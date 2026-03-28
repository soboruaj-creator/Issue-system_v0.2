"""업로드 라우터"""
from fastapi import APIRouter, UploadFile, File, HTTPException
from datetime import datetime
from database import get_collection
from services.excel_service import read_excel_with_drm, read_qdata_excel_with_drm
from services.voc_parser import process_voc_row, merge_similar_chipsets, convert_qdata_date
import pandas as pd

router = APIRouter(prefix="/api/upload", tags=["upload"])


@router.post("/internal_voc")
async def upload_internal_voc(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")

    try:
        df = await read_excel_with_drm(file)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    voc_col = get_collection("internal_voc")
    chipset_col = get_collection("chipset_mapping")
    app_col = get_collection("app_keywords")

    # 칩셋 맵, 앱 키워드 캐시
    chipset_docs = await chipset_col.find({}, {"model_name": 1, "chipset": 1}).to_list(None)
    chipset_map = {d["model_name"]: d["chipset"] for d in chipset_docs}

    app_docs = await app_col.find({}).to_list(None)

    success_count = 0
    error_count = 0
    unmapped_models = set()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for _, row in df.iterrows():
        case_code, voc_data, is_unmapped = process_voc_row(
            row, file.filename or "", chipset_map, app_docs
        )
        if case_code is None:
            continue
        if is_unmapped and voc_data.get("model_name"):
            unmapped_models.add(voc_data["model_name"])

        try:
            existing = await voc_col.find_one({"case_code": case_code})
            if existing:
                await voc_col.update_one(
                    {"case_code": case_code},
                    {"$set": {
                        "model_name": voc_data["model_name"],
                        "cause": voc_data["cause"],
                        "solution": voc_data["solution"],
                        "uploaded_date": now_str,
                    }},
                )
            else:
                voc_data["case_code"] = case_code
                voc_data["uploaded_date"] = now_str
                await voc_col.insert_one(voc_data)
            success_count += 1
        except Exception as e:
            error_count += 1
            print(f"Insert error: {e}")

    message = f"업로드 완료: {success_count}건 성공, {error_count}건 실패"
    if unmapped_models:
        message += f"\n칩셋 미매핑 모델: {len(unmapped_models)}개"

    return {
        "success": True,
        "message": message,
        "unmapped_models": list(unmapped_models),
    }


@router.post("/chipset_mapping")
async def upload_chipset_mapping(file: UploadFile = File(...)):
    try:
        df = await read_excel_with_drm(file)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    col = get_collection("chipset_mapping")
    voc_col = get_collection("internal_voc")

    pairs = []
    for _, row in df.iterrows():
        model_name = str(row.iloc[0]).strip()
        chipset = str(row.iloc[1]).strip()
        if model_name and chipset and model_name != "nan" and chipset != "nan":
            pairs.append((model_name, chipset))

    all_chipsets = [c for _, c in pairs]
    merged = merge_similar_chipsets(all_chipsets)

    success_count = update_count = 0
    for model_name, chipset in pairs:
        final_chipset = merged.get(chipset, chipset)
        existing = await col.find_one({"model_name": model_name})
        if existing:
            await col.update_one({"model_name": model_name}, {"$set": {"chipset": final_chipset}})
            update_count += 1
        else:
            await col.insert_one({"model_name": model_name, "chipset": final_chipset})
            success_count += 1
        # VOC 데이터의 칩셋도 업데이트
        await voc_col.update_many({"model_name": model_name}, {"$set": {"chipset": final_chipset}})

    msg = f"{success_count}개 등록"
    if update_count:
        msg += f", {update_count}개 업데이트"
    return {"success": True, "message": msg}


@router.post("/app_keywords")
async def upload_app_keywords(file: UploadFile = File(...)):
    try:
        df = await read_excel_with_drm(file)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    col = get_collection("app_keywords")
    await col.delete_many({})

    docs = []
    for _, row in df.iterrows():
        app_name = str(row.iloc[0]).strip()
        keywords = str(row.iloc[1]).strip()
        if app_name and keywords and app_name != "nan":
            docs.append({"app_name": app_name, "keywords": keywords})

    if docs:
        await col.insert_many(docs)

    return {"success": True, "message": f"{len(docs)}개의 앱 키워드가 등록되었습니다."}


@router.post("/qdata")
async def upload_qdata(file: UploadFile = File(...)):
    try:
        df = await read_qdata_excel_with_drm(file)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    # 날짜 변환
    df["service_date"] = df["service_date"].apply(convert_qdata_date)
    df = df.dropna(subset=["service_date", "model_name"])

    col = get_collection("q_data")
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    success_count = duplicate_count = error_count = 0

    for _, row in df.iterrows():
        doc = row.where(pd.notna(row), None).to_dict()
        doc["uploaded_date"] = now_str

        serial = doc.get("serial_number")
        service_date = doc.get("service_date")

        try:
            if serial and service_date:
                existing = await col.find_one({"serial_number": serial, "service_date": service_date})
                if existing:
                    duplicate_count += 1
                    continue
            await col.insert_one(doc)
            success_count += 1
        except Exception as e:
            error_count += 1
            print(f"Q-data insert error: {e}")

    return {
        "success": True,
        "message": f"Q-data 업로드 완료: {success_count}건 성공, {duplicate_count}건 중복, {error_count}건 실패",
        "success_count": success_count,
        "duplicate_count": duplicate_count,
        "error_count": error_count,
    }
