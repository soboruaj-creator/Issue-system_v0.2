"""업로드 라우터"""
import re
from fastapi import APIRouter, UploadFile, File, HTTPException
from datetime import datetime
from database import get_collection
from services.excel_service import read_excel_with_drm, read_qdata_excel_with_drm
from services.voc_parser import (process_voc_row, merge_similar_chipsets, convert_qdata_date,
                                 extract_watch_model, extract_model_from_title, map_model_name)
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
                # close 상태는 open/resolve로 되돌릴 수 없음
                new_status = voc_data.get("status", "open")
                if existing.get("status") == "close":
                    new_status = "close"
                await voc_col.update_one(
                    {"case_code": case_code},
                    {"$set": {
                        "model_name": voc_data["model_name"],
                        "status": new_status,
                        "cause": voc_data["cause"],
                        "solution": voc_data["solution"],
                        "resolve_option": voc_data.get("resolve_option"),
                        "uploaded_date": now_str,
                    }},
                )
            else:
                voc_data["case_code"] = case_code
                voc_data["uploaded_date"] = now_str
                voc_data.setdefault("is_pending", False)
                voc_data.setdefault("pending_memo", "")
                voc_data.setdefault("pending_attachments", [])
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

    # J열 필터: 'GPS 수신 이상' 항목만 처리
    unique_vals = df["j_category"].astype(str).unique().tolist()
    print(f"[Q-data] J열 고유값 ({len(unique_vals)}개): {unique_vals[:20]}")
    before_count = len(df)
    df = df[df["j_category"].astype(str).str.contains("GPS 수신 이상", na=False)]
    print(f"[Q-data] 필터 전: {before_count}행 → 필터 후: {len(df)}행")

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


@router.post("/dev_issues")
async def upload_dev_issues(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")
    try:
        df = await read_excel_with_drm(file, header=2)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    EXCLUDE = ["members", "k zone", "rdm", "디버그분석"]
    col = get_collection("dev_issues")
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    success_count = skip_count = error_count = 0

    for _, row in df.iterrows():
        try:
            def safe(val):
                return str(val) if pd.notna(val) else None

            case_code = safe(row.iloc[0])
            title = safe(row.iloc[7])
            status_raw = safe(row.iloc[6])

            if not case_code:
                continue

            title_lower = (title or "").lower()
            if any(kw in title_lower for kw in EXCLUDE):
                skip_count += 1
                continue

            issue_type = "UT" if title and "ut" in title_lower else "자체"

            col_d = safe(row.iloc[3])  # D열: 모델명 직접 기재

            watch = extract_watch_model(title)
            if watch:
                model_name = watch
            else:
                model_name = extract_model_from_title(title)
                if not model_name and col_d:
                    model_name = extract_watch_model(col_d) or extract_model_from_title(col_d) or col_d.strip()
            model_name = map_model_name(model_name) if model_name else None

            status_norm = (status_raw or "").strip().lower()
            if status_norm.startswith("resolve"):
                status_norm = "resolve"
            elif status_norm not in ("open", "close"):
                status_norm = "open"

            created_date = None
            if case_code and case_code.upper().startswith("P"):
                dm = re.search(r"P(\d{6})", case_code, re.IGNORECASE)
                if dm:
                    try:
                        created_date = datetime.strptime(dm.group(1), "%y%m%d").strftime("%Y-%m-%d")
                    except Exception:
                        pass

            existing = await col.find_one({"case_code": case_code})
            if existing:
                # close 상태는 open/resolve로 되돌릴 수 없음
                final_status = status_norm if existing.get("status") != "close" else "close"
                update_fields = {
                    "title": title,
                    "model_name": model_name,
                    "status": final_status,
                    "issue_type": issue_type,
                    "uploaded_date": now_str,
                }
                if created_date and not existing.get("created_date"):
                    update_fields["created_date"] = created_date
                await col.update_one({"case_code": case_code}, {"$set": update_fields})
            else:
                await col.insert_one({
                    "case_code": case_code,
                    "title": title,
                    "model_name": model_name,
                    "status": status_norm,
                    "issue_type": issue_type,
                    "created_date": created_date,
                    "is_pending": False,
                    "pending_memo": "",
                    "pending_attachments": [],
                    "uploaded_date": now_str,
                })
            success_count += 1
        except Exception as e:
            error_count += 1
            print(f"Dev issue row error: {e}")

    return {
        "success": True,
        "message": f"개발이슈 업로드 완료: {success_count}건 처리, {skip_count}건 제외, {error_count}건 오류",
    }


@router.post("/launch_dates")
async def upload_launch_dates(file: UploadFile = File(...)):
    """출시일 업로드 (A열: 모델명, B열: 출시일)"""
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="엑셀 파일만 업로드 가능합니다.")

    try:
        df = await read_excel_with_drm(file)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    col = get_collection("launch_dates")
    success_count = 0

    for _, row in df.iterrows():
        model_name = str(row.iloc[0]).strip()
        if not model_name or model_name == "nan":
            continue

        launch_date_raw = row.iloc[1] if len(row) > 1 else None

        # 날짜 파싱 (datetime 객체 또는 문자열, 공란이면 None으로 처리)
        launch_date_str = None
        if pd.notna(launch_date_raw) if launch_date_raw is not None else False:
            if hasattr(launch_date_raw, "strftime"):
                launch_date_str = launch_date_raw.strftime("%Y-%m-%d")
            elif isinstance(launch_date_raw, str) and launch_date_raw.strip() not in ("", "nan"):
                for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
                    try:
                        launch_date_str = datetime.strptime(launch_date_raw.strip(), fmt).strftime("%Y-%m-%d")
                        break
                    except ValueError:
                        continue

        # C열 마케팅명 (선택)
        marketing_name = None
        if len(row) > 2:
            raw = row.iloc[2]
            if pd.notna(raw) and str(raw).strip() not in ("", "nan"):
                marketing_name = str(raw).strip()

        # 출시일도 마케팅명도 없으면 건너뜀
        if not launch_date_str and not marketing_name:
            continue

        update_fields = {"model_name": model_name}
        if launch_date_str:
            update_fields["launch_date"] = launch_date_str
        if marketing_name is not None:
            update_fields["marketing_name"] = marketing_name

        await col.update_one(
            {"model_name": model_name},
            {"$set": update_fields},
            upsert=True,
        )
        success_count += 1

    return {"success": True, "message": f"{success_count}개 모델 출시일 등록 완료"}
