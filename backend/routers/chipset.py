"""칩셋 관리 라우터"""
from fastapi import APIRouter, HTTPException
from database import get_collection
from services.voc_parser import merge_similar_chipsets

router = APIRouter(prefix="/api", tags=["chipset"])


@router.get("/unmapped-models")
async def get_unmapped_models():
    voc_col = get_collection("internal_voc")
    chipset_col = get_collection("chipset_mapping")

    # 매핑된 모델 목록
    mapped_docs = await chipset_col.find({}, {"model_name": 1}).to_list(None)
    mapped_models = {d["model_name"] for d in mapped_docs}

    # 칩셋 없는 VOC 모델
    pipeline = [
        {"$match": {"model_name": {"$ne": None}, "chipset": None}},
        {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}},
        {"$project": {"model_name": "$_id", "count": 1, "_id": 0}},
    ]
    unmapped = await voc_col.aggregate(pipeline).to_list(None)
    return [u for u in unmapped if u["model_name"] not in mapped_models]


@router.post("/chipset-mapping/add")
async def add_chipset_mapping(body: dict):
    model_name = body.get("model_name", "").strip()
    chipset = body.get("chipset", "").strip()
    if not model_name or not chipset:
        raise HTTPException(status_code=400, detail="model_name과 chipset이 필요합니다.")

    col = get_collection("chipset_mapping")
    voc_col = get_collection("internal_voc")

    existing = await col.find_one({"model_name": model_name})
    if existing:
        await col.update_one({"model_name": model_name}, {"$set": {"chipset": chipset}})
    else:
        await col.insert_one({"model_name": model_name, "chipset": chipset})

    await voc_col.update_many({"model_name": model_name}, {"$set": {"chipset": chipset}})
    return {"success": True, "message": f"{model_name} → {chipset} 매핑 완료"}


@router.post("/chipset-mapping/batch")
async def add_chipset_mapping_batch(body: dict):
    mappings = body.get("mappings", [])
    if not mappings:
        raise HTTPException(status_code=400, detail="mappings 배열이 필요합니다.")

    col = get_collection("chipset_mapping")
    voc_col = get_collection("internal_voc")

    all_chipsets = [m.get("chipset", "") for m in mappings]
    merged = merge_similar_chipsets(all_chipsets)

    success = 0
    for m in mappings:
        model_name = m.get("model_name", "").strip()
        chipset = merged.get(m.get("chipset", ""), m.get("chipset", ""))
        if not model_name or not chipset:
            continue
        existing = await col.find_one({"model_name": model_name})
        if existing:
            await col.update_one({"model_name": model_name}, {"$set": {"chipset": chipset}})
        else:
            await col.insert_one({"model_name": model_name, "chipset": chipset})
        await voc_col.update_many({"model_name": model_name}, {"$set": {"chipset": chipset}})
        success += 1

    return {"success": True, "message": f"{success}개 매핑 처리 완료"}


@router.post("/chipset/merge")
async def merge_chipsets_endpoint(body: dict):
    source = body.get("source_chipset", "").strip()
    target = body.get("target_chipset", "").strip()
    if not source or not target:
        raise HTTPException(status_code=400, detail="source_chipset과 target_chipset이 필요합니다.")

    voc_col = get_collection("internal_voc")
    chipset_col = get_collection("chipset_mapping")

    result = await voc_col.update_many({"chipset": source}, {"$set": {"chipset": target}})
    await chipset_col.update_many({"chipset": source}, {"$set": {"chipset": target}})

    return {"success": True, "updated_count": result.modified_count, "message": f"{source} → {target} 병합 완료"}


@router.post("/chipset/rename")
async def rename_chipset(body: dict):
    old_name = body.get("old_name", "").strip()
    new_name = body.get("new_name", "").strip()
    if not old_name or not new_name:
        raise HTTPException(status_code=400, detail="old_name과 new_name이 필요합니다.")

    voc_col = get_collection("internal_voc")
    chipset_col = get_collection("chipset_mapping")

    await voc_col.update_many({"chipset": old_name}, {"$set": {"chipset": new_name}})
    await chipset_col.update_many({"chipset": old_name}, {"$set": {"chipset": new_name}})

    return {"success": True, "message": f"칩셋명 변경: {old_name} → {new_name}"}


@router.post("/model/update-watch")
async def update_watch_models():
    from services.voc_parser import extract_watch_model, map_model_name
    voc_col = get_collection("internal_voc")

    docs = await voc_col.find({"title": {"$ne": None}}, {"_id": 1, "title": 1, "model_name": 1}).to_list(None)
    updated = 0
    for doc in docs:
        watch = extract_watch_model(doc.get("title"))
        if watch:
            mapped = map_model_name(watch)
            if mapped != doc.get("model_name"):
                await voc_col.update_one({"_id": doc["_id"]}, {"$set": {"model_name": mapped}})
                updated += 1

    return {"success": True, "updated_count": updated}


@router.post("/model/update-mapping")
async def update_model_mapping():
    voc_col = get_collection("internal_voc")
    chipset_col = get_collection("chipset_mapping")

    chipset_docs = await chipset_col.find({}).to_list(None)
    chipset_map = {d["model_name"]: d["chipset"] for d in chipset_docs}

    updated = 0
    for model_name, chipset in chipset_map.items():
        result = await voc_col.update_many(
            {"model_name": model_name, "chipset": None},
            {"$set": {"chipset": chipset}},
        )
        updated += result.modified_count

    return {"success": True, "updated_count": updated}


@router.get("/chipset-mappings")
async def list_chipset_mappings():
    col = get_collection("chipset_mapping")
    docs = await col.find({}, {"_id": 0}).sort("model_name", 1).to_list(None)
    return docs
