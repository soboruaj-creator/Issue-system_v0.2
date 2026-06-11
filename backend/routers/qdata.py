"""Q-data 관리 라우터"""
from fastapi import APIRouter, HTTPException, Query
from typing import Optional
from datetime import datetime, timedelta
from database import get_collection

router = APIRouter(prefix="/api/qdata", tags=["qdata"])


def _serialize(doc: dict) -> dict:
    if doc and "_id" in doc:
        doc["_id"] = str(doc["_id"])
    return doc


@router.get("")
async def list_qdata(
    page: int = Query(1, ge=1),
    size: int = Query(20, ge=1, le=100),
    model_name: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
):
    col = get_collection("q_data")
    match = {}
    if model_name:
        match["model_name"] = model_name
    if start_date and end_date:
        match["service_date"] = {"$gte": start_date, "$lte": end_date}

    total = await col.count_documents(match)
    skip = (page - 1) * size
    docs = await col.find(match).sort("service_date", -1).skip(skip).limit(size).to_list(None)

    return {
        "total": total,
        "page": page,
        "size": size,
        "items": [_serialize(d) for d in docs],
    }


@router.get("/check-duplicates")
async def check_qdata_duplicates():
    col = get_collection("q_data")
    pipeline = [
        {"$match": {"serial_number": {"$ne": None}, "service_date": {"$ne": None}}},
        {
            "$group": {
                "_id": {"serial_number": "$serial_number", "service_date": "$service_date"},
                "count": {"$sum": 1},
                "ids": {"$push": {"$toString": "$_id"}},
            }
        },
        {"$match": {"count": {"$gt": 1}}},
        {"$sort": {"count": -1}},
        {
            "$project": {
                "serial_number": "$_id.serial_number",
                "service_date": "$_id.service_date",
                "count": 1,
                "_id": 0,
            }
        },
    ]
    duplicates = await col.aggregate(pipeline).to_list(None)
    return {"duplicate_count": len(duplicates), "duplicates": duplicates}


@router.post("/remove-duplicates")
async def remove_qdata_duplicates():
    col = get_collection("q_data")
    pipeline = [
        {"$match": {"serial_number": {"$ne": None}, "service_date": {"$ne": None}}},
        {
            "$group": {
                "_id": {"serial_number": "$serial_number", "service_date": "$service_date"},
                "ids": {"$push": "$_id"},
                "count": {"$sum": 1},
            }
        },
        {"$match": {"count": {"$gt": 1}}},
    ]
    groups = await col.aggregate(pipeline).to_list(None)
    removed = 0
    for group in groups:
        ids_to_remove = group["ids"][1:]  # 첫 번째 제외하고 삭제
        from bson import ObjectId
        await col.delete_many({"_id": {"$in": ids_to_remove}})
        removed += len(ids_to_remove)

    return {"success": True, "removed_count": removed}


@router.post("/reset")
async def reset_qdata():
    col = get_collection("q_data")
    result = await col.delete_many({})
    return {"success": True, "deleted_count": result.deleted_count}


@router.post("/reset-all")
async def reset_voc_data():
    voc_col = get_collection("internal_voc")
    result = await voc_col.delete_many({})
    return {"success": True, "deleted_count": result.deleted_count, "message": "VOC 데이터가 초기화되었습니다."}


@router.get("/debug")
async def debug_qdata():
    """Q-data 컬렉션 현황 진단 엔드포인트"""
    col = get_collection("q_data")
    total = await col.count_documents({})

    # 최근 30일 날짜 범위
    today = datetime.now().date()
    from_date = (today - timedelta(days=30)).strftime("%Y-%m-%d")
    recent_count = await col.count_documents({"service_date": {"$gte": from_date}})

    # 날짜별 건수 (최근 30일)
    daily_pipeline = [
        {"$match": {"service_date": {"$gte": from_date}}},
        {"$group": {"_id": "$service_date", "count": {"$sum": 1}}},
        {"$sort": {"_id": -1}},
    ]
    daily = await col.aggregate(daily_pipeline).to_list(None)

    # 모델별 건수 (최근 30일)
    model_pipeline = [
        {"$match": {"service_date": {"$gte": from_date}}},
        {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}},
        {"$limit": 10},
    ]
    by_model = await col.aggregate(model_pipeline).to_list(None)

    # 가장 최근 업로드된 5건 샘플
    sample = await col.find(
        {}, {"_id": 0, "service_date": 1, "model_name": 1, "serial_number": 1, "uploaded_date": 1}
    ).sort("uploaded_date", -1).limit(5).to_list(None)

    return {
        "total": total,
        "recent_30days": recent_count,
        "daily_counts": [{"date": d["_id"], "count": d["count"]} for d in daily],
        "by_model_recent": [{"model": m["_id"], "count": m["count"]} for m in by_model],
        "latest_5_sample": sample,
    }
