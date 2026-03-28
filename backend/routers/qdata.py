"""Q-data 관리 라우터"""
from fastapi import APIRouter, HTTPException, Query
from typing import Optional
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
