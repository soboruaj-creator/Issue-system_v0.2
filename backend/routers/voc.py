"""VOC CRUD 라우터"""
from fastapi import APIRouter, HTTPException, Query
from typing import Optional, List
from datetime import datetime
from bson import ObjectId
from database import get_collection

router = APIRouter(prefix="/api/voc", tags=["voc"])


def _serialize(doc: dict) -> dict:
    if doc and "_id" in doc:
        doc["_id"] = str(doc["_id"])
    return doc


@router.get("")
async def list_voc(
    page: int = Query(1, ge=1),
    size: int = Query(20, ge=1, le=100),
    model_name: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    search: Optional[str] = None,
):
    col = get_collection("internal_voc")
    query = {}
    if model_name:
        query["model_name"] = model_name
    if start_date and end_date:
        query["created_date"] = {"$gte": start_date, "$lte": end_date}
    if search:
        query["$or"] = [
            {"case_code": {"$regex": search, "$options": "i"}},
            {"title": {"$regex": search, "$options": "i"}},
            {"problem": {"$regex": search, "$options": "i"}},
        ]

    total = await col.count_documents(query)
    skip = (page - 1) * size
    docs = await col.find(query).sort("created_date", -1).skip(skip).limit(size).to_list(None)

    return {
        "total": total,
        "page": page,
        "size": size,
        "items": [_serialize(d) for d in docs],
    }


@router.get("/{case_code}")
async def get_voc_detail(case_code: str):
    col = get_collection("internal_voc")
    comment_col = get_collection("comments")

    doc = await col.find_one({"case_code": case_code})
    if not doc:
        raise HTTPException(status_code=404, detail="VOC를 찾을 수 없습니다.")

    comments = await comment_col.find({"voc_case_code": case_code}).sort("created_date", 1).to_list(None)

    result = _serialize(doc)
    result["comments"] = [_serialize(c) for c in comments]
    return result


@router.post("/{case_code}/comment")
async def add_comment(case_code: str, body: dict):
    comment_text = body.get("comment", "").strip()
    if not comment_text:
        raise HTTPException(status_code=400, detail="댓글 내용이 없습니다.")

    col = get_collection("comments")
    doc = {
        "voc_case_code": case_code,
        "voc_type": body.get("voc_type", "internal"),
        "comment": comment_text,
        "created_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    result = await col.insert_one(doc)
    doc["_id"] = str(result.inserted_id)
    return doc


@router.get("/by-model/{model_name}")
async def get_voc_by_model(
    model_name: str,
    month: Optional[str] = None,
    page: int = Query(1, ge=1),
    size: int = Query(50, ge=1, le=200),
):
    col = get_collection("internal_voc")
    query = {"model_name": model_name}
    if month:
        query["created_date"] = {"$regex": f"^{month}"}

    total = await col.count_documents(query)
    skip = (page - 1) * size
    docs = await col.find(query).sort("created_date", -1).skip(skip).limit(size).to_list(None)

    return {
        "total": total,
        "model_name": model_name,
        "items": [_serialize(d) for d in docs],
    }
