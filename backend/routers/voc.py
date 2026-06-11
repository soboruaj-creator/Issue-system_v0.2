"""VOC CRUD 라우터"""
from fastapi import APIRouter, HTTPException, Query
from typing import Optional
from datetime import datetime
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


@router.get("/resolved-with-fix")
async def get_resolved_with_fix():
    from routers.statistics import _get_marketing_map
    col = get_collection("internal_voc")
    docs = await col.find(
        {"resolve_option": {"$regex": "수정완료", "$options": "i"},
         "status": "resolve"},
        {"_id": 0, "case_code": 1, "model_name": 1, "title": 1, "created_date": 1},
    ).sort("created_date", -1).to_list(None)
    mmap = await _get_marketing_map()
    for d in docs:
        raw = d.get("model_name") or "미분류"
        d["model_name"] = mmap.get(raw) or raw
    return docs


@router.get("/search")
async def search_voc(case_code: str):
    col = get_collection("internal_voc")
    doc = await col.find_one({"case_code": case_code}, {"_id": 0})
    if not doc:
        raise HTTPException(status_code=404, detail="이슈를 찾을 수 없습니다.")
    return doc


@router.post("/{case_code}/pending")
async def update_voc_pending(case_code: str, body: dict):
    col = get_collection("internal_voc")
    if not await col.find_one({"case_code": case_code}):
        raise HTTPException(status_code=404, detail="이슈를 찾을 수 없습니다.")
    await col.update_one(
        {"case_code": case_code},
        {"$set": {"is_pending": body.get("is_pending", False), "pending_memo": body.get("memo", "")}},
    )
    return await col.find_one({"case_code": case_code}, {"_id": 0})


@router.get("/model-active/{model_name}")
async def get_voc_model_active(model_name: str):
    from routers.statistics import _get_marketing_map, _extract_summary
    mmap = await _get_marketing_map()
    matching = {model_name}
    for raw, mkt in mmap.items():
        if mkt == model_name:
            matching.add(raw)
        if raw == model_name and mkt:
            matching.add(mkt)

    names = [n for n in matching if n is not None and n != "미분류"]
    include_null = None in matching or "미분류" in matching
    if include_null and names:
        query = {"$or": [{"model_name": {"$in": names}}, {"model_name": None}, {"model_name": ""}]}
    elif include_null:
        query = {"$or": [{"model_name": None}, {"model_name": ""}]}
    else:
        query = {"model_name": {"$in": names}}

    col = get_collection("internal_voc")
    docs = await col.find(
        query,
        {"_id": 0, "case_code": 1, "title": 1, "status": 1, "is_pending": 1, "pending_memo": 1,
         "created_date": 1, "problem": 1, "original_content": 1, "reproduction_path": 1},
    ).to_list(None)

    total = len(docs)
    open_count = sum(1 for d in docs if (d.get("status") or "open") == "open")
    resolve_count = sum(1 for d in docs if (d.get("status") or "open") == "resolve")
    close_count = sum(1 for d in docs if (d.get("status") or "open") == "close" and not d.get("is_pending"))
    pending_count = sum(1 for d in docs if d.get("is_pending"))

    active = [d for d in docs if (d.get("status") or "open") != "close" or d.get("is_pending")]
    active.sort(key=lambda x: (-(1 if x.get("is_pending") else 0), x.get("status") or "open"))

    for d in active:
        d["summary"] = _extract_summary(d.get("problem"), d.get("original_content"), d.get("reproduction_path"))
        d.pop("problem", None)
        d.pop("original_content", None)
        d.pop("reproduction_path", None)

    return {
        "stats": {"total": total, "open": open_count, "resolve": resolve_count, "pending": pending_count, "close": close_count},
        "issues": active,
    }


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
