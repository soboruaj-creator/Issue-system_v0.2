"""모델 정보(이벤트 로그) CRUD 라우터"""
from fastapi import APIRouter, HTTPException
from bson import ObjectId
from database import get_collection

router = APIRouter(prefix="/api/model-notes", tags=["model-notes"])


def _serialize(doc: dict) -> dict:
    doc["id"] = str(doc.pop("_id"))
    return doc


@router.get("/{model_name}")
async def get_notes(model_name: str):
    """모델의 이벤트 로그 목록 (날짜 내림차순)"""
    col = get_collection("model_notes")
    docs = await col.find(
        {"model_name": model_name},
        {"model_name": 0},
    ).sort("date", -1).to_list(None)
    return [_serialize(d) for d in docs]


@router.post("/{model_name}")
async def create_note(model_name: str, body: dict):
    """이벤트 로그 추가"""
    date = (body.get("date") or "").strip()
    content = (body.get("content") or "").strip()
    if not date or not content:
        raise HTTPException(status_code=400, detail="date와 content는 필수입니다")
    col = get_collection("model_notes")
    result = await col.insert_one({"model_name": model_name, "date": date, "content": content})
    doc = await col.find_one({"_id": result.inserted_id}, {"model_name": 0})
    return _serialize(doc)


@router.put("/{note_id}")
async def update_note(note_id: str, body: dict):
    """이벤트 로그 수정"""
    date = (body.get("date") or "").strip()
    content = (body.get("content") or "").strip()
    if not date or not content:
        raise HTTPException(status_code=400, detail="date와 content는 필수입니다")
    try:
        oid = ObjectId(note_id)
    except Exception:
        raise HTTPException(status_code=400, detail="잘못된 note_id")
    col = get_collection("model_notes")
    await col.update_one({"_id": oid}, {"$set": {"date": date, "content": content}})
    doc = await col.find_one({"_id": oid}, {"model_name": 0})
    if not doc:
        raise HTTPException(status_code=404, detail="노트를 찾을 수 없습니다")
    return _serialize(doc)


@router.delete("/{note_id}")
async def delete_note(note_id: str):
    """이벤트 로그 삭제"""
    try:
        oid = ObjectId(note_id)
    except Exception:
        raise HTTPException(status_code=400, detail="잘못된 note_id")
    col = get_collection("model_notes")
    result = await col.delete_one({"_id": oid})
    if result.deleted_count == 0:
        raise HTTPException(status_code=404, detail="노트를 찾을 수 없습니다")
    return {"ok": True}
