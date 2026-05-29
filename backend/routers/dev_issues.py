"""개발모델 이슈 라우터"""
import os
import shutil
from datetime import datetime
from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
from database import get_collection

router = APIRouter(prefix="/api/dev-issues", tags=["dev-issues"])

ATTACH_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "uploads", "dev_attachments")
os.makedirs(ATTACH_DIR, exist_ok=True)

EXCLUDE_KEYWORDS = ["members", "k zone", "rdm", "디버그분석"]


def _should_exclude(title: str) -> bool:
    t = (title or "").lower()
    return any(kw in t for kw in EXCLUDE_KEYWORDS)


@router.get("/dashboard")
async def get_dev_dashboard():
    col = get_collection("dev_issues")
    docs = await col.find({}, {"_id": 0, "model_name": 1, "status": 1, "is_pending": 1}).to_list(None)

    models: dict = {}
    for d in docs:
        m = d.get("model_name") or "미분류"
        if m not in models:
            models[m] = {"model_name": m, "open": 0, "resolve": 0, "close": 0, "pending": 0, "total": 0}
        s = (d.get("status") or "").lower()
        if s in ("open", "resolve", "close"):
            models[m][s] += 1
        models[m]["total"] += 1
        if d.get("is_pending"):
            models[m]["pending"] += 1

    result = [v for v in models.values() if v["open"] > 0 or v["resolve"] > 0 or v["pending"] > 0]
    return sorted(result, key=lambda x: x["open"] + x["resolve"] + x["pending"], reverse=True)


@router.get("/model/{model_name}")
async def get_model_dev_issues(model_name: str):
    from routers.statistics import _get_marketing_map
    mmap = await _get_marketing_map()

    matching = {model_name}
    for raw, mkt in mmap.items():
        if mkt == model_name:
            matching.add(raw)
        if raw == model_name and mkt:
            matching.add(mkt)

    col = get_collection("dev_issues")
    all_docs = await col.find(
        {"model_name": {"$in": list(matching)}},
        {"_id": 0, "case_code": 1, "title": 1, "status": 1, "issue_type": 1,
         "is_pending": 1, "pending_memo": 1, "uploaded_date": 1},
    ).to_list(None)

    stats = {"total": len(all_docs), "open": 0, "resolve": 0, "close": 0, "pending": 0}
    for d in all_docs:
        s = (d.get("status") or "").lower()
        if s in stats:
            stats[s] += 1
        if d.get("is_pending"):
            stats["pending"] += 1

    active = [d for d in all_docs if (d.get("status") or "").lower() != "close"]
    active.sort(key=lambda x: (-(1 if x.get("is_pending") else 0), x.get("status", "")))
    return {"stats": stats, "issues": active}


@router.get("/search")
async def search_dev_issue(case_code: str):
    col = get_collection("dev_issues")
    doc = await col.find_one({"case_code": case_code}, {"_id": 0})
    if not doc:
        raise HTTPException(status_code=404, detail="이슈를 찾을 수 없습니다.")
    return doc


@router.post("/{case_code}/pending")
async def update_pending(case_code: str, body: dict):
    col = get_collection("dev_issues")
    if not await col.find_one({"case_code": case_code}):
        raise HTTPException(status_code=404, detail="이슈를 찾을 수 없습니다.")
    await col.update_one(
        {"case_code": case_code},
        {"$set": {"is_pending": body.get("is_pending", False), "pending_memo": body.get("memo", "")}},
    )
    return await col.find_one({"case_code": case_code}, {"_id": 0})


@router.post("/{case_code}/attachment")
async def upload_attachment(case_code: str, file: UploadFile = File(...)):
    col = get_collection("dev_issues")
    if not await col.find_one({"case_code": case_code}):
        raise HTTPException(status_code=404, detail="이슈를 찾을 수 없습니다.")

    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    stored_name = f"{case_code}_{now}_{file.filename}"
    with open(os.path.join(ATTACH_DIR, stored_name), "wb") as f:
        shutil.copyfileobj(file.file, f)

    attachment = {
        "filename": file.filename,
        "stored_name": stored_name,
        "uploaded_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    await col.update_one({"case_code": case_code}, {"$push": {"pending_attachments": attachment}})
    return attachment


@router.delete("/{case_code}/attachment/{stored_name}")
async def delete_attachment(case_code: str, stored_name: str):
    path = os.path.join(ATTACH_DIR, stored_name)
    if os.path.exists(path):
        os.remove(path)
    col = get_collection("dev_issues")
    await col.update_one({"case_code": case_code}, {"$pull": {"pending_attachments": {"stored_name": stored_name}}})
    return {"success": True}


@router.get("/{case_code}/attachment/{stored_name}")
async def download_attachment(case_code: str, stored_name: str):
    col = get_collection("dev_issues")
    doc = await col.find_one(
        {"case_code": case_code, "pending_attachments.stored_name": stored_name},
        {"pending_attachments": 1},
    )
    original_name = stored_name
    if doc:
        for a in doc.get("pending_attachments", []):
            if a["stored_name"] == stored_name:
                original_name = a["filename"]
                break
    path = os.path.join(ATTACH_DIR, stored_name)
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")
    return FileResponse(path, filename=original_name)
