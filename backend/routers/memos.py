"""메모 관리 라우터"""
from fastapi import APIRouter, HTTPException
from datetime import datetime
from database import get_collection

router = APIRouter(prefix="/api", tags=["memos"])


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ─── 월별 메모 ───────────────────────────────────────────────────────────────

@router.get("/monthly-memos")
async def get_monthly_memos():
    col = get_collection("monthly_memos")
    docs = await col.find({}, {"_id": 0}).sort("month", 1).to_list(None)
    return docs


@router.post("/monthly-memos")
async def add_monthly_memo(body: dict):
    month = body.get("month", "").strip()
    memo = body.get("memo", "").strip()
    if not month or not memo:
        raise HTTPException(status_code=400, detail="month와 memo 필드가 필요합니다.")

    col = get_collection("monthly_memos")
    existing = await col.find_one({"month": month})
    now = _now()
    if existing:
        await col.update_one({"month": month}, {"$set": {"memo": memo, "updated_date": now}})
    else:
        await col.insert_one({"month": month, "memo": memo, "created_date": now, "updated_date": now})
    return {"success": True, "month": month}


@router.put("/monthly-memos/{month}")
async def update_monthly_memo(month: str, body: dict):
    memo = body.get("memo", "").strip()
    if not memo:
        raise HTTPException(status_code=400, detail="memo 필드가 필요합니다.")
    col = get_collection("monthly_memos")
    await col.update_one({"month": month}, {"$set": {"memo": memo, "updated_date": _now()}})
    return {"success": True}


@router.delete("/monthly-memos/{month}")
async def delete_monthly_memo(month: str):
    col = get_collection("monthly_memos")
    await col.delete_one({"month": month})
    return {"success": True}


# ─── 주별 메모 ────────────────────────────────────────────────────────────────

@router.get("/weekly-memos")
async def get_weekly_memos():
    col = get_collection("weekly_memos")
    docs = await col.find({}, {"_id": 0}).sort("week", 1).to_list(None)
    return docs


@router.post("/weekly-memos")
async def add_weekly_memo(body: dict):
    week = body.get("week", "").strip()
    memo = body.get("memo", "").strip()
    if not week or not memo:
        raise HTTPException(status_code=400, detail="week와 memo 필드가 필요합니다.")

    col = get_collection("weekly_memos")
    existing = await col.find_one({"week": week})
    now = _now()
    if existing:
        await col.update_one({"week": week}, {"$set": {"memo": memo, "updated_date": now}})
    else:
        await col.insert_one({"week": week, "memo": memo, "created_date": now, "updated_date": now})
    return {"success": True, "week": week}


@router.put("/weekly-memos/{week}")
async def update_weekly_memo(week: str, body: dict):
    memo = body.get("memo", "").strip()
    if not memo:
        raise HTTPException(status_code=400, detail="memo 필드가 필요합니다.")
    col = get_collection("weekly_memos")
    await col.update_one({"week": week}, {"$set": {"memo": memo, "updated_date": _now()}})
    return {"success": True}


@router.delete("/weekly-memos/{week}")
async def delete_weekly_memo(week: str):
    col = get_collection("weekly_memos")
    await col.delete_one({"week": week})
    return {"success": True}


# ─── 모델별 월별 메모 ──────────────────────────────────────────────────────────

@router.get("/model-monthly-memos/{model_name}")
async def get_model_monthly_memos(model_name: str):
    col = get_collection("model_monthly_memos")
    docs = await col.find({"model_name": model_name}, {"_id": 0}).sort("month", 1).to_list(None)
    return docs


@router.post("/model-monthly-memos")
async def add_model_monthly_memo(body: dict):
    model_name = body.get("model_name", "").strip()
    month = body.get("month", "").strip()
    memo = body.get("memo", "").strip()
    if not model_name or not month or not memo:
        raise HTTPException(status_code=400, detail="model_name, month, memo 필드가 필요합니다.")

    col = get_collection("model_monthly_memos")
    existing = await col.find_one({"model_name": model_name, "month": month})
    now = _now()
    if existing:
        await col.update_one(
            {"model_name": model_name, "month": month},
            {"$set": {"memo": memo, "updated_date": now}},
        )
    else:
        await col.insert_one({
            "model_name": model_name, "month": month,
            "memo": memo, "created_date": now, "updated_date": now,
        })
    return {"success": True}


@router.put("/model-monthly-memos/{model_name}/{month}")
async def update_model_monthly_memo(model_name: str, month: str, body: dict):
    memo = body.get("memo", "").strip()
    if not memo:
        raise HTTPException(status_code=400, detail="memo 필드가 필요합니다.")
    col = get_collection("model_monthly_memos")
    await col.update_one(
        {"model_name": model_name, "month": month},
        {"$set": {"memo": memo, "updated_date": _now()}},
    )
    return {"success": True}


@router.delete("/model-monthly-memos/{model_name}/{month}")
async def delete_model_monthly_memo(model_name: str, month: str):
    col = get_collection("model_monthly_memos")
    await col.delete_one({"model_name": model_name, "month": month})
    return {"success": True}


# ─── 메모 백업/복원 ────────────────────────────────────────────────────────────

@router.post("/memos/backup")
async def backup_memos():
    result = {}
    for col_name in ["monthly_memos", "weekly_memos", "model_monthly_memos"]:
        col = get_collection(col_name)
        docs = await col.find({}, {"_id": 0}).to_list(None)
        result[col_name] = docs
    return result


@router.post("/memos/restore")
async def restore_memos(body: dict):
    restored = 0
    for col_name in ["monthly_memos", "weekly_memos", "model_monthly_memos"]:
        if col_name in body and body[col_name]:
            col = get_collection(col_name)
            await col.delete_many({})
            await col.insert_many(body[col_name])
            restored += len(body[col_name])
    return {"success": True, "restored": restored}
