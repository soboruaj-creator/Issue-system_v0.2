"""통계 라우터"""
from fastapi import APIRouter, Query
from typing import Optional, List
from datetime import datetime, timedelta
from database import get_collection

router = APIRouter(prefix="/api/statistics", tags=["statistics"])


def _date_filter(start_date: Optional[str], end_date: Optional[str]) -> dict:
    if start_date and end_date:
        return {"created_date": {"$gte": start_date, "$lte": end_date + "T"}}
    return {}


@router.get("/dashboard")
async def get_dashboard():
    col = get_collection("internal_voc")
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

    total = await col.count_documents({})
    daily_count = await col.count_documents({"created_date": {"$regex": f"^{yesterday}"}})

    pipeline = [
        {"$match": {"created_date": {"$regex": f"^{yesterday}"}, "model_name": {"$ne": None}}},
        {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}},
        {"$limit": 10},
        {"$project": {"model_name": "$_id", "count": 1, "_id": 0}},
    ]
    top10 = await col.aggregate(pipeline).to_list(None)

    return {
        "yesterday_date": yesterday,
        "total_count": total,
        "daily_count": daily_count,
        "top10_models": top10,
    }


@router.get("/model")
async def get_model_statistics(
    start_date: Optional[str] = None, end_date: Optional[str] = None
):
    col = get_collection("internal_voc")
    match = {"model_name": {"$ne": None}}
    if start_date and end_date:
        match["created_date"] = {"$gte": start_date, "$lte": end_date + "~"}

    pipeline = [
        {"$match": match},
        {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}},
        {"$project": {"model_name": "$_id", "count": 1, "_id": 0}},
    ]
    return await col.aggregate(pipeline).to_list(None)


@router.get("/weekly")
async def get_weekly_statistics(
    start_date: Optional[str] = None, end_date: Optional[str] = None
):
    col = get_collection("internal_voc")
    memo_col = get_collection("weekly_memos")

    match = {"created_date": {"$ne": None}}
    if start_date and end_date:
        match["created_date"]["$gte"] = start_date
        match["created_date"]["$lte"] = end_date + "~"

    # MongoDB에서 주차 집계 (ISO week)
    pipeline = [
        {"$match": match},
        {
            "$group": {
                "_id": {
                    "$dateToString": {
                        "format": "%Y-%V",
                        "date": {
                            "$dateFromString": {
                                "dateString": {"$substr": ["$created_date", 0, 10]},
                                "onError": None,
                            }
                        },
                    }
                },
                "count": {"$sum": 1},
            }
        },
        {"$match": {"_id": {"$ne": None}}},
        {"$sort": {"_id": 1}},
        {"$project": {"week": "$_id", "count": 1, "_id": 0}},
    ]
    weekly = await col.aggregate(pipeline).to_list(None)

    memos_docs = await memo_col.find({}, {"_id": 0, "week": 1, "memo": 1}).to_list(None)
    memos = {d["week"]: d["memo"] for d in memos_docs}

    for item in weekly:
        item["memo"] = memos.get(item["week"], "")

    return weekly


@router.get("/monthly")
async def get_monthly_statistics(
    start_date: Optional[str] = None, end_date: Optional[str] = None
):
    col = get_collection("internal_voc")
    memo_col = get_collection("monthly_memos")

    match = {"created_date": {"$ne": None}}
    if start_date and end_date:
        match["created_date"] = {"$gte": start_date, "$lte": end_date + "~"}

    pipeline = [
        {"$match": match},
        {
            "$group": {
                "_id": {"$substr": ["$created_date", 0, 7]},
                "count": {"$sum": 1},
            }
        },
        {"$sort": {"_id": 1}},
        {"$project": {"month": "$_id", "count": 1, "_id": 0}},
    ]
    monthly = await col.aggregate(pipeline).to_list(None)

    memos_docs = await memo_col.find({}, {"_id": 0, "month": 1, "memo": 1}).to_list(None)
    memos = {d["month"]: d["memo"] for d in memos_docs}

    for item in monthly:
        item["memo"] = memos.get(item["month"], "")

    return monthly


@router.get("/chipset")
async def get_chipset_statistics(
    start_date: Optional[str] = None, end_date: Optional[str] = None
):
    col = get_collection("internal_voc")
    match = {"chipset": {"$ne": None}}
    if start_date and end_date:
        match["created_date"] = {"$gte": start_date, "$lte": end_date + "~"}

    pipeline = [
        {"$match": match},
        {"$group": {"_id": "$chipset", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}},
        {"$project": {"chipset": "$_id", "count": 1, "_id": 0}},
    ]
    return await col.aggregate(pipeline).to_list(None)


@router.get("/app")
async def get_app_statistics(
    start_date: Optional[str] = None, end_date: Optional[str] = None
):
    col = get_collection("internal_voc")
    match = {"third_party_app": {"$ne": None}}
    if start_date and end_date:
        match["created_date"] = {"$gte": start_date, "$lte": end_date + "~"}

    pipeline = [
        {"$match": match},
        {"$group": {"_id": "$third_party_app", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}},
        {"$project": {"app_name": "$_id", "count": 1, "_id": 0}},
    ]
    return await col.aggregate(pipeline).to_list(None)


@router.get("/model/{model_name}/monthly")
async def get_model_monthly_statistics(model_name: str):
    col = get_collection("internal_voc")
    pipeline = [
        {"$match": {"model_name": model_name, "created_date": {"$ne": None}}},
        {
            "$group": {
                "_id": {"$substr": ["$created_date", 0, 7]},
                "count": {"$sum": 1},
            }
        },
        {"$sort": {"_id": 1}},
        {"$project": {"month": "$_id", "count": 1, "_id": 0}},
    ]
    monthly = await col.aggregate(pipeline).to_list(None)

    memo_col = get_collection("model_monthly_memos")
    memos_docs = await memo_col.find(
        {"model_name": model_name}, {"_id": 0, "month": 1, "memo": 1}
    ).to_list(None)
    memos = {d["month"]: d["memo"] for d in memos_docs}

    for item in monthly:
        item["memo"] = memos.get(item["month"], "")

    return {"model_name": model_name, "monthly": monthly}


@router.post("/models/monthly")
async def get_models_monthly_statistics(body: dict):
    models = body.get("models", [])
    if not models:
        return []

    col = get_collection("internal_voc")
    pipeline = [
        {"$match": {"model_name": {"$in": models}, "created_date": {"$ne": None}}},
        {
            "$group": {
                "_id": {
                    "model_name": "$model_name",
                    "month": {"$substr": ["$created_date", 0, 7]},
                },
                "count": {"$sum": 1},
            }
        },
        {"$sort": {"_id.month": 1}},
        {"$project": {"model_name": "$_id.model_name", "month": "$_id.month", "count": 1, "_id": 0}},
    ]
    return await col.aggregate(pipeline).to_list(None)


@router.get("/qdata/model")
async def get_qdata_model_statistics(
    start_date: Optional[str] = None, end_date: Optional[str] = None
):
    col = get_collection("q_data")
    match = {"model_name": {"$ne": None}}
    if start_date and end_date:
        match["service_date"] = {"$gte": start_date, "$lte": end_date}

    pipeline = [
        {"$match": match},
        {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}},
        {"$project": {"model_name": "$_id", "count": 1, "_id": 0}},
    ]
    return await col.aggregate(pipeline).to_list(None)


@router.get("/qdata/monthly")
async def get_qdata_monthly_statistics(
    start_date: Optional[str] = None, end_date: Optional[str] = None
):
    col = get_collection("q_data")
    match = {"service_date": {"$ne": None}}
    if start_date and end_date:
        match["service_date"] = {"$gte": start_date, "$lte": end_date}

    pipeline = [
        {"$match": match},
        {
            "$group": {
                "_id": {"$substr": ["$service_date", 0, 7]},
                "count": {"$sum": 1},
            }
        },
        {"$sort": {"_id": 1}},
        {"$project": {"month": "$_id", "count": 1, "_id": 0}},
    ]
    return await col.aggregate(pipeline).to_list(None)


@router.post("/qdata/models/monthly")
async def get_qdata_models_monthly(body: dict):
    models = body.get("models", [])
    if not models:
        return []

    col = get_collection("q_data")
    pipeline = [
        {"$match": {"model_name": {"$in": models}, "service_date": {"$ne": None}}},
        {
            "$group": {
                "_id": {
                    "model_name": "$model_name",
                    "month": {"$substr": ["$service_date", 0, 7]},
                },
                "count": {"$sum": 1},
            }
        },
        {"$sort": {"_id.month": 1}},
        {"$project": {"model_name": "$_id.model_name", "month": "$_id.month", "count": 1, "_id": 0}},
    ]
    return await col.aggregate(pipeline).to_list(None)
