"""출시일/개통일 비교 라우터"""
from fastapi import APIRouter, Query
from datetime import datetime, date, timedelta
from database import get_collection

router = APIRouter(prefix="/api/launch", tags=["launch"])


@router.get("/models")
async def get_launch_models():
    """출시일이 등록된 모델 목록 반환"""
    col = get_collection("launch_dates")
    docs = await col.find({}, {"_id": 0, "model_name": 1, "launch_date": 1}).to_list(None)
    return sorted(docs, key=lambda x: x["model_name"])


@router.get("/compare")
async def compare_by_activation_day(models: str = Query(...)):
    """
    여러 모델의 개통일별 VOC/Q-data 건수 비교
    models: 쉼표로 구분된 모델명 (예: "ModelA,ModelB")
    기준 모델(첫 번째)의 현재 개통일 수를 기준으로 비교
    """
    model_names = [m.strip() for m in models.split(",") if m.strip()]
    if not model_names:
        return []

    launch_col = get_collection("launch_dates")
    voc_col = get_collection("internal_voc")
    qdata_col = get_collection("q_data")
    today = date.today()

    # 기준 모델(첫 번째)의 최대 개통일 수 계산
    ref_launch_doc = await launch_col.find_one({"model_name": model_names[0]})
    if not ref_launch_doc:
        return [{"model_name": model_names[0], "error": "출시일 미등록"}]

    try:
        ref_launch = datetime.strptime(ref_launch_doc["launch_date"], "%Y-%m-%d").date()
        max_days = (today - ref_launch).days + 1
    except Exception:
        return [{"model_name": model_names[0], "error": "출시일 형식 오류"}]

    if max_days <= 0:
        return [{"model_name": model_names[0], "error": "출시일이 미래입니다"}]

    result = []

    for model_name in model_names:
        launch_doc = await launch_col.find_one({"model_name": model_name})
        if not launch_doc:
            result.append({"model_name": model_name, "error": "출시일 미등록"})
            continue

        try:
            launch_date = datetime.strptime(launch_doc["launch_date"], "%Y-%m-%d").date()
        except Exception:
            result.append({"model_name": model_name, "error": "출시일 형식 오류"})
            continue

        # 이 모델의 max_days 범위 날짜
        range_start = launch_date
        range_end = launch_date + timedelta(days=max_days - 1)
        range_start_str = range_start.strftime("%Y-%m-%d")
        range_end_str = range_end.strftime("%Y-%m-%d")

        # VOC 날짜별 집계
        voc_pipeline = [
            {"$match": {
                "model_name": model_name,
                "created_date": {"$gte": range_start_str, "$lte": range_end_str + "~"},
            }},
            {"$project": {"date_str": {"$substr": ["$created_date", 0, 10]}}},
            {"$group": {"_id": "$date_str", "count": {"$sum": 1}}},
        ]
        voc_docs = await voc_col.aggregate(voc_pipeline).to_list(None)
        voc_by_date = {d["_id"]: d["count"] for d in voc_docs}

        # Q-data 날짜별 집계
        qdata_pipeline = [
            {"$match": {
                "model_name": model_name,
                "service_date": {"$gte": range_start_str, "$lte": range_end_str},
            }},
            {"$group": {"_id": "$service_date", "count": {"$sum": 1}}},
        ]
        qdata_docs = await qdata_col.aggregate(qdata_pipeline).to_list(None)
        qdata_by_date = {d["_id"]: d["count"] for d in qdata_docs}

        daily_data = []
        for day_num in range(1, max_days + 1):
            day_date = launch_date + timedelta(days=day_num - 1)
            day_str = day_date.strftime("%Y-%m-%d")
            daily_data.append({
                "day": day_num,
                "date": day_str,
                "voc_count": voc_by_date.get(day_str, 0),
                "qdata_count": qdata_by_date.get(day_str, 0),
            })

        result.append({
            "model_name": model_name,
            "display_name": launch_doc.get("marketing_name") or model_name,
            "launch_date": launch_doc["launch_date"],
            "max_days": max_days,
            "daily_data": daily_data,
        })

    return result
