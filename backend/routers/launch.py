"""출시일/개통일 비교 라우터"""
from fastapi import APIRouter, Query
from datetime import datetime, date, timedelta
from database import get_collection

router = APIRouter(prefix="/api/launch", tags=["launch"])


def _build_eff_map(docs: list) -> dict:
    """Returns {effective_name: [doc, ...]} grouping by marketing_name || model_name."""
    eff_map: dict = {}
    for doc in docs:
        eff = (doc.get("marketing_name") or "").strip() or doc["model_name"]
        if eff not in eff_map:
            eff_map[eff] = []
        eff_map[eff].append(doc)
    return eff_map


@router.get("/models")
async def get_launch_models():
    """마케팅명으로 그룹핑된 모델 목록 반환"""
    col = get_collection("launch_dates")
    docs = await col.find(
        {}, {"_id": 0, "model_name": 1, "launch_date": 1, "marketing_name": 1}
    ).to_list(None)

    eff_map = _build_eff_map(docs)
    result = []
    for eff_name, group in eff_map.items():
        earliest_date = min(d["launch_date"] for d in group)
        result.append({
            "effective_name": eff_name,
            "model_names": [d["model_name"] for d in group],
            "launch_date": earliest_date,
        })
    return sorted(result, key=lambda x: x["effective_name"])


@router.get("/compare")
async def compare_by_activation_day(models: str = Query(...)):
    """
    여러 그룹(마케팅명)의 개통일별 VOC/Q-data 건수 비교
    models: 쉼표로 구분된 effective_name (마케팅명 또는 모델명)
    기준 모델(첫 번째)의 현재 개통일 수를 기준으로 비교
    """
    eff_names = [m.strip() for m in models.split(",") if m.strip()]
    if not eff_names:
        return []

    launch_col = get_collection("launch_dates")
    voc_col = get_collection("internal_voc")
    qdata_col = get_collection("q_data")
    today = date.today()

    all_docs = await launch_col.find(
        {}, {"_id": 0, "model_name": 1, "launch_date": 1, "marketing_name": 1}
    ).to_list(None)
    eff_map = _build_eff_map(all_docs)

    # 기준: 첫 번째 그룹의 가장 이른 출시일 → max_days 산출
    first_group = eff_map.get(eff_names[0])
    if not first_group:
        return [{"model_name": eff_names[0], "display_name": eff_names[0], "error": "출시일 미등록"}]

    try:
        ref_launch = datetime.strptime(
            min(d["launch_date"] for d in first_group), "%Y-%m-%d"
        ).date()
        max_days = (today - ref_launch).days + 1
    except Exception:
        return [{"model_name": eff_names[0], "display_name": eff_names[0], "error": "출시일 형식 오류"}]

    if max_days <= 0:
        return [{"model_name": eff_names[0], "display_name": eff_names[0], "error": "출시일이 미래입니다"}]

    result = []

    for eff_name in eff_names:
        group_docs = eff_map.get(eff_name)
        if not group_docs:
            result.append({"model_name": eff_name, "display_name": eff_name, "error": "출시일 미등록"})
            continue

        try:
            launch_date = datetime.strptime(
                min(d["launch_date"] for d in group_docs), "%Y-%m-%d"
            ).date()
        except Exception:
            result.append({"model_name": eff_name, "display_name": eff_name, "error": "출시일 형식 오류"})
            continue

        model_names_in_group = [d["model_name"] for d in group_docs]
        range_start_str = launch_date.strftime("%Y-%m-%d")
        range_end_str = (launch_date + timedelta(days=max_days - 1)).strftime("%Y-%m-%d")

        # VOC 날짜별 집계 (그룹 내 모든 모델 합산)
        voc_pipeline = [
            {"$match": {
                "model_name": {"$in": model_names_in_group},
                "created_date": {"$gte": range_start_str, "$lte": range_end_str + "~"},
            }},
            {"$project": {"date_str": {"$substr": ["$created_date", 0, 10]}}},
            {"$group": {"_id": "$date_str", "count": {"$sum": 1}}},
        ]
        voc_docs = await voc_col.aggregate(voc_pipeline).to_list(None)
        voc_by_date = {d["_id"]: d["count"] for d in voc_docs}

        # Q-data 날짜별 집계 (그룹 내 모든 모델 합산)
        qdata_pipeline = [
            {"$match": {
                "model_name": {"$in": model_names_in_group},
                "service_date": {"$gte": range_start_str, "$lte": range_end_str},
            }},
            {"$group": {"_id": "$service_date", "count": {"$sum": 1}}},
        ]
        qdata_docs = await qdata_col.aggregate(qdata_pipeline).to_list(None)
        qdata_by_date = {d["_id"]: d["count"] for d in qdata_docs}

        # PPM 개통일차별 조회
        ppm_col = get_collection("ppm_data")
        ppm_docs = await ppm_col.find(
            {"model_name": {"$in": model_names_in_group}},
            {"_id": 0, "activation_day": 1, "ppm": 1, "upload_date": 1},
        ).to_list(None)
        ppm_by_day: dict = {}
        for d in ppm_docs:
            day = d["activation_day"]
            if day not in ppm_by_day or d.get("upload_date", "") >= ppm_by_day[day].get("upload_date", ""):
                ppm_by_day[day] = d
        ppm_by_day = {k: v["ppm"] for k, v in ppm_by_day.items()}

        cur_day = (today - launch_date).days + 1
        latest_ppm = ppm_by_day.get(cur_day)
        if latest_ppm is None and ppm_by_day:
            latest_ppm = ppm_by_day[max(ppm_by_day.keys())]

        daily_data = []
        for day_num in range(1, max_days + 1):
            day_date = launch_date + timedelta(days=day_num - 1)
            day_str = day_date.strftime("%Y-%m-%d")
            daily_data.append({
                "day": day_num,
                "date": day_str,
                "voc_count": voc_by_date.get(day_str, 0),
                "qdata_count": qdata_by_date.get(day_str, 0),
                "ppm": ppm_by_day.get(day_num),
            })

        result.append({
            "model_name": eff_name,
            "display_name": eff_name,
            "launch_date": launch_date.strftime("%Y-%m-%d"),
            "max_days": (today - launch_date).days + 1,
            "latest_ppm": latest_ppm,
            "model_names": model_names_in_group,
            "daily_data": daily_data,
        })

    return result
