"""통계 라우터"""
import asyncio
import re
import calendar as cal
from fastapi import APIRouter, Query
from typing import Optional, List
from datetime import datetime, timedelta
from database import get_collection

router = APIRouter(prefix="/api/statistics", tags=["statistics"])


# ── 마케팅명 헬퍼 ─────────────────────────────────────────────────────────────

async def _get_marketing_map() -> dict:
    """Returns {model_name: marketing_name} for models that have a marketing name."""
    col = get_collection("launch_dates")
    docs = await col.find(
        {"marketing_name": {"$exists": True, "$nin": [None, ""]}},
        {"_id": 0, "model_name": 1, "marketing_name": 1},
    ).to_list(None)
    return {
        d["model_name"]: d["marketing_name"]
        for d in docs
        if d.get("marketing_name", "").strip()
    }


def _resolve_name(orig: str, mmap: dict) -> str:
    """Resolve effective marketing name.
    1) Exact match
    2) SM- prefix: first 4 chars after 'SM-' must match (e.g. SM-L325N → SM-L325 key)
    """
    if not orig:
        return orig
    if orig in mmap:
        return mmap[orig]
    upper = orig.upper()
    if upper.startswith("SM-") and len(orig) >= 7:
        prefix = orig[:7]  # "SM-" + 4 chars
        for key, val in mmap.items():
            if key[:7] == prefix:
                return val
    return orig


def _apply_marketing(items: list, mmap: dict) -> list:
    """Group model-level stats by effective name (marketing_name || model_name)."""
    merged: dict = {}
    for item in items:
        orig = item.get("model_name") or ""
        eff = _resolve_name(orig, mmap)
        if not eff:
            continue
        if eff not in merged:
            merged[eff] = {"model_name": eff, "count": 0}
        merged[eff]["count"] += item.get("count", 0)
    return sorted(merged.values(), key=lambda x: x["count"], reverse=True)

def _date_filter(start_date: Optional[str], end_date: Optional[str]) -> dict:
    if start_date and end_date:
        return {"created_date": {"$gte": start_date, "$lte": end_date + "T"}}
    return {}


def _extract_summary(problem, original_content, reproduction) -> str:
    """M열(problem) 또는 N열(original_content/reproduction)에서 요약 추출"""
    if problem:
        m = re.search(r'내용\s*:\s*(.+)', str(problem), re.DOTALL)
        if m:
            return re.sub(r'\s+', ' ', m.group(1)).strip()[:160]
    if original_content:
        text = re.sub(r'[Qq]ueti?on_?[Tt]ime.*', '', str(original_content), flags=re.DOTALL)
        text = re.sub(r'\s+', ' ', text).strip()
        if text:
            return text[:160]
    if reproduction:
        text = re.sub(r'\[[^\]]+\]\s*', '', str(reproduction))
        text = re.sub(r'\s+', ' ', text).strip()
        return text[:160]
    return ""


@router.get("/dashboard")
async def get_dashboard():
    voc_col = get_collection("internal_voc")
    qdata_col = get_collection("q_data")

    today_dt = datetime.now().date()
    yesterday = today_dt - timedelta(days=1)
    yest_str = yesterday.strftime("%Y-%m-%d")

    # Q-data 유효 날짜: 어제 데이터가 없으면 가장 최근 데이터 날짜 사용 (주말/공휴일 대응)
    recent_qdata = await qdata_col.find_one(
        {"service_date": {"$lte": yest_str}},
        sort=[("service_date", -1)],
        projection={"service_date": 1, "_id": 0},
    )
    effective_qdata_date = recent_qdata["service_date"] if recent_qdata else yest_str

    # WoW 비교: 최근 7일 vs 이전 7일
    d7_start = (yesterday - timedelta(days=6)).strftime("%Y-%m-%d")
    d7_prev_start = (yesterday - timedelta(days=13)).strftime("%Y-%m-%d")
    d7_prev_end = (yesterday - timedelta(days=7)).strftime("%Y-%m-%d")
    # 이상탐지 기준: 어제 기준 7일 평균
    anomaly_start = (yesterday - timedelta(days=7)).strftime("%Y-%m-%d")
    anomaly_end = (yesterday - timedelta(days=1)).strftime("%Y-%m-%d")

    # 10주 추이 (월~일 기준)
    monday = today_dt - timedelta(days=today_dt.weekday())
    week_ranges, week_labels = [], []
    for i in range(9, -1, -1):
        ws = monday - timedelta(weeks=i)
        we = ws + timedelta(days=6)
        week_ranges.append((ws.strftime("%Y-%m-%d"), we.strftime("%Y-%m-%d")))
        week_labels.append("이번주" if i == 0 else "지난주" if i == 1 else f"{i+1}주전")

    # 모든 쿼리 병렬 실행
    results = await asyncio.gather(
        qdata_col.aggregate([
            {"$match": {"service_date": effective_qdata_date}},
            {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
            {"$sort": {"count": -1}}, {"$limit": 5},
            {"$project": {"model_name": "$_id", "count": 1, "_id": 0}},
        ]).to_list(None),
        qdata_col.aggregate([
            {"$match": {"service_date": {"$gte": d7_start, "$lte": yest_str}}},
            {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
        ]).to_list(None),
        qdata_col.aggregate([
            {"$match": {"service_date": {"$gte": d7_prev_start, "$lte": d7_prev_end}}},
            {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
        ]).to_list(None),
        voc_col.find(
            {"created_date": {"$regex": f"^{yest_str}"}},
            {"_id": 0, "case_code": 1, "model_name": 1, "problem": 1,
             "original_content": 1, "reproduction_path": 1},
        ).sort("model_name", 1).to_list(200),
        voc_col.aggregate([
            {"$match": {"created_date": {"$gte": anomaly_start, "$lte": anomaly_end + "~"}}},
            {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
        ]).to_list(None),
        qdata_col.count_documents({"service_date": effective_qdata_date}),
        *[voc_col.count_documents({"created_date": {"$gte": ws, "$lte": we + "~"}}) for ws, we in week_ranges],
        *[qdata_col.count_documents({"service_date": {"$gte": ws, "$lte": we}}) for ws, we in week_ranges],
        voc_col.find_one({}, sort=[("uploaded_date", -1)], projection={"uploaded_date": 1, "_id": 0}),
        qdata_col.find_one({}, sort=[("uploaded_date", -1)], projection={"uploaded_date": 1, "_id": 0}),
    )

    (qdata_top5, qdata_this_w, qdata_prev_w, voc_yest_docs, voc_7d_docs,
     qdata_yest_count, *rest) = results
    n_weeks = len(week_ranges)
    trend_voc = rest[:n_weeks]
    trend_qdata = rest[n_weeks:n_weeks * 2]
    voc_last_doc, qdata_last_doc = rest[n_weeks * 2], rest[n_weeks * 2 + 1]

    # 마케팅명 매핑
    mmap = await _get_marketing_map()

    # Members issue 어제 리스트 + 요약
    voc_entries = [
        {
            "case_code": d.get("case_code"),
            "model_name": mmap.get(d.get("model_name"), d.get("model_name")) or "-",
            "summary": _extract_summary(d.get("problem"), d.get("original_content"), d.get("reproduction_path")),
        }
        for d in voc_yest_docs
    ]

    # 이상 탐지
    yest_model_map: dict = {}
    for d in voc_yest_docs:
        m = d.get("model_name") or "미분류"
        yest_model_map[m] = yest_model_map.get(m, 0) + 1
    d7_avg_map = {d["_id"]: d["count"] / 7.0 for d in voc_7d_docs if d["_id"]}
    anomalies = []
    for model, ycount in yest_model_map.items():
        avg = d7_avg_map.get(model, 0)
        if (avg > 0 and ycount >= avg * 2 and ycount >= 2) or (avg == 0 and ycount >= 5):
            anomalies.append({
                "model_name": mmap.get(model, model),
                "yesterday_count": ycount,
                "avg_7d": round(avg, 1),
                "ratio": round(ycount / avg, 1) if avg > 0 else None,
            })
    anomalies.sort(key=lambda x: x["yesterday_count"], reverse=True)

    # Q-data Top5 마케팅명 적용
    qdata_top5 = [
        {"model_name": mmap.get(d["model_name"], d["model_name"]), "count": d["count"]}
        for d in qdata_top5
    ]

    # Q-data WoW Top5 + 마케팅명
    this_map = {d["_id"]: d["count"] for d in qdata_this_w if d["_id"]}
    prev_map = {d["_id"]: d["count"] for d in qdata_prev_w if d["_id"]}
    wow = sorted(
        [{"model_name": mmap.get(m, m), "this_week": c, "last_week": prev_map.get(m, 0), "diff": c - prev_map.get(m, 0)}
         for m, c in this_map.items() if c - prev_map.get(m, 0) > 0],
        key=lambda x: x["diff"], reverse=True
    )[:5]

    return {
        "yesterday": yest_str,
        "qdata_date": effective_qdata_date,
        "members_issue": {
            "yesterday_count": len(voc_entries),
            "entries": voc_entries,
            "anomaly": anomalies[:5],
        },
        "qdata": {
            "yesterday_count": qdata_yest_count,
            "top5_yesterday": qdata_top5,
            "top5_wow": wow,
        },
        "weekly_trend": [
            {"label": lbl, "week_start": ws, "members_issue": vc, "qdata": qc}
            for lbl, (ws, _), vc, qc in zip(week_labels, week_ranges, trend_voc, trend_qdata)
        ],
        "upload_status": {
            "members_issue": voc_last_doc.get("uploaded_date") if voc_last_doc else None,
            "qdata": qdata_last_doc.get("uploaded_date") if qdata_last_doc else None,
        },
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
    result = await col.aggregate(pipeline).to_list(None)
    mmap = await _get_marketing_map()
    return _apply_marketing(result, mmap)


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
    voc_col = get_collection("internal_voc")
    qdata_col = get_collection("q_data")
    chipset_col = get_collection("chipset_mapping")

    voc_match = {"chipset": {"$ne": None}}
    if start_date and end_date:
        voc_match["created_date"] = {"$gte": start_date, "$lte": end_date + "~"}

    voc_docs = await voc_col.aggregate([
        {"$match": voc_match},
        {"$group": {"_id": "$chipset", "voc_count": {"$sum": 1}}},
    ]).to_list(None)
    voc_map = {d["_id"]: d["voc_count"] for d in voc_docs}

    # Q-data: model_name → chipset via chipset_mapping
    mapping_docs = await chipset_col.find({}, {"_id": 0, "model_name": 1, "chipset": 1}).to_list(None)
    model_to_chipset = {d["model_name"]: d["chipset"] for d in mapping_docs}

    qdata_match = {"model_name": {"$ne": None}}
    if start_date and end_date:
        qdata_match["service_date"] = {"$gte": start_date, "$lte": end_date}

    qdata_docs = await qdata_col.aggregate([
        {"$match": qdata_match},
        {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
    ]).to_list(None)

    qdata_map: dict = {}
    for d in qdata_docs:
        chipset = model_to_chipset.get(d["_id"])
        if chipset:
            qdata_map[chipset] = qdata_map.get(chipset, 0) + d["count"]

    all_chipsets = sorted(set(voc_map) | set(qdata_map),
                          key=lambda c: voc_map.get(c, 0) + qdata_map.get(c, 0), reverse=True)
    return [
        {"chipset": c, "voc_count": voc_map.get(c, 0), "qdata_count": qdata_map.get(c, 0)}
        for c in all_chipsets
    ]


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
    result, settings_doc = await asyncio.gather(
        col.aggregate(pipeline).to_list(None),
        get_collection("qdata_settings").find_one({"key": "last_ppm"}),
    )

    last_ppm = settings_doc["value"] if settings_doc else None

    for item in result:
        item["ppm"] = last_ppm

    mmap = await _get_marketing_map()
    result = _apply_marketing(result, mmap)

    # 건수 내림차순 정렬 (PPM은 전체 공통 값)
    result.sort(key=lambda x: -x["count"])
    return result


@router.get("/qdata/weekly")
async def get_qdata_weekly_statistics(
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
                "_id": {
                    "$dateToString": {
                        "format": "%Y-%V",
                        "date": {
                            "$dateFromString": {
                                "dateString": {"$substr": ["$service_date", 0, 10]},
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


# ── 마케팅명 통합 엔드포인트 ──────────────────────────────────────────────────

@router.get("/effective-models")
async def get_effective_models():
    """VOC + Q-data 전체 모델 목록을 마케팅명으로 통합해 반환"""
    voc_col = get_collection("internal_voc")
    qdata_col = get_collection("q_data")
    mmap = await _get_marketing_map()

    voc_models = set(await voc_col.distinct("model_name")) - {None}
    qdata_models = set(await qdata_col.distinct("model_name")) - {None}
    effective = {mmap.get(m, m) for m in (voc_models | qdata_models) if m}
    return sorted(effective)


@router.get("/effective-name/monthly")
async def get_effective_name_monthly(
    name: str = Query(...),
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
):
    """마케팅명(또는 모델명) 기준으로 VOC + Q-data 월별 통계 반환"""
    mmap = await _get_marketing_map()
    models = [m for m, mn in mmap.items() if mn == name] or [name]

    voc_col = get_collection("internal_voc")
    qdata_col = get_collection("q_data")
    memo_col = get_collection("monthly_memos")

    voc_match = {"model_name": {"$in": models}, "created_date": {"$ne": None}}
    if start_date and end_date:
        voc_match["created_date"] = {"$gte": start_date, "$lte": end_date + "~"}

    voc_monthly = await voc_col.aggregate([
        {"$match": voc_match},
        {"$group": {"_id": {"$substr": ["$created_date", 0, 7]}, "count": {"$sum": 1}}},
        {"$sort": {"_id": 1}},
        {"$project": {"month": "$_id", "count": 1, "_id": 0}},
    ]).to_list(None)

    qdata_match = {"model_name": {"$in": models}, "service_date": {"$ne": None}}
    if start_date and end_date:
        qdata_match["service_date"] = {"$gte": start_date, "$lte": end_date}

    qdata_monthly = await qdata_col.aggregate([
        {"$match": qdata_match},
        {"$group": {"_id": {"$substr": ["$service_date", 0, 7]}, "count": {"$sum": 1}}},
        {"$sort": {"_id": 1}},
        {"$project": {"month": "$_id", "count": 1, "_id": 0}},
    ]).to_list(None)

    memos = {d["month"]: d["memo"] for d in await memo_col.find(
        {}, {"_id": 0, "month": 1, "memo": 1}
    ).to_list(None)}
    for item in voc_monthly:
        item["memo"] = memos.get(item["month"], "")

    return {"voc_monthly": voc_monthly, "qdata_monthly": qdata_monthly}


@router.get("/month-detail/{month}")
async def get_month_detail(month: str):
    """월별 상세: MoM 증감, 칩셋별, 주차별, 3개월 Top5 연속, 신규 모델"""
    year, mon = int(month[:4]), int(month[5:7])
    _, last_day = cal.monthrange(year, mon)
    month_start = f"{month}-01"
    month_end = f"{month}-{last_day:02d}"

    py, pm = (year - 1, 12) if mon == 1 else (year, mon - 1)
    prev_month = f"{py}-{pm:02d}"
    _, prev_last = cal.monthrange(py, pm)
    prev_start, prev_end = f"{prev_month}-01", f"{prev_month}-{prev_last:02d}"

    py2, pm2 = (py - 1, 12) if pm == 1 else (py, pm - 1)
    two_ago = f"{py2}-{pm2:02d}"
    _, two_last = cal.monthrange(py2, pm2)
    two_start, two_end = f"{two_ago}-01", f"{two_ago}-{two_last:02d}"

    mmap = await _get_marketing_map()
    voc_col = get_collection("internal_voc")
    qdata_col = get_collection("q_data")

    async def voc_model_agg(s: str, e: str) -> dict:
        docs = await voc_col.aggregate([
            {"$match": {"created_date": {"$gte": s, "$lte": e + "~"}, "model_name": {"$ne": None}}},
            {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
            {"$project": {"model_name": "$_id", "count": 1, "_id": 0}},
        ]).to_list(None)
        return {i["model_name"]: i["count"] for i in _apply_marketing(docs, mmap)}

    async def qdata_model_agg(s: str, e: str) -> dict:
        docs = await qdata_col.aggregate([
            {"$match": {"service_date": {"$gte": s, "$lte": e}, "model_name": {"$ne": None}}},
            {"$group": {"_id": "$model_name", "count": {"$sum": 1}}},
            {"$project": {"model_name": "$_id", "count": 1, "_id": 0}},
        ]).to_list(None)
        return {i["model_name"]: i["count"] for i in _apply_marketing(docs, mmap)}

    voc_cur, voc_prev, voc_two, qdata_cur, qdata_prev = await asyncio.gather(
        voc_model_agg(month_start, month_end),
        voc_model_agg(prev_start, prev_end),
        voc_model_agg(two_start, two_end),
        qdata_model_agg(month_start, month_end),
        qdata_model_agg(prev_start, prev_end),
    )

    def top5_set(d: dict) -> set:
        return set(sorted(d, key=d.get, reverse=True)[:5])

    streak_models = top5_set(voc_cur) & top5_set(voc_prev) & top5_set(voc_two)

    prev_raw = set(await voc_col.distinct("model_name", {"created_date": {"$lt": month_start}}))
    prev_effective = {mmap.get(m, m) for m in prev_raw if m}

    def build_model_list(cur: dict, prev: dict, with_badges: bool = False) -> list:
        rows = []
        for name in set(cur) | set(prev):
            c, p = cur.get(name, 0), prev.get(name, 0)
            diff = c - p
            row = {
                "name": name, "count": c, "prev_count": p, "diff": diff,
                "diff_pct": round(diff / p * 100, 1) if p else None,
            }
            if with_badges:
                row["is_top5_streak"] = name in streak_models
                row["is_new"] = name not in prev_effective and c > 0
            rows.append(row)
        return sorted(rows, key=lambda x: x["count"], reverse=True)

    chipsets = await voc_col.aggregate([
        {"$match": {"created_date": {"$gte": month_start, "$lte": month_end + "~"}, "chipset": {"$ne": None}}},
        {"$group": {"_id": "$chipset", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}},
        {"$project": {"chipset": "$_id", "count": 1, "_id": 0}},
    ]).to_list(None)

    def weekly_agg(date_field: str, start: str, end: str, suffix: str = "") -> list:
        return [
            {"$match": {date_field: {"$gte": start, "$lte": end + suffix}}},
            {"$project": {"day": {"$toInt": {"$substr": [f"${date_field}", 8, 2]}}}},
            {"$project": {"week": {"$switch": {"branches": [
                {"case": {"$lte": ["$day", 7]}, "then": 1},
                {"case": {"$lte": ["$day", 14]}, "then": 2},
                {"case": {"$lte": ["$day", 21]}, "then": 3},
                {"case": {"$lte": ["$day", 28]}, "then": 4},
            ], "default": 5}}}},
            {"$group": {"_id": "$week", "count": {"$sum": 1}}},
            {"$sort": {"_id": 1}},
            {"$project": {"week": "$_id", "count": 1, "_id": 0}},
        ]

    voc_weekly, qdata_weekly = await asyncio.gather(
        voc_col.aggregate(weekly_agg("created_date", month_start, month_end, "~")).to_list(None),
        qdata_col.aggregate(weekly_agg("service_date", month_start, month_end)).to_list(None),
    )

    return {
        "month": month,
        "prev_month": prev_month,
        "streak_models": list(streak_models),
        "voc": {
            "total": sum(voc_cur.values()),
            "prev_total": sum(voc_prev.values()),
            "models": build_model_list(voc_cur, voc_prev, with_badges=True),
            "chipsets": chipsets,
            "weekly": voc_weekly,
        },
        "qdata": {
            "total": sum(qdata_cur.values()),
            "prev_total": sum(qdata_prev.values()),
            "models": build_model_list(qdata_cur, qdata_prev),
            "weekly": qdata_weekly,
        },
    }
