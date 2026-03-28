"""엑셀 내보내기 라우터"""
from fastapi import APIRouter, Query
from fastapi.responses import StreamingResponse
from typing import Optional
from database import get_collection
import pandas as pd
import io

router = APIRouter(prefix="/api/export", tags=["export"])


@router.get("/excel")
async def export_voc_excel(
    start_date: Optional[str] = Query(None),
    end_date: Optional[str] = Query(None),
    model_name: Optional[str] = Query(None),
):
    col = get_collection("internal_voc")
    match = {}
    if start_date and end_date:
        match["created_date"] = {"$gte": start_date, "$lte": end_date + "~"}
    if model_name:
        match["model_name"] = model_name

    docs = await col.find(match, {"_id": 0}).sort("created_date", -1).to_list(None)

    if not docs:
        docs = []

    df = pd.DataFrame(docs)
    columns_order = [
        "case_code", "title", "model_name", "model_no", "chipset",
        "build_version", "os_version", "issue_type", "problem",
        "original_content", "reproduction_path", "resolver", "resolve_option",
        "cause", "solution", "third_party_app", "created_date", "uploaded_date",
    ]
    for col_name in columns_order:
        if col_name not in df.columns:
            df[col_name] = None
    df = df[[c for c in columns_order if c in df.columns]]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="VOC 데이터")
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=voc_export.xlsx"},
    )


@router.get("/qdata/excel")
async def export_qdata_excel(
    start_date: Optional[str] = Query(None),
    end_date: Optional[str] = Query(None),
    model_name: Optional[str] = Query(None),
):
    col = get_collection("q_data")
    match = {}
    if start_date and end_date:
        match["service_date"] = {"$gte": start_date, "$lte": end_date}
    if model_name:
        match["model_name"] = model_name

    docs = await col.find(match, {"_id": 0}).sort("service_date", -1).to_list(None)

    df = pd.DataFrame(docs if docs else [])
    columns_order = [
        "service_date", "process_type", "repair_name", "repair_detail",
        "detail_content", "model_name", "serial_number", "log_id",
        "sw_before", "sw_after", "uploaded_date",
    ]
    for col_name in columns_order:
        if col_name not in df.columns:
            df[col_name] = None
    if not df.empty:
        df = df[[c for c in columns_order if c in df.columns]]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Q-data")
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=qdata_export.xlsx"},
    )
