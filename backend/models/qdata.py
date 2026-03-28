from pydantic import BaseModel, Field
from typing import Optional


class QDataBase(BaseModel):
    service_date: Optional[str] = None
    process_type: Optional[str] = None
    repair_name: Optional[str] = None
    repair_detail: Optional[str] = None
    detail_content: Optional[str] = None
    model_name: Optional[str] = None
    serial_number: Optional[str] = None
    log_id: Optional[str] = None
    sw_before: Optional[str] = None
    sw_after: Optional[str] = None
    uploaded_date: Optional[str] = None


class QDataCreate(QDataBase):
    pass


class QDataInDB(QDataBase):
    id: Optional[str] = Field(default=None, alias="_id")

    model_config = {"populate_by_name": True, "arbitrary_types_allowed": True}


class MemoBase(BaseModel):
    memo: str


class MonthlyMemo(MemoBase):
    month: str  # YYYY-MM format
    created_date: Optional[str] = None
    updated_date: Optional[str] = None


class WeeklyMemo(MemoBase):
    week: str  # YYYY-WW format
    created_date: Optional[str] = None
    updated_date: Optional[str] = None


class ModelMonthlyMemo(MemoBase):
    model_name: str
    month: str
    created_date: Optional[str] = None
    updated_date: Optional[str] = None
