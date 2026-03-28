from pydantic import BaseModel, Field
from typing import Optional
from datetime import datetime
from bson import ObjectId


class PyObjectId(str):
    @classmethod
    def __get_validators__(cls):
        yield cls.validate

    @classmethod
    def validate(cls, v, info=None):
        if not ObjectId.is_valid(v):
            raise ValueError("Invalid ObjectId")
        return str(v)


class VocBase(BaseModel):
    case_code: str
    title: Optional[str] = None
    model_name: Optional[str] = None
    model_no: Optional[str] = None
    chipset: Optional[str] = None
    build_version: Optional[str] = None
    os_version: Optional[str] = None
    issue_type: Optional[str] = None
    problem: Optional[str] = None
    original_content: Optional[str] = None
    reproduction_path: Optional[str] = None
    resolver: Optional[str] = None
    resolve_option: Optional[str] = None
    cause: Optional[str] = None
    solution: Optional[str] = None
    third_party_app: Optional[str] = None
    created_date: Optional[str] = None
    uploaded_date: Optional[str] = None


class VocCreate(VocBase):
    pass


class VocUpdate(BaseModel):
    model_name: Optional[str] = None
    chipset: Optional[str] = None
    cause: Optional[str] = None
    solution: Optional[str] = None


class VocInDB(VocBase):
    id: Optional[str] = Field(default=None, alias="_id")

    model_config = {"populate_by_name": True, "arbitrary_types_allowed": True}


class CommentBase(BaseModel):
    voc_case_code: str
    voc_type: str = "internal"
    comment: str


class CommentCreate(CommentBase):
    pass


class CommentInDB(CommentBase):
    id: Optional[str] = Field(default=None, alias="_id")
    created_date: str

    model_config = {"populate_by_name": True}


class ChipsetMapping(BaseModel):
    model_name: str
    chipset: str


class AppKeyword(BaseModel):
    app_name: str
    keywords: str


class NotificationSettings(BaseModel):
    enabled: bool = True
    notification_time: str = "09:00"
