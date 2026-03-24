"""
MongoDB 연결 모듈
"""
from pymongo import MongoClient, ASCENDING, DESCENDING

MONGO_URI = 'mongodb://localhost:27017'
DB_NAME = 'voc_database'

_client = MongoClient(MONGO_URI)
_db = _client[DB_NAME]

# 컬렉션
internal_voc = _db['internal_voc']
chipset_mapping = _db['chipset_mapping']
app_keywords = _db['app_keywords']
comments = _db['comments']
notification_settings = _db['notification_settings']
monthly_memos = _db['monthly_memos']
weekly_memos = _db['weekly_memos']
model_monthly_memos = _db['model_monthly_memos']
q_data = _db['q_data']


def init_db():
    """인덱스 생성"""
    internal_voc.create_index('case_code', unique=True)
    chipset_mapping.create_index('model_name', unique=True)
    monthly_memos.create_index('month', unique=True)
    weekly_memos.create_index('week', unique=True)
    model_monthly_memos.create_index(
        [('model_name', ASCENDING), ('month', ASCENDING)], unique=True
    )
    q_data.create_index(
        [('serial_number', ASCENDING), ('service_date', ASCENDING)],
        unique=True,
        sparse=True
    )
    q_data.create_index('model_name')
    q_data.create_index('service_date')
    q_data.create_index('repair_name')
    q_data.create_index('process_type')
    print("MongoDB 인덱스 생성 완료")
