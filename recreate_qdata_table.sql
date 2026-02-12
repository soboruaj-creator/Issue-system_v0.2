-- Q-data 테이블 재생성 (중복 방지 강화)

-- 기존 테이블 삭제 (주의: 데이터 모두 삭제됨!)
DROP TABLE IF EXISTS q_data;

-- Q-data 테이블 생성
CREATE TABLE q_data (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    service_date TEXT NOT NULL,           -- F열: 서비스일자 (YYYY-MM-DD)
    process_type TEXT,                     -- M열: 처리유형 (자재미사용/자재사용/반품)
    repair_name TEXT,                      -- P열: 수리명
    repair_detail TEXT,                    -- Q열: 수리 세부 내용
    detail_content TEXT,                   -- T열: 상세내용
    model_name TEXT NOT NULL,              -- Z열: 모델명
    serial_number TEXT NOT NULL,           -- AD열: S/N (고유, NOT NULL 추가)
    log_id TEXT NOT NULL,                  -- AR열: LOG ID (고유, NOT NULL 추가)
    sw_before TEXT,                        -- BE열: 수리전 S/W
    sw_after TEXT,                         -- BF열: 수리 S/W
    uploaded_date TEXT NOT NULL,           -- 업로드 일시
    
    -- 중복 방지: S/N + LOG ID 조합이 고유해야 함
    UNIQUE(serial_number, log_id)
);

-- 인덱스 생성 (검색 성능 향상)
CREATE INDEX idx_q_data_model ON q_data(model_name);
CREATE INDEX idx_q_data_service_date ON q_data(service_date);
CREATE INDEX idx_q_data_repair_name ON q_data(repair_name);
CREATE INDEX idx_q_data_process_type ON q_data(process_type);
CREATE INDEX idx_q_data_sn_logid ON q_data(serial_number, log_id);  -- 중복 체크 성능 향상
