"""VOC 데이터 파싱 유틸리티"""
import re
import pandas as pd
from datetime import datetime
from typing import Optional


def extract_watch_model(title: Optional[str]) -> Optional[str]:
    if not title:
        return None
    match = re.search(r"watch\d*", title, re.IGNORECASE)
    if match:
        return match.group(0).upper()
    match = re.search(r"워치\d*", title, re.IGNORECASE)
    if match:
        return match.group(0).upper()
    if "watch" in title.lower():
        return "WATCH"
    if "워치" in title.lower():
        return "워치"
    return None


def map_model_name(model_name: Optional[str]) -> Optional[str]:
    if not model_name:
        return None
    upper = model_name.upper()
    if upper in ["SM-L705N", "SM-L705"]:
        return "워치울"
    if upper in ["SM-L310N", "SM-L310", "SM-L305N", "WATCH7"]:
        return "워치7"
    if upper in ["SM-R890", "SM-R870"]:
        return "워치4"
    if upper in ["SM-R935N", "SM-R960", "SM-R950", "SM-R940", "WATCH6"]:
        return "워치6"
    return model_name


def extract_model_from_title(title: Optional[str]) -> Optional[str]:
    if not title:
        return None
    match = re.search(r"SM-[A-Z0-9]{4,5}", title, re.IGNORECASE)
    return match.group(0).upper() if match else None


def extract_model_from_reproduction(text: Optional[str]) -> Optional[str]:
    if not text:
        return None
    match = re.search(r"\[Model No\.\]\s*([^\[\n]+)", text, re.IGNORECASE)
    if match:
        model = match.group(1).strip()
        sm_match = re.search(r"SM-[A-Z0-9]{4,5}", model, re.IGNORECASE)
        return sm_match.group(0).upper() if sm_match else model
    return None


def extract_build_version(text: Optional[str]) -> Optional[str]:
    if not text:
        return None
    match = re.search(r"\[Build No\.\]\s*([^\[\n]+)", text, re.IGNORECASE)
    if match:
        build = match.group(1).strip()
        return build[-3:] if len(build) >= 3 else build
    return None


def extract_os_version(text: Optional[str]) -> Optional[str]:
    if not text:
        return None
    match = re.search(r"\[OS Ver\.\]\s*([^\[\n]+)", text, re.IGNORECASE)
    return match.group(1).strip() if match else None


def extract_original_content(text: Optional[str]) -> Optional[str]:
    if not text:
        return None
    match = re.search(r"\[Original Contents\]\s*([^\[]+)", text, re.IGNORECASE)
    return match.group(1).strip() if match else None


def detect_issue_type(problem: Optional[str], reproduction: Optional[str]) -> str:
    text = f"{problem or ''} {reproduction or ''}"
    external_keywords = ["samsung members", "k zone", "rdm"]
    for keyword in external_keywords:
        if keyword.lower() in text.lower():
            return "외부이슈"
    return "내부이슈"


def normalize_chipset_name(chipset: Optional[str]) -> str:
    if not chipset:
        return ""
    if chipset.strip() == "JDM T618":
        return "jdm t618"
    if chipset.strip().upper().startswith("SM"):
        return chipset.strip().lower()
    normalized = chipset.lower()
    normalized = re.sub(r"[\s\(\)\[\]]+", "", normalized)
    normalized = re.sub(r"[^a-z0-9]", "", normalized)
    return normalized


def calculate_string_similarity(str1: str, str2: str) -> float:
    if not str1 or not str2:
        return 0.0
    longer = str1 if len(str1) >= len(str2) else str2
    shorter = str2 if len(str1) >= len(str2) else str1
    if len(longer) == 0:
        return 1.0
    common_chars = set(longer) & set(shorter)
    return len(common_chars) / len(set(longer))


def find_similar_chipset(new_chipset: str, existing_chipsets: list, threshold: float = 0.7) -> Optional[str]:
    if not new_chipset or not existing_chipsets:
        return None
    new_normalized = normalize_chipset_name(new_chipset)
    for existing in existing_chipsets:
        existing_normalized = normalize_chipset_name(existing)
        if new_normalized == existing_normalized:
            return existing
        if new_normalized in existing_normalized or existing_normalized in new_normalized:
            return existing
        if calculate_string_similarity(new_normalized, existing_normalized) >= threshold:
            return existing
    return None


def merge_similar_chipsets(chipsets: list) -> dict:
    if not chipsets:
        return {}
    excluded_patterns = [r"^sm", r"^jdm t618$"]
    chipset_groups = {}
    for chipset in chipsets:
        if any(re.match(p, chipset.lower()) for p in excluded_patterns):
            chipset_groups[chipset] = [chipset]
            continue
        normalized = normalize_chipset_name(chipset)
        if normalized not in chipset_groups:
            chipset_groups[normalized] = []
        chipset_groups[normalized].append(chipset)

    merged = {}
    for normalized, group in chipset_groups.items():
        if len(group) == 1:
            merged[group[0]] = group[0]
        else:
            representative = max(group, key=len)
            for c in group:
                merged[c] = representative
    return merged


def convert_qdata_date(date_str) -> Optional[str]:
    if pd.isna(date_str):
        return None
    try:
        s = str(int(date_str))
        return f"20{s[:2]}-{s[2:4]}-{s[4:6]}"
    except Exception:
        return None


def _parse_date_cell(val) -> Optional[str]:
    """Excel 날짜 셀 값을 YYYY-MM-DD 문자열로 변환. 실패 시 None 반환."""
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(val, "strftime"):
        try:
            return val.strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            return None
    if isinstance(val, (int, float)):
        try:
            ts = pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(val))
            result = ts.strftime("%Y-%m-%d")
            if "2000-" <= result <= "2031-":
                return result
        except Exception:
            pass
        return None
    if isinstance(val, str):
        s = val.strip()
        if not s or s.lower() in ("nan", "none", "-", ""):
            return None
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
            try:
                return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue
        try:
            ts = pd.to_datetime(s, dayfirst=False, errors="coerce")
            if pd.notna(ts):
                return ts.strftime("%Y-%m-%d")
        except Exception:
            pass
    return None


def process_voc_row(row: pd.Series, filename: str, chipset_map: dict, app_keywords: list) -> tuple:
    """단일 VOC 행 처리, (case_code, data_dict, is_unmapped) 반환"""
    try:
        def safe_str(val):
            return str(val) if pd.notna(val) else None

        case_code = safe_str(row.iloc[0])
        status_raw = safe_str(row.iloc[6])   # G열: 진행상태
        title = safe_str(row.iloc[7])
        problem = safe_str(row.iloc[12])
        reproduction = safe_str(row.iloc[13])
        resolver = safe_str(row.iloc[14])
        resolve_option = safe_str(row.iloc[17])
        cause = safe_str(row.iloc[20])
        solution = safe_str(row.iloc[21])

        if not case_code:
            return None, None, None

        watch_model = extract_watch_model(title)
        if watch_model:
            model_name = watch_model
            model_no = None
        else:
            model_from_title = extract_model_from_title(title)
            model_no = extract_model_from_reproduction(reproduction)
            model_name = model_no or model_from_title

        model_name = map_model_name(model_name)
        build_version = extract_build_version(reproduction)
        os_version = extract_os_version(reproduction)
        original_content = extract_original_content(reproduction)
        issue_type = detect_issue_type(problem, reproduction)

        chipset = chipset_map.get(model_name) if model_name else None

        search_text = f"{problem or ''} {original_content or ''}"
        third_party_app = _detect_third_party_app(search_text, app_keywords)

        # 날짜 추출 순서:
        # 1. 케이스코드 (P + 6자리 YYMMDD)
        # 2. 엑셀 B열 또는 C열 (등록일 컬럼)
        # 3. 파일명 8자리 숫자 (YYYYMMDD)
        created_date = None
        if case_code and case_code.startswith("P"):
            date_match = re.search(r"P(\d{6})", case_code)
            if date_match:
                try:
                    created_date = datetime.strptime(date_match.group(1), "%y%m%d").strftime("%Y-%m-%d")
                except Exception:
                    pass
        if not created_date:
            for col_idx in (1, 2):
                if col_idx < len(row):
                    created_date = _parse_date_cell(row.iloc[col_idx])
                    if created_date:
                        break
        if not created_date:
            date_match = re.search(r"(\d{8})", filename)
            if date_match:
                try:
                    created_date = datetime.strptime(date_match.group(1), "%Y%m%d").strftime("%Y-%m-%d")
                except Exception:
                    pass

        s = (status_raw or "").strip().lower()
        if s.startswith("resolve"):
            status = "resolve"
        elif s == "close":
            status = "close"
        else:
            status = "open"

        return case_code, {
            "status": status,
            "title": title,
            "model_name": model_name,
            "model_no": model_no if not watch_model else None,
            "chipset": chipset,
            "build_version": build_version,
            "os_version": os_version,
            "issue_type": issue_type,
            "problem": problem,
            "original_content": original_content,
            "reproduction_path": reproduction,
            "resolver": resolver,
            "resolve_option": resolve_option,
            "cause": cause,
            "solution": solution,
            "third_party_app": third_party_app,
            "created_date": created_date,
        }, (chipset is None and model_name is not None)

    except Exception as e:
        print(f"Row processing error: {e}")
        return None, None, None


def _detect_third_party_app(text: str, app_keywords: list) -> Optional[str]:
    if not text or not app_keywords:
        return None
    detected = []
    for item in app_keywords:
        keywords = [k.strip() for k in item.get("keywords", "").split(",")]
        for kw in keywords:
            if kw and kw.lower() in text.lower():
                detected.append(item["app_name"])
                break
    return ", ".join(detected) if detected else None
