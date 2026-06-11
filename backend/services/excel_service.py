"""엑셀 파일 읽기 서비스 (DRM 처리 포함)"""
import io
import tempfile
import os
import pandas as pd
from fastapi import UploadFile


async def read_excel_with_drm(file: UploadFile, header: int = 0) -> pd.DataFrame:
    """DRM 우회 엑셀 읽기 - 8가지 방법 시도"""
    content = await file.read()
    last_error = None

    # 방법 1: openpyxl (메모리)
    try:
        df = pd.read_excel(io.BytesIO(content), engine="openpyxl", header=header)
        print("DRM 처리 성공: openpyxl (메모리)")
        return df
    except Exception as e:
        last_error = str(e)

    # 방법 2: xlrd (메모리)
    try:
        df = pd.read_excel(io.BytesIO(content), engine="xlrd", header=header)
        print("DRM 처리 성공: xlrd (메모리)")
        return df
    except Exception as e:
        last_error = str(e)

    # 방법 3: pyxlsb (메모리)
    try:
        df = pd.read_excel(io.BytesIO(content), engine="pyxlsb", header=header)
        print("DRM 처리 성공: pyxlsb (메모리)")
        return df
    except Exception as e:
        last_error = str(e)

    # 방법 4: openpyxl (임시 파일)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(content)
            temp_path = tmp.name
        df = pd.read_excel(temp_path, engine="openpyxl", header=header)
        print("DRM 처리 성공: openpyxl (임시 파일)")
        return df
    except Exception as e:
        last_error = str(e)
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

    # 방법 5: xlrd (임시 파일)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
            tmp.write(content)
            temp_path = tmp.name
        df = pd.read_excel(temp_path, engine="xlrd", header=header)
        print("DRM 처리 성공: xlrd (임시 파일)")
        return df
    except Exception as e:
        last_error = str(e)
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

    # 방법 6: pyxlsb (임시 파일)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsb") as tmp:
            tmp.write(content)
            temp_path = tmp.name
        df = pd.read_excel(temp_path, engine="pyxlsb", header=header)
        print("DRM 처리 성공: pyxlsb (임시 파일)")
        return df
    except Exception as e:
        last_error = str(e)
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

    # 방법 7: 기본 엔진 (메모리)
    try:
        df = pd.read_excel(io.BytesIO(content), header=header)
        print("DRM 처리 성공: 기본 엔진 (메모리)")
        return df
    except Exception as e:
        last_error = str(e)

    # 방법 8: 기본 엔진 (임시 파일)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp.write(content)
            temp_path = tmp.name
        df = pd.read_excel(temp_path, header=header)
        print("DRM 처리 성공: 기본 엔진 (임시 파일)")
        return df
    except Exception as e:
        last_error = str(e)
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

    raise Exception(
        f"모든 DRM 처리 방법이 실패했습니다. 마지막 오류: {last_error}\n"
        "해결 방법: 엑셀 파일을 열고 '다른 이름으로 저장' → Excel 통합 문서(*.xlsx) 형식으로 저장 후 재업로드"
    )


async def read_qdata_excel_with_drm(file: UploadFile) -> pd.DataFrame:
    """Q-data 전용 엑셀 읽기 — Z열(모델명 'SM-') 위치로 헤더행 자동 감지"""
    usecols = [5, 9, 12, 15, 16, 19, 25, 29, 43, 50, 51]  # F,J,M,P,Q,T,Z,AD,AR,BE,BF
    content = await file.read()
    last_error = None
    engines = ["openpyxl", "xlrd", "pyxlsb", None]

    def _detect_header_row(raw_df: pd.DataFrame) -> int:
        """Z열(절대 col index 25)에서 'SM-'으로 시작하는 첫 데이터행 직전을 헤더행으로 반환"""
        if raw_df.shape[1] <= 25:
            return 8  # fallback
        for i in range(min(30, len(raw_df))):
            val = str(raw_df.iloc[i, 25]).strip()
            if val.upper().startswith("SM-"):
                detected = max(0, i - 1)
                print(f"[Q-data] Z열 SM- 감지: 엑셀 {i + 1}행, 헤더행={detected + 1}")
                return detected
        print("[Q-data] Z열 SM- 미감지, 기본 헤더행=9 사용")
        return 8  # fallback

    suffix_map = {"openpyxl": ".xlsx", "xlrd": ".xls", "pyxlsb": ".xlsb", None: ""}

    for engine in engines:
        eng_kwargs = {"engine": engine} if engine else {}

        # 메모리 방식
        try:
            raw_df = pd.read_excel(io.BytesIO(content), header=None, **eng_kwargs)
            header_row = _detect_header_row(raw_df)
            df = pd.read_excel(io.BytesIO(content), header=header_row, usecols=usecols, **eng_kwargs)
            if df is not None and not df.empty:
                print(f"Q-data DRM 성공: {engine or '기본'} (메모리)")
                return _rename_qdata_columns(df)
        except Exception as e:
            last_error = str(e)

        # 임시 파일 방식
        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix_map.get(engine, "")) as tmp:
                tmp.write(content)
                temp_path = tmp.name
            raw_df = pd.read_excel(temp_path, header=None, **eng_kwargs)
            header_row = _detect_header_row(raw_df)
            df = pd.read_excel(temp_path, header=header_row, usecols=usecols, **eng_kwargs)
            if df is not None and not df.empty:
                print(f"Q-data DRM 성공: {engine or '기본'} (임시 파일)")
                return _rename_qdata_columns(df)
        except Exception as e:
            last_error = str(e)
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

    raise Exception(f"Q-data 엑셀 파일 읽기 실패: {last_error}")


def _rename_qdata_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [
        "service_date",
        "j_category",
        "process_type",
        "repair_name",
        "repair_detail",
        "detail_content",
        "model_name",
        "serial_number",
        "log_id",
        "sw_before",
        "sw_after",
    ]
    return df
