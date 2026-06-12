"""
auto_synthesize_gui_dpv2.py
v2: 다중 엑셀 파일 일괄 마스킹 + 컬럼 관계 자동 추천 + 사번 기반 MD 생성

[v1 대비 추가 기능]
1. 엑셀 파일 여러 개 추가 → 한 번에 마스킹
2. 파일 간 컬럼 관계 자동 감지 + 추천 (이름 유사도 + 값 중복도)
   → 관계가 있는 컬럼은 마스킹 값을 공유하여 관계 유지
3. 사번 입력 → 클로드코드용 테이블 생성 작업서(.md) 자동 생성
   - DBconn() 사용 패턴
   - 테이블명: <사번>_<엑셀파일명>
   - 컬럼 간 FK 관계 명시

[자립형 구조]
- 본 파일은 자립형(self-contained) — auto_synthesize_gui_dp.py 없어도 단독 실행
- v1 파일은 그대로 보존되며 별도 GUI로 계속 동작
- v1에서 사용하던 합성/검증 유틸리티는 모두 본 파일에 인라인되어 있음

[실행]
    python auto_synthesize_gui_dpv2.py
"""

import os, sys, json, warnings, threading, re
from datetime import datetime
from collections import defaultdict
from difflib import SequenceMatcher

import numpy as np
import pandas as pd
from scipy import stats   # 인라인된 v1 유틸(validate_quality, Gaussian Copula 등)에서 사용
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

warnings.filterwarnings('ignore')

# ──────────────────────────────────────────────────────────────
# 옵셔널 외부 라이브러리 — PIL / xlwings / openpyxl 감지
# ──────────────────────────────────────────────────────────────

# PIL — 로고 이미지 리사이즈용 (없으면 로고 생략)
HAS_PIL = False
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    pass

# xlwings는 선택사항 — 없거나 Excel 미설치면 openpyxl로 폴백
HAS_XLWINGS = False
try:
    import xlwings as xw
    HAS_XLWINGS = True
except (ImportError, OSError, Exception):
    pass

try:
    import openpyxl
except ImportError:
    if not HAS_XLWINGS:
        raise ImportError(
            "xlwings 또는 openpyxl 중 하나가 필요합니다.\n"
            "  pip install openpyxl   또는   pip install xlwings")


def _resource_path(filename):
    """PyInstaller EXE에서도 동작하는 리소스 경로 반환."""
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


# ──────────────────────────────────────────────────────────────
# DB 접속 (DB_app.py 패턴 이식) — 테이블명 중복 검사용
# ──────────────────────────────────────────────────────────────

HAS_DB = False
try:
    import pyodbc
    try:
        from dotenv import load_dotenv
        load_dotenv(_resource_path(".env"))
    except ImportError:
        # dotenv가 없어도 환경변수로 동작 가능
        pass
    HAS_DB = True
except ImportError:
    pyodbc = None


def _db_connect(timeout=3):
    """MSSQL 연결 — DB_app.py와 동일한 환경변수 사용. 실패 시 None."""
    if not HAS_DB:
        return None
    server = os.getenv("DB_SERVER")
    port = os.getenv("DB_PORT")
    database = os.getenv("DB_NAME")
    user = os.getenv("DB_USER")
    pw = os.getenv("DB_PASSWORD")
    driver = os.getenv("DB_DRIVER", "ODBC Driver 17 for SQL Server")
    if not all([server, database, user, pw]):
        return None
    conn_str = (
        f"DRIVER={{{driver}}};"
        f"SERVER={server},{port};"
        f"DATABASE={database};"
        f"UID={user};"
        f"PWD={pw};"
        "TrustServerCertificate=yes;"
    )
    try:
        return pyodbc.connect(conn_str, timeout=timeout)
    except Exception:
        return None


def _db_table_exists(full_table_name, conn=None):
    """주어진 테이블명이 DB에 존재하는지 여부.

    Returns:
        True/False — 존재/비존재
        None — DB 연결 실패 (검사를 건너뜀)
    """
    own_conn = False
    if conn is None:
        conn = _db_connect()
        own_conn = True
    if conn is None:
        return None
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT 1 FROM INFORMATION_SCHEMA.TABLES "
            "WHERE TABLE_NAME = ?", full_table_name)
        return cur.fetchone() is not None
    except Exception:
        return None
    finally:
        if own_conn:
            try:
                conn.close()
            except Exception:
                pass


# ══════════════════════════════════════════════════════════════
# v1에서 인라인된 핵심 유틸리티 함수들 (자립형 단일 파일 구성)
#  - 원본 위치: auto_synthesize_gui_dp.py L73~L1218
#  - v1 GUI 클래스(SynthesizeApp)는 인라인하지 않음
# ══════════════════════════════════════════════════════════════

# ── 문자열 동적 생성기 (자동 채우기용) ────────────────────────
_LAST_NAMES  = ['김','이','박','최','정','강','윤','임','한','오',
                '신','홍','문','류','배','전','조','남','서','권']
_FIRST_PARTS = ['민','서','지','수','도','하','준','소','예','태',
                '재','채','나','기','윤','성','우','세','진','아']
_LAST_PARTS  = ['준','연','호','아','서','윤','혁','율','훈','린',
                '민','원','영','현','양','나','재','진','수','은']
_CO_PREFIX = ['Alpha','Beta','Gamma','Delta','Epsilon','Zeta','Eta',
              'Theta','Iota','Kappa','Lambda','Mu','Nu','Xi','Omicron',
              'Pi','Rho','Sigma','Tau','Upsilon','Nova','Apex','Nexus',
              'Prime','Vertex','Zenith','Orion','Titan','Vega','Polaris']
_CO_SUFFIX = ['Corp','Group','Co','Inc','Ltd','Partners','Holdings',
              'Solutions','Systems','Global','Industries','Ventures',
              'Marine','Logistics','Energy','Shipping','Tech','Works']


def generate_fake_persons(n, seed=42):
    """가짜 한글 이름 n개 생성 — 중복 없음."""
    rng, names, result, attempt = np.random.default_rng(seed), set(), [], 0
    while len(result) < n:
        name = (_LAST_NAMES[rng.integers(len(_LAST_NAMES))]
              + _FIRST_PARTS[rng.integers(len(_FIRST_PARTS))]
              + _LAST_PARTS[rng.integers(len(_LAST_PARTS))])
        if name not in names:
            names.add(name); result.append(name)
        attempt += 1
        if attempt > n * 200:
            result.append(f'직원_{len(result):04d}')
    return result


def generate_fake_companies(n, seed=42):
    """가짜 회사명 n개 생성 — 중복 없음."""
    rng, names, result, attempt = np.random.default_rng(seed), set(), [], 0
    while len(result) < n:
        name = f'{_CO_PREFIX[rng.integers(len(_CO_PREFIX))]}-{_CO_SUFFIX[rng.integers(len(_CO_SUFFIX))]}'
        if name not in names:
            names.add(name); result.append(name)
        attempt += 1
        if attempt > n * 200:
            result.append(f'Company-{len(result):04d}')
    return result


def generate_auto_codes(col_name, n, seed=42):
    """컬럼명 기반 자동 코드 생성 (예: 'col_A', 'col_B', ...)."""
    cs = str(col_name)
    alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    return [f'{cs}_{alpha[i] if i<26 else str(i)}' for i in range(n)]


# ── 한글 포함 날짜 파싱 유틸리티 ──────────────────────────────

def _parse_korean_date(val):
    """한글이 포함된 날짜 문자열을 파싱한다.

    예: '2024년 3월 15일', '2024년03월15일', '2025-11-17 오전 12:00:00' 등
    """
    if not isinstance(val, str):
        return val
    s = val.strip()

    # 오전/오후 (Korean AM/PM) 처리
    m_ampm = re.match(
        r'(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})\s*(오전|오후)\s*(\d{1,2})[:\s](\d{2})(?:[:\s](\d{2}))?',
        s
    )
    if m_ampm:
        date_part = m_ampm.group(1)
        ampm = m_ampm.group(2)
        h, mi = int(m_ampm.group(3)), int(m_ampm.group(4))
        sc = int(m_ampm.group(5)) if m_ampm.group(5) else 0
        if ampm == '오후' and h < 12:
            h += 12
        elif ampm == '오전' and h == 12:
            h = 0
        try:
            base = pd.to_datetime(date_part)
            return base.replace(hour=h, minute=mi, second=sc)
        except Exception:
            pass

    # 일반 datetime+시간 혼합
    m_dt = re.match(
        r'(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})\s+(\d{1,2}:\d{2}(?::\d{2})?)', s
    )
    if m_dt:
        try:
            return pd.to_datetime(s)
        except Exception:
            pass

    # 한글 날짜 패턴
    m = re.match(
        r'(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일'
        r'(?:\s*(\d{1,2})\s*시\s*(\d{1,2})\s*분(?:\s*(\d{1,2})\s*초)?)?',
        s
    )
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        h = int(m.group(4)) if m.group(4) else 0
        mi = int(m.group(5)) if m.group(5) else 0
        sc = int(m.group(6)) if m.group(6) else 0
        try:
            return datetime(y, mo, d, h, mi, sc)
        except ValueError:
            return pd.NaT
    # 한글만 제거 후 재시도
    cleaned = re.sub(r'[년월일시분초]', ' ', s).strip()
    cleaned = re.sub(r'\s+', ' ', cleaned)
    if cleaned != s:
        try:
            return pd.to_datetime(cleaned)
        except Exception:
            pass
    return None


def _clean_datetime_column(series):
    """datetime 컬럼을 순수 Date 또는 Time 문자열로 정리.

    - 모든 시간이 00:00:00 → date only ('YYYY-MM-DD')
    - 모든 날짜가 동일 → time only ('HH:MM:SS')
    - 그 외 → date only (시간 제거)
    """
    if not pd.api.types.is_datetime64_any_dtype(series):
        return series
    valid = series.dropna()
    if len(valid) == 0:
        return series
    times = valid.dt.time
    dates = valid.dt.date
    all_midnight = (times == pd.Timestamp('00:00:00').time()).all()
    unique_dates = dates.nunique()
    if unique_dates <= 1 and not all_midnight:
        return series.dt.strftime('%H:%M:%S').where(series.notna(), other=np.nan)
    return series.dt.strftime('%Y-%m-%d').where(series.notna(), other=np.nan)


def _safe_to_datetime(series):
    """한글 날짜를 포함할 수 있는 시리즈를 안전하게 datetime으로 변환. null 보존."""
    result = series.copy()
    for idx, val in series.items():
        if pd.isna(val) or val is None:
            result.at[idx] = pd.NaT
            continue
        if isinstance(val, datetime):
            continue
        parsed = _parse_korean_date(val)
        if parsed is not None:
            result.at[idx] = parsed
        else:
            try:
                result.at[idx] = pd.to_datetime(val)
            except Exception:
                result.at[idx] = pd.NaT
    return pd.to_datetime(result, errors='coerce')


# ── 데이터프레임 후처리 (컬럼별 타입 추론) ────────────────────

def _postprocess_dataframe(df):
    """로드된 DataFrame에 대해 컬럼별 타입 추론 (한글 날짜, 숫자 등). null 보존."""
    for col in df.columns:
        s = df[col]
        # datetime 객체 감지
        if s.dtype == object:
            smp = s.dropna().head(10)
            if len(smp) and smp.apply(lambda x: isinstance(x, datetime)).all():
                df[col] = pd.to_datetime(s, errors='coerce')
                continue
        # 오전/오후 패턴 감지
        if s.dtype == object:
            smp_str = s.dropna().head(20).astype(str)
            has_ampm = smp_str.str.contains(r'오전|오후', na=False).mean() > 0.3
            if has_ampm:
                converted = _safe_to_datetime(s)
                if converted.notna().mean() > 0.5:
                    df[col] = converted
                    continue
        # 한글 날짜 패턴 감지
        if s.dtype == object:
            smp_str = s.dropna().head(20).astype(str)
            has_korean_date = smp_str.str.contains(r'\d+\s*년', na=False).mean() > 0.5
            if has_korean_date:
                converted = _safe_to_datetime(s)
                if converted.notna().mean() > 0.5:
                    df[col] = converted
                    continue
        # 일반 날짜 문자열 — 먼저 앞쪽 50개 샘플로 날짜 여부 선검사 후
        # 날짜로 보일 때만 전체 변환 (날짜 아닌 대용량 텍스트에 전체 파싱하던 비용 제거)
        if s.dtype == object:
            try:
                smp = s.dropna().head(50)
                if len(smp) and pd.to_datetime(
                        smp, errors='coerce').notna().mean() > 0.8:
                    conv = pd.to_datetime(s, errors='coerce')
                    if conv.notna().mean() > 0.8:
                        df[col] = conv
                        continue
            except Exception:
                pass
        # 숫자 변환 (콤마 제거)
        if s.dtype == object:
            try:
                num_conv = pd.to_numeric(
                    s.astype(str).str.replace(',', '', regex=False).str.strip(),
                    errors='coerce'
                )
                if num_conv.notna().mean() >= 0.5:
                    df[col] = num_conv
            except Exception:
                pass
    df.dropna(how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


# ── Excel 로드 (xlwings 우선, openpyxl 폴백) ──────────────────

def _get_sheet_names_xlwings(excel_path):
    """xlwings로 시트 이름 목록 반환."""
    app, wb = None, None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        wb = app.books.open(os.path.abspath(excel_path))
        return [s.name for s in wb.sheets]
    finally:
        if wb:  wb.close()
        if app: app.quit()


def _load_excel_xlwings(excel_path, sheet_name=None):
    """xlwings로 Excel 로드 (DRM 해제 시 사용 가능)."""
    app, wb, close_after = None, None, False
    try:
        abs_path = os.path.abspath(excel_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"파일 없음: {abs_path}")
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        wb = app.books.open(abs_path)
        close_after = True
        ws = wb.sheets[sheet_name] if sheet_name else wb.sheets[0]
        raw_data = ws.used_range.value
        if raw_data is None:
            raise ValueError(f"시트 '{ws.name}' 데이터 없음")
        if not isinstance(raw_data[0], list):
            raw_data = [raw_data]
        headers = [h if h is not None else f'col_{i}'
                   for i, h in enumerate(raw_data[0])]
        df = pd.DataFrame(raw_data[1:], columns=headers)
        df = _postprocess_dataframe(df)
        info = {
            'file_name': wb.name, 'file_path': excel_path or wb.fullname,
            'sheet_name': ws.name, 'all_sheets': [s.name for s in wb.sheets],
            'rows': len(df), 'cols': len(df.columns),
        }
        return df, info
    finally:
        if close_after:
            if wb:  wb.close()
            if app: app.quit()


def _get_sheet_names_openpyxl(excel_path):
    """openpyxl로 시트 이름 읽기 (read_only 모드)."""
    abs_path = os.path.abspath(excel_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"파일 없음: {abs_path}")
    wb = openpyxl.load_workbook(abs_path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def _load_excel_openpyxl(excel_path, sheet_name=None):
    """openpyxl + pandas로 Excel 로드."""
    abs_path = os.path.abspath(excel_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"파일 없음: {abs_path}")
    file_name = os.path.basename(abs_path)
    xls = pd.ExcelFile(abs_path, engine='openpyxl')
    all_sheets = xls.sheet_names
    target_sheet = sheet_name if sheet_name else all_sheets[0]
    df = pd.read_excel(xls, sheet_name=target_sheet, engine='openpyxl')
    xls.close()
    df = _postprocess_dataframe(df)
    info = {
        'file_name': file_name, 'file_path': abs_path,
        'sheet_name': target_sheet, 'all_sheets': all_sheets,
        'rows': len(df), 'cols': len(df.columns),
    }
    return df, info


def get_sheet_names(excel_path):
    """시트 이름 목록 반환. xlwings → openpyxl 순서로 시도."""
    errors = []
    if HAS_XLWINGS:
        try:
            return _get_sheet_names_xlwings(excel_path), 'xlwings'
        except Exception as e:
            errors.append(f"xlwings: {e}")
    try:
        return _get_sheet_names_openpyxl(excel_path), 'openpyxl'
    except Exception as e:
        errors.append(f"openpyxl: {e}")
    raise RuntimeError("시트 목록 로드 실패:\n" + "\n".join(errors))


def load_excel(excel_path, sheet_name=None):
    """Excel 로드. xlwings → openpyxl 순서로 시도."""
    errors = []
    if HAS_XLWINGS:
        try:
            return _load_excel_xlwings(excel_path, sheet_name), 'xlwings'
        except Exception as e:
            errors.append(f"xlwings: {e}")
    try:
        return _load_excel_openpyxl(excel_path, sheet_name), 'openpyxl'
    except Exception as e:
        errors.append(f"openpyxl: {e}")
    raise RuntimeError("파일 로드 실패:\n" + "\n".join(errors))


# ── 컬럼 타입 / PII / ID 감지 ─────────────────────────────────

def auto_detect_column_type(series, sample_size=1000):
    """series의 데이터 타입을 'numerical'/'datetime'/'categorical' 중 하나로 추론.

    정규식 판별은 앞쪽 sample_size개 샘플로만 수행 (대용량 성능)."""
    sample = series.dropna()
    if len(sample) == 0:
        return 'categorical'
    if pd.api.types.is_datetime64_any_dtype(series):
        return 'datetime'
    if series.dtype == object:
        smp = sample if len(sample) <= sample_size else sample.head(sample_size)
        smp_str = smp.astype(str)
        if smp_str.str.contains(r'\d+\s*년\s*\d+\s*월', na=False).mean() > 0.5:
            return 'datetime'
        if smp_str.str.match(r'\d{4}[-/]\d{2}[-/]\d{2}').mean() > 0.8:
            return 'datetime'
    if pd.api.types.is_numeric_dtype(series):
        return 'numerical'
    if series.dtype == object:
        if series.nunique() / len(series) < 0.15:
            return 'categorical'
    return 'categorical'


def _is_id_col(series, sample_size=1000):
    """ID 패턴(영문/한글 1-3자 + 선택 구분자 + 3자리 이상 숫자) 매칭율 > 0.7.

    대용량 컬럼은 앞쪽 sample_size개만으로 판별 (성능)."""
    s = series.dropna()
    if len(s) == 0:
        return False
    if len(s) > sample_size:
        s = s.head(sample_size)
    return s.astype(str).str.match(
        r'^[A-Za-z가-힣]{1,3}[\-_]?\d{3,}$').mean() > 0.7


# PII 자동 감지 패턴 (컬럼명 + 값 정규식)
_PII_NAME_PATTERNS = [
    r'이름|성명|성함|name(?!_id)',
    r'담당자?|책임자?|관리자|작성자|등록자|매니저|manager|owner',
    r'\bPM\b|프로젝트매니저|project[_\s]?manager',
    r'회사|법인|기업|업체|company|corp',
    r'전화|연락처|휴대폰|핸드폰|phone|tel|mobile',
    r'주민|ssn|resident',
    r'이메일|메일|email',
    r'주소|address',
    r'생년월일|birth',
]
_PII_VALUE_PATTERNS = [
    re.compile(r'^\d{6}[-\s]?\d{7}$'),                                # 주민번호
    re.compile(r'^01[016789][-\s]?\d{3,4}[-\s]?\d{4}$'),              # 휴대폰
    re.compile(r'^[\w.+-]+@[\w-]+\.[\w.-]+$'),                        # 이메일
]
_KOREAN_SURNAMES = (
    '김|이|박|최|정|강|조|윤|장|임|한|오|서|신|권|황|안|송|류|전|'
    '홍|고|문|양|손|배|백|허|유|남|심|노|하|곽|성|차|주|우|구|민'
)
_KOREAN_NAME_PATTERN = re.compile(
    rf'^({_KOREAN_SURNAMES})[가-힣]{{1,2}}$'
)


def detect_pii_columns(df, col_types):
    """컬럼명 패턴 + 값 정규식 + 한글 인명 휴리스틱으로 PII 컬럼 자동 감지."""
    pii = set()
    if df is None or len(df) == 0:
        return pii
    name_re = re.compile('|'.join(_PII_NAME_PATTERNS), re.IGNORECASE)
    for col in df.columns:
        if name_re.search(str(col)):
            pii.add(col)
            continue
        if col_types.get(col) != 'categorical':
            continue
        sample = df[col].dropna().astype(str).head(50)
        if len(sample) == 0:
            continue
        matched_pii = False
        for pat in _PII_VALUE_PATTERNS:
            if sample.str.match(pat).mean() > 0.5:
                pii.add(col)
                matched_pii = True
                break
        if matched_pii:
            continue
        if sample.str.match(_KOREAN_NAME_PATTERN).mean() > 0.5:
            pii.add(col)
    return pii


def detect_id_columns(df, col_types):
    """비즈니스 식별자(ID) 컬럼 감지 — PII와 분리 (GUI에서 🆔 배지로 표시)."""
    ids = set()
    if df is None or len(df) == 0:
        return ids
    for col in df.columns:
        if col_types.get(col) != 'categorical':
            continue
        try:
            if _is_id_col(df[col]):
                ids.add(col)
        except Exception:
            pass
    return ids


def _extract_id_prefix(values):
    """ID 값 리스트에서 공통 접두사(영문/한글) 추출. 없으면 빈 문자열."""
    if not values:
        return ''
    m = re.match(r'^([A-Za-z가-힣]+)', str(values[0]))
    return m.group(1) if m else ''


_ID_PREFIX_CANDIDATES = ['X', 'F', 'Z', 'Q', 'W', 'J', 'V', 'P', 'R', 'T']


def _pick_different_prefix(orig_prefix):
    """원본 접두사와 겹치지 않는 가짜 접두사 선택."""
    upper = orig_prefix.upper()
    for candidate in _ID_PREFIX_CANDIDATES:
        if candidate != upper:
            return candidate
    return 'FID'


# ── 문자열 합성 (매핑 기반) — null 보존 ───────────────────────

def synthesize_text_columns(df, col_types, mapping_dict):
    """mapping_dict ({컬럼: {원본값: 가짜값}})를 df에 적용한 새 DataFrame 반환. null 보존."""
    syn, desc_map = df.copy(), {}
    str_cols = [c for c in df.columns
                if (df[c].dtype == object or pd.api.types.is_string_dtype(df[c]))
                and col_types.get(c) == 'categorical']
    for col in str_cols:
        if col not in mapping_dict:
            continue
        mapping = mapping_dict[col]
        if not mapping:
            continue
        null_mask = df[col].isna()
        mapped = df[col].astype(str).map(mapping)
        mapped = mapped.where(mapped.notna(), df[col])
        mapped = mapped.where(~null_mask, other=np.nan)
        syn[col] = mapped
        desc_map[col] = {'method': '사용자 매핑', 'mapping': mapping}
    return syn, desc_map


# ── 카테고리 마스킹 (k-익명성 + 빈도 보존 합성) ───────────────

def synthesize_categorical_masked(df_original, df_text, col_types, col_modes,
                                  pii_cols, k_anon):
    """범주형 컬럼을 모드별로 후처리.

    Args:
        df_original : 원본 DataFrame (빈도 계산용)
        df_text     : synthesize_text_columns 결과 (매핑 적용된 상태)
        col_types   : {col: 'numerical'/'datetime'/'categorical'}
        col_modes   : {col: 'masked'/'keep'} — 사용자 선택
        pii_cols    : set — PII/관계 그룹 등 재샘플링 스킵 대상
        k_anon      : 빈도 < k_anon 카테고리 → 'OTHER' 통합

    Returns:
        (df_result, exposed_categories, other_counts)
    """
    syn = df_text.copy()
    exposed = {}
    other_counts = {}
    cat_cols = [c for c, t in col_types.items()
                if t == 'categorical' and c in syn.columns]
    if not cat_cols:
        return syn, exposed, other_counts
    n = len(syn)
    for col in cat_cols:
        if col in pii_cols:
            continue
        mode = col_modes.get(col, 'masked')
        if mode == 'keep':
            vals = df_original[col].dropna()
            if len(vals) == 0:
                continue
            null_mask = syn[col].isna()
            sampled = np.random.choice(vals.values, size=n, replace=True)
            syn[col] = sampled
            syn.loc[null_mask, col] = np.nan
            exposed[col] = sorted(set(map(str, vals.unique())))
            continue
        # mode == 'masked' — k-익명성 + 빈도 보존 합성
        masked_vals = df_text[col].dropna().astype(str)
        if len(masked_vals) == 0:
            continue
        freq = masked_vals.value_counts()
        rare_mask = freq < int(k_anon)
        n_rare = int(rare_mask.sum())
        other_counts[col] = n_rare
        if n_rare > 0:
            kept_freq = freq[~rare_mask].copy()
            kept_freq['OTHER'] = int(freq[rare_mask].sum())
            freq = kept_freq
        categories = freq.index.tolist()
        counts = freq.values.astype(float)
        total = counts.sum()
        probs = counts / total if total > 0 else None
        if probs is not None:
            sampled = np.random.choice(categories, size=n, replace=True, p=probs)
        else:
            sampled = np.random.choice(categories, size=n, replace=True)
        null_mask = syn[col].isna()
        syn[col] = sampled
        syn.loc[null_mask, col] = np.nan
        exposed[col] = sorted(set(map(str, categories)))
    return syn, exposed, other_counts


# ── 상관관계 / 제약조건 ───────────────────────────────────────

def analyze_correlations(df, col_types):
    """수치 컬럼들의 상관 행렬 + |r|>=0.5 강한 쌍 리스트 반환."""
    num_cols = [c for c, t in col_types.items() if t == 'numerical']
    if len(num_cols) < 2:
        return {}
    corr = df[num_cols].corr()
    pairs = []
    for i, c1 in enumerate(num_cols):
        for c2 in num_cols[i+1:]:
            r = corr.loc[c1, c2]
            if abs(r) >= 0.5:
                pairs.append({'col1': c1, 'col2': c2, 'r': round(r, 3)})
    return {'correlation_matrix': corr.round(3).to_dict(), 'strong_pairs': pairs}


def auto_detect_constraints(df, col_types):
    """제약조건 자동 감지: positive / range_0_100 / inequality(시작<종료)."""
    constraints = []
    num_cols  = [c for c, t in col_types.items() if t == 'numerical']
    date_cols = [c for c, t in col_types.items() if t == 'datetime']
    for col in num_cols:
        col_data = df[col].dropna()
        if len(col_data) == 0:
            continue
        if col_data.min() >= 0:
            constraints.append({'type': 'positive', 'column': col})
        if col_data.min() >= 0 and col_data.max() <= 100:
            constraints.append({'type': 'range_0_100', 'column': col})
    start_kw = ['시작', '착수', '계획', 'start', 'begin', 'from']
    end_kw   = ['종료', '완료', '끝', 'end', 'finish', 'to']
    s_cols = [c for c in date_cols if any(k in str(c) for k in start_kw)]
    e_cols = [c for c in date_cols if any(k in str(c) for k in end_kw)]
    for s in s_cols:
        for e in e_cols:
            try:
                mask = df[s].notna() & df[e].notna()
                if mask.sum() > 0 and (df.loc[mask, s] < df.loc[mask, e]).mean() > 0.9:
                    constraints.append({'type': 'inequality', 'low': s, 'high': e})
            except Exception:
                pass
    return constraints


# ── 마스킹 헬퍼 (v1 원본 — 컬럼 셔플 / 코드 생성) ────────────

_MASK_CHARS = 'ABCDEFGHIJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789'


def generate_masked_column_names(orig_columns, seed=None):
    """무작위 영문/숫자 코드(3자+_+2자)로 컬럼명 마스킹. 충돌 없음.

    v2에서는 파일 간 충돌까지 막는 generate_unique_masked_column_names를 사용하지만,
    호환을 위해 원본 v1 함수도 유지.
    """
    rng = np.random.default_rng(seed)
    chars = np.array(list(_MASK_CHARS))
    digits = np.array(list('0123456789'))
    used = set()
    mapping = {}
    for c in orig_columns:
        while True:
            prefix = ''.join(rng.choice(chars, 3))
            suffix = ''.join(rng.choice(digits, 2))
            code = f'{prefix}_{suffix}'
            if code not in used:
                used.add(code)
                break
        mapping[c] = code
    return mapping


def generate_masked_category_codes(unique_values, col_alpha):
    """컬럼별 prefix + 일련번호로 무의미 코드 생성. 예: '선행의장' → 'CAT_C_001'."""
    return {v: f'CAT_{col_alpha}_{i+1:03d}'
            for i, v in enumerate(unique_values)}


def shuffle_columns_safely(columns, df_original, col_types, seed=None,
                           strong_corr_threshold=0.5, max_attempts=200):
    """컬럼 순서 셔플 + 강한 상관 쌍이 인접하지 않도록 swap (max_attempts 회)."""
    rng = np.random.default_rng(seed)
    columns = list(columns)
    if len(columns) <= 2:
        return columns
    num_cols = [c for c in columns
                if c in df_original.columns and col_types.get(c) == 'numerical']
    strong_pairs = set()
    if len(num_cols) >= 2:
        try:
            corr = df_original[num_cols].corr().abs()
            for i, a in enumerate(num_cols):
                for b in num_cols[i+1:]:
                    if corr.loc[a, b] > strong_corr_threshold:
                        strong_pairs.add(frozenset([a, b]))
        except Exception:
            pass
    order = columns.copy()
    rng.shuffle(order)
    if not strong_pairs:
        return order
    for _ in range(max_attempts):
        adjacent = []
        for i in range(len(order) - 1):
            if frozenset([order[i], order[i+1]]) in strong_pairs:
                adjacent.append(i)
        if not adjacent:
            break
        idx = adjacent[0]
        far_candidates = [j for j in range(len(order))
                          if abs(j - idx) > 2 and abs(j - idx - 1) > 2]
        if not far_candidates:
            break
        target = int(rng.choice(far_candidates))
        order[idx], order[target] = order[target], order[idx]
    return order


# ── 수치/날짜 합성 (통계 보존 Gaussian Copula) ───────────────

def generate_numeric_datetime_masked(df, col_types, constraints, num_rows=None,
                                     n_quantiles=64):
    """비-DP 통계 보존 Gaussian Copula 합성.

    수치 + 단독 datetime을 모두 copula에 통합해 진짜 공분산·분위수 격자로 합성.
    노이즈 주입 없음 → 평균/분산/상관관계가 원본과 사실상 일치.
    외부 공유 안전성은 컬럼명/카테고리 마스킹 + 순서 셔플로 확보 (이 함수 밖).
    """
    n = num_rows or len(df)
    syn = pd.DataFrame(index=range(n))
    num_cols = [c for c, t in col_types.items() if t == 'numerical']
    date_cols = [c for c, t in col_types.items() if t == 'datetime']
    # null 비율 기록
    null_ratios = {}
    for col in num_cols + date_cols:
        total = len(df[col])
        null_count = df[col].isna().sum()
        null_ratios[col] = null_count / total if total > 0 else 0
    # datetime → epoch 변환 (copula 통합)
    ineq_low_set = {c['low'] for c in constraints if c['type'] == 'inequality'}
    ineq_high_set = {c['high'] for c in constraints if c['type'] == 'inequality'}
    dt_in_copula = [c for c in date_cols
                    if c not in ineq_low_set and c not in ineq_high_set
                    and df[c].notna().sum() > 0]
    if dt_in_copula:
        df = df.copy()
        for dcol in dt_in_copula:
            ts = pd.to_datetime(df[dcol], errors='coerce')
            na_mask = np.asarray(ts.isna().values)
            epoch = np.array(
                ts.astype('datetime64[s]').astype('int64').astype(float).values,
                dtype=float, copy=True)
            epoch[na_mask] = np.nan
            df[dcol] = epoch
        num_cols = num_cols + dt_in_copula
    pos_cols = {c['column'] for c in constraints if c['type'] == 'positive'}
    range_cols = {c['column'] for c in constraints if c['type'] == 'range_0_100'}
    # 수치 컬럼 — 통계 보존 Gaussian Copula
    if num_cols:
        col_data_map = {}
        for col in num_cols:
            vals = df[col].dropna().values.astype(float)
            if len(vals) > 0:
                col_data_map[col] = vals
        valid_num_cols = [c for c in num_cols if c in col_data_map]
        if len(valid_num_cols) > 1:
            sub = df[valid_num_cols].dropna()
            if len(sub) > 1:
                d = len(valid_num_cols)
                normals = np.zeros((len(sub), d))
                params = {}
                for i, col in enumerate(valid_num_cols):
                    vals = sub[col].values.astype(float)
                    normals[:, i] = stats.norm.ppf(
                        stats.rankdata(vals) / (len(vals) + 1))
                    sorted_vals = np.sort(col_data_map[col])
                    if col in pos_cols:
                        sorted_vals = sorted_vals[sorted_vals >= 0]
                    if col in range_cols:
                        sorted_vals = sorted_vals[(sorted_vals >= 0) & (sorted_vals <= 100)]
                    if len(sorted_vals) == 0:
                        sorted_vals = np.sort(col_data_map[col])
                    params[col] = sorted_vals
                cov = np.cov(normals.T) + np.eye(d) * 1e-6
                sample = np.random.multivariate_normal(np.zeros(d), cov, n)
                unif = stats.norm.cdf(sample)
                grid_levels = np.linspace(0.0, 1.0, n_quantiles)
                for i, col in enumerate(valid_num_cols):
                    grid = np.quantile(params[col], grid_levels)
                    syn[col] = np.interp(unif[:, i], grid_levels, grid)
                    if df[col].dtype in [np.int32, np.int64]:
                        syn[col] = syn[col].round().astype(int)
            elif len(valid_num_cols) == 1:
                col = valid_num_cols[0]
                sorted_vals = np.sort(col_data_map[col])
                idx = np.random.randint(0, len(sorted_vals), n)
                syn[col] = sorted_vals[idx]
        elif len(valid_num_cols) == 1:
            col = valid_num_cols[0]
            sorted_vals = np.sort(col_data_map[col])
            idx = np.random.randint(0, len(sorted_vals), n)
            syn[col] = sorted_vals[idx]
    # inequality 쌍 datetime — 별도 처리 (sampling)
    ineq = {c['low']: c['high'] for c in constraints if c['type'] == 'inequality'}
    for col in date_cols:
        if col in ineq.values():
            continue
        if col in dt_in_copula:
            continue
        ts = df[col].dropna()
        if len(ts) == 0:
            continue
        ts_vals = pd.to_datetime(ts).values.astype('datetime64[s]').astype(np.int64)
        sampled = np.random.choice(ts_vals, size=n)
        if col in ineq:
            end_col = ineq[col]
            mask = df[col].notna() & df[end_col].notna()
            dur = (pd.to_datetime(df.loc[mask, end_col]) -
                   pd.to_datetime(df.loc[mask, col])).dt.days.values
            if len(dur) == 0:
                dur = np.array([0])
            dur = np.abs(dur)
            syn[col] = pd.to_datetime(sampled, unit='s')
            syn[end_col] = syn[col] + pd.to_timedelta(
                np.random.choice(dur, n), unit='D'
            )
        else:
            syn[col] = pd.to_datetime(sampled, unit='s')
    # 제약조건 재분배
    for c in constraints:
        col = c.get('column')
        if col is None or col not in syn.columns:
            continue
        col_vals = np.array(syn[col].values, dtype=float, copy=True)
        non_null = ~np.isnan(col_vals)
        if c['type'] == 'positive':
            violations = non_null & (col_vals < 0)
            if violations.any():
                valid = col_vals[non_null & (col_vals >= 0)]
                if len(valid) > 0:
                    lo = valid.min()
                    hi = np.percentile(valid, 25)
                    if lo == hi:
                        hi = lo + 1
                    col_vals[violations] = np.random.uniform(lo, hi, violations.sum())
                else:
                    col_vals[violations] = np.abs(col_vals[violations])
                syn[col] = col_vals
                if df[col].dtype in [np.int32, np.int64]:
                    syn[col] = syn[col].round().astype(int)
        elif c['type'] == 'range_0_100':
            violations = non_null & ((col_vals < 0) | (col_vals > 100))
            if violations.any():
                valid = col_vals[non_null & (col_vals >= 0) & (col_vals <= 100)]
                if len(valid) > 0:
                    col_vals[violations] = np.random.choice(valid, violations.sum())
                else:
                    col_vals[violations] = np.clip(col_vals[violations], 0, 100)
                syn[col] = col_vals
                if df[col].dtype in [np.int32, np.int64]:
                    syn[col] = syn[col].round().astype(int)
    # null 비율 복원
    for col in syn.columns:
        if col in null_ratios and null_ratios[col] > 0:
            null_count = int(round(n * null_ratios[col]))
            if null_count > 0:
                null_indices = np.random.choice(n, size=null_count, replace=False)
                syn.loc[null_indices, col] = np.nan
    # epoch → datetime 역변환
    for dcol in dt_in_copula:
        if dcol in syn.columns:
            epoch_vals = pd.to_numeric(syn[dcol], errors='coerce')
            syn[dcol] = pd.to_datetime(epoch_vals, unit='s', errors='coerce')
    return syn


# ── 함수 종속성 / 품질 검증 ───────────────────────────────────

def detect_functional_dependencies(df, col_types):
    """범주형 컬럼 간 함수 종속성 감지.

    A → B: A의 각 고유값이 항상 같은 B값에 매핑되면 A가 B를 결정.
    """
    cat_cols = [c for c, t in col_types.items() if t == 'categorical']
    deps = []
    for a in cat_cols:
        for b in cat_cols:
            if a == b:
                continue
            sub = df[[a, b]].dropna()
            if len(sub) < 2:
                continue
            grouped = sub.groupby(a)[b].nunique()
            if (grouped == 1).all():
                deps.append({'from': a, 'to': b})
    return deps


def validate_quality(real, synthetic, col_types):
    """원본 vs 합성 데이터의 품질 점수 (컬럼별 + 종합 평균)."""
    scores = {}
    for col, ctype in col_types.items():
        if col not in real.columns or col not in synthetic.columns:
            continue
        try:
            if ctype == 'numerical':
                r_clean = real[col].dropna()
                s_clean = synthetic[col].dropna()
                if len(r_clean) == 0 or len(s_clean) == 0:
                    scores[col] = 0.5
                    continue
                ks, _ = stats.ks_2samp(r_clean, s_clean)
                scores[col] = round(1 - ks, 3)
            elif ctype == 'categorical':
                r = real[col].value_counts(normalize=True)
                s = synthetic[col].value_counts(normalize=True)
                r_s = r.sort_values(ascending=False).values
                s_s = s.sort_values(ascending=False).values
                mn = min(len(r_s), len(s_s))
                scores[col] = round(1 - np.abs(r_s[:mn] - s_s[:mn]).sum() / 2, 3)
            elif ctype == 'datetime':
                r_clean = real[col].dropna()
                s_clean = synthetic[col].dropna()
                if len(r_clean) == 0 or len(s_clean) == 0:
                    scores[col] = 0.5
                    continue
                r_ts = r_clean.values.astype('datetime64[s]').astype(np.int64)
                s_ts = s_clean.values.astype('datetime64[s]').astype(np.int64)
                ks, _ = stats.ks_2samp(r_ts, s_ts)
                scores[col] = round(1 - ks, 3)
        except Exception:
            scores[col] = 0.5
    overall = round(float(np.mean(list(scores.values()))), 3) if scores else 0.0
    return overall, scores


# ══════════════════════════════════════════════════════════════
# 마스킹 컬럼명 — 파일 간 충돌 방지 버전 (v2 신규)
# ══════════════════════════════════════════════════════════════

_MASK_CHARS_V2 = 'ABCDEFGHIJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789'


def generate_unique_masked_column_names(orig_columns, excluded=None, seed=None):
    """v1의 generate_masked_column_names와 같은 형식이나,
    `excluded` 집합에 든 코드는 사용하지 않는다 (파일 간 중복 방지).

    Args:
        orig_columns: 원본 컬럼명 리스트
        excluded: 이미 다른 파일에서 사용 중인 마스킹 코드 집합
        seed: 난수 시드

    Returns:
        {원본명: 마스킹 코드} 매핑 (excluded와 충돌 없음)
    """
    rng = np.random.default_rng(seed)
    chars = np.array(list(_MASK_CHARS_V2))
    digits = np.array(list('0123456789'))
    used = set(excluded) if excluded else set()
    mapping = {}
    for c in orig_columns:
        while True:
            prefix = ''.join(rng.choice(chars, 3))
            suffix = ''.join(rng.choice(digits, 2))
            code = f'{prefix}_{suffix}'
            if code not in used:
                used.add(code)
                break
        mapping[c] = code
    return mapping


# ══════════════════════════════════════════════════════════════
# 컬럼 관계 자동 감지
# ══════════════════════════════════════════════════════════════

def _normalize_col_name(name):
    """비교용 정규화: 소문자화 + 공백/특수문자 제거."""
    s = str(name).lower()
    s = re.sub(r'[\s_\-/.()\[\]]+', '', s)
    return s


def compute_name_similarity(name1, name2):
    """컬럼명 유사도 (0~1)."""
    n1 = _normalize_col_name(name1)
    n2 = _normalize_col_name(name2)
    if not n1 or not n2:
        return 0.0
    if n1 == n2:
        return 1.0
    if n1 in n2 or n2 in n1:
        shorter = min(len(n1), len(n2))
        longer = max(len(n1), len(n2))
        return 0.7 + 0.25 * (shorter / longer)
    return SequenceMatcher(None, n1, n2).ratio()


# 관계 감지 시 Jaccard 추정에 사용할 고유값 최대 개수 (초대용량 컬럼 폭주 방지)
_OVERLAP_VALUE_CAP = 50000


def compute_value_overlap(values1, values2):
    """두 컬럼 값의 Jaccard 유사도 (0~1).

    파이썬 루프 대신 pandas 벡터화로 문자열 set을 만든다.
    고유값이 _OVERLAP_VALUE_CAP을 넘으면 '등간격(stride) 샘플링'으로 줄여
    값 분포를 보존한 채 추정한다 (앞부분만 자르면 겹침 구간을 놓치므로 금지).
    """
    def _to_set(vals):
        s = pd.Series(vals)
        s = s[s.notna()]
        if len(s) == 0:
            return set()
        uniq = pd.unique(s.astype(str).str.strip().to_numpy())
        if len(uniq) <= _OVERLAP_VALUE_CAP:
            return set(uniq)
        # 초대용량: 값 해시 기반 일관 샘플링 — 동일 값은 양쪽 컬럼에서
        # 똑같이 선택/탈락하므로 교집합(겹침)이 보존된다 (MinHash식 추정).
        m = len(uniq) // _OVERLAP_VALUE_CAP + 1
        h = pd.util.hash_pandas_object(
            pd.Series(uniq), index=False).to_numpy()
        return set(uniq[(h % m) == 0])
    s1 = _to_set(values1)
    s2 = _to_set(values2)
    if not s1 or not s2:
        return 0.0
    inter = len(s1 & s2)
    union = len(s1 | s2)
    return inter / union if union > 0 else 0.0


def detect_relationships(files_data,
                         name_threshold=0.5,
                         value_threshold=0.15,
                         confidence_threshold=0.45):
    """파일 간 컬럼 관계 자동 감지.

    Args:
        files_data: FileData 리스트
        name_threshold: 이름 유사도 최소값 (이보다 너무 낮으면 값 검사 생략)
        value_threshold: 값 중복 최소값
        confidence_threshold: 추천 신뢰도 임계값

    Returns:
        [{
            from_file_idx, from_col, to_file_idx, to_col,
            name_score, value_score, confidence, reason,
            enabled (bool, 추천 여부와 동일하게 초기화)
        }, ...]
    """
    relationships = []
    n_files = len(files_data)

    for i in range(n_files):
        f1 = files_data[i]
        if f1.df is None:
            continue
        for j in range(i + 1, n_files):
            f2 = files_data[j]
            if f2.df is None:
                continue

            for c1 in f1.df.columns:
                # datetime/numerical 컬럼은 관계 검사 대상에서 제외
                if f1.col_types.get(c1) not in ('categorical',):
                    # 단, 정수형 ID 같은 수치는 검사 — id_suspected에 들어있으면 포함
                    if c1 not in f1.id_suspected:
                        continue
                vals1_full = f1.df[c1].dropna()
                if len(vals1_full) == 0:
                    continue
                vals1 = vals1_full.unique()

                for c2 in f2.df.columns:
                    if f2.col_types.get(c2) not in ('categorical',):
                        if c2 not in f2.id_suspected:
                            continue
                    name_score = compute_name_similarity(c1, c2)
                    # 이름이 너무 다르면 값 검사도 안 하고 스킵
                    if name_score < name_threshold * 0.4:
                        # 단, ID 의심 컬럼끼리는 이름 무관하게 값으로 검사
                        if not (c1 in f1.id_suspected and c2 in f2.id_suspected):
                            continue

                    vals2_full = f2.df[c2].dropna()
                    if len(vals2_full) == 0:
                        continue
                    vals2 = vals2_full.unique()

                    value_score = compute_value_overlap(vals1, vals2)

                    id_bonus = 0.0
                    if c1 in f1.id_suspected or c2 in f2.id_suspected:
                        id_bonus = 0.1

                    confidence = 0.5 * name_score + 0.5 * value_score + id_bonus
                    confidence = min(1.0, confidence)

                    if confidence < confidence_threshold and value_score < value_threshold:
                        continue

                    reason_parts = []
                    if name_score >= 0.95:
                        reason_parts.append("이름 일치")
                    elif name_score >= 0.7:
                        reason_parts.append("이름 유사")
                    elif name_score >= name_threshold:
                        reason_parts.append("이름 일부 유사")
                    if value_score >= 0.5:
                        reason_parts.append(f"값 {value_score:.0%} 중복")
                    elif value_score >= value_threshold:
                        reason_parts.append(f"값 {value_score:.0%} 중복")
                    if id_bonus > 0:
                        reason_parts.append("ID 추정")
                    reason = " / ".join(reason_parts) or "낮은 유사도"

                    relationships.append({
                        'from_file_idx': i, 'from_col': str(c1),
                        'to_file_idx': j, 'to_col': str(c2),
                        'name_score': round(name_score, 3),
                        'value_score': round(value_score, 3),
                        'confidence': round(confidence, 3),
                        'reason': reason,
                        'enabled': confidence >= confidence_threshold,
                    })

    relationships.sort(key=lambda r: -r['confidence'])
    return relationships


def build_relationship_groups(relationships):
    """활성화된 관계를 union-find로 그룹화.

    Returns: [[(file_idx, col), ...], ...]
    """
    parent = {}

    def find(x):
        """Union-Find의 path-compression find 연산."""
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(x, y):
        """Union-Find의 union 연산."""
        rx, ry = find(x), find(y)
        if rx != ry:
            parent[rx] = ry

    for r in relationships:
        if not r.get('enabled'):
            continue
        a = (r['from_file_idx'], r['from_col'])
        b = (r['to_file_idx'], r['to_col'])
        if a not in parent:
            parent[a] = a
        if b not in parent:
            parent[b] = b
        union(a, b)

    groups_dict = defaultdict(list)
    for key in parent:
        root = find(key)
        groups_dict[root].append(key)
    return list(groups_dict.values())


# ══════════════════════════════════════════════════════════════
# MD 파일 생성 (클로드코드용 테이블 생성 작업서)
# ══════════════════════════════════════════════════════════════

def _safe_table_name(employee_id, file_name):
    """안전한 SQL 테이블명 생성: <사번>_<파일명에서 확장자 제거 + 정제>.

    file_name으로는 합성 데이터 파일명(예: '합성데이터_사원정보.xlsx')을 넘기는 것을
    권장. 원본 파일명을 넘기면 그것으로 테이블이 만들어짐.
    """
    base = os.path.splitext(file_name)[0]
    base = re.sub(r'[^\w가-힣]', '_', base)
    base = re.sub(r'_+', '_', base).strip('_')
    eid = re.sub(r'[^\w가-힣]', '_', str(employee_id))
    return f"{eid}_{base}"


def _table_name_for_file(employee_id, fd):
    """FileData에서 테이블명을 도출 — `<프로젝트명>_<파일별 테이블명>`.

    1순위: 사용자 지정 `fd.table_name` (영문 소문자)이 있으면 `<eid>_<table>`.
    2순위: 합성 데이터 파일명 기반 자동 생성 (구 로직).
    """
    user_table = (getattr(fd, 'table_name', '') or '').strip()
    eid = (employee_id or '').strip()
    if user_table:
        return f"{eid}_{user_table}" if eid else user_table
    name_source = (os.path.basename(fd.output_xlsx)
                   if getattr(fd, 'output_xlsx', None) else fd.file_name)
    return _safe_table_name(eid, name_source)


def _is_pk_candidate(fd, orig_col):
    """원본 데이터(fd.df) 기준 PK 적용 가능 여부 판정.

    합성 전 단계에서 dim 테이블 보호(replace=False) 결정에 사용.
    조건: NULL 없음 + 모든 값 고유 + 컬럼이 비어있지 않음.
    """
    if fd is None or fd.df is None or orig_col not in fd.df.columns:
        return False
    series = fd.df[orig_col]
    if len(series) == 0:
        return False
    if series.isna().any():
        return False
    if series.duplicated().any():
        return False
    return True


def _is_pk_candidate_synth(fd, orig_col):
    """합성된 데이터 기준 PK 적용 가능 여부 판정.

    MD 생성 시점에 사용. 원본이 unique여도 합성 후 중복이 있으면
    PK를 적용하지 않아야 INSERT 실패를 방지할 수 있다.
    """
    if fd is None or not getattr(fd, 'synth_unique_cols', None):
        return False
    masked = fd.column_mask_map.get(orig_col, orig_col)
    return str(masked) in fd.synth_unique_cols


def _pandas_dtype_to_sql(dtype, col_type):
    """pandas dtype + 추론 타입 → MSSQL (T-SQL) 타입.

    - 정수 → BIGINT
    - 실수 → FLOAT
    - datetime → DATETIME2 (정밀도 7자리, MSSQL 권장)
    - 그 외(범주/문자) → NVARCHAR(255) — 한글 안전을 위해 N-prefix
    """
    if col_type == 'datetime':
        return 'DATETIME2'
    if col_type == 'numerical':
        s = str(dtype)
        if 'bool' in s:
            return 'BIT'
        if 'int' in s:
            return 'BIGINT'
        return 'FLOAT'
    return 'NVARCHAR(255)'


def generate_table_creation_md(employee_id, files_data, relationships, save_path):
    """Claude CLI용 테이블 생성 지시서 MD 생성 — SQL 지시만 포함 (Python 래퍼 없음)."""
    lines = []
    L = lines.append

    table_names = {}
    for idx, fd in enumerate(files_data):
        # 합성 데이터 파일명(.xlsx) 기준으로 테이블명 생성 — 원본 파일명 미사용
        table_names[idx] = _table_name_for_file(employee_id, fd)

    active_rels = [r for r in relationships if r.get('enabled')]

    L(f"# 테이블 생성 지시서 (프로젝트명: `{employee_id}`)")
    L("")
    L(f"- 생성 시각: {datetime.now().isoformat(timespec='seconds')}")
    L(f"- 대상 테이블 수: **{len(files_data)}개**")
    L(f"- 활성 관계 수: **{len(active_rels)}개**")
    L(f"- **SQL Dialect: Microsoft SQL Server (T-SQL)**")
    L("")
    L("---")
    L("")
    L("## 📌 작업 지시")
    L("")
    L("**`data_processor.py`의 `dbconn()` 함수로 MSSQL에 연결한 뒤 아래 SQL들을 순서대로 실행**하여")
    L("테이블과 컬럼 관계(인덱스)를 생성하고, 합성 데이터(`.xlsx`)를 적재한다.")
    L("")
    L("> 🚀 **빠른 실행**: 본 MD 와 같은 폴더에 `tableanddata.py` 가 함께 생성됩니다. ")
    L("> `python tableanddata.py` 한 번 실행하면 아래의 모든 SQL + 데이터 INSERT 가 ")
    L("> 자동으로 처리됩니다 (테이블 생성 → INSERT → PK → FK/INDEX 순서). ")
    L("> 본 MD 의 SQL 은 검토/수동 실행/감사 목적의 참고용입니다.")
    L("")
    L("**수동 실행 시 패턴**")
    L("")
    L("```python")
    L("from data_processor import dbconn")
    L("")
    L("conn = dbconn()")
    L("cur = conn.cursor()")
    L("# 아래 SQL 블록을 순서대로 cur.execute(...) 로 실행")
    L("conn.commit()")
    L("conn.close()")
    L("```")
    L("")
    L("**T-SQL 작성 규칙**")
    L("")
    L("- 식별자 quote: 대괄호 `[name]`")
    L("- 문자열 컬럼: 한글 안전을 위해 `NVARCHAR(255)` (전부 NULL 허용)")
    L("- 정수: `BIGINT`,  실수: `FLOAT`,  날짜시간: `DATETIME2`,  불리언: `BIT`")
    L("- 테이블 존재 검사: `IF OBJECT_ID(N'테이블', N'U') IS NULL ... BEGIN ... END;`")
    L("- 인덱스 존재 검사: "
      "`IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name=N'idx' AND object_id=OBJECT_ID(N'테이블'))`")
    L("")
    L(f"**테이블명 규칙**: 파일별 사용자 지정 (영문 소문자) — 프로젝트명 `{employee_id}`")
    L("")
    L("### 🔑 PK / FK 자동 판정")
    L("")
    L("- 원본 데이터를 분석하여 **NULL 없음 + 모든 값 고유**인 컬럼은 `PRIMARY KEY` 적용 ")
    L("  대상으로 분류되며, 그 PK를 참조하는 쪽에는 `FOREIGN KEY` 제약이 추가됩니다.")
    L("- PK 적용이 불가능한(NULL이나 중복이 있는) 관계는 `CREATE INDEX`만으로 JOIN 성능을 ")
    L("  보장합니다 (FK 미적용 — 데이터 적재 단계에서 거부 방지).")
    L("- 본 문서에 명시된 SQL을 **순서대로 실행**하면 됩니다. ")
    L("  순서를 바꾸면 ALTER COLUMN/PK 추가가 실패할 수 있습니다 ")
    L("  (인덱스 있는 컬럼은 NOT NULL 변경 불가).")
    L("")
    L("---")
    L("")
    L("## 1. 테이블 정의")
    L("")
    L("**각 테이블에 대해 아래 SQL을 그대로 실행한다.**")
    L("")

    for idx, fd in enumerate(files_data):
        tname = table_names[idx]
        L(f"### 1.{idx+1}. `[{tname}]`")
        L("")
        if fd.output_xlsx:
            L(f"- 합성 데이터 파일: `{os.path.basename(fd.output_xlsx)}`")
        L(f"- 행 수: {len(fd.df) if fd.df is not None else '?'}")
        L("")
        L("**컬럼 (마스킹된 이름 / 원본명 비공개)**")
        L("")
        L("| # | 컬럼명 | 타입 | SQL 타입 | 설명 |")
        L("|---|---|---|---|---|")

        df = fd.df
        ct = fd.col_types
        descs = fd.col_descriptions or {}
        mask_map = fd.column_mask_map or {}

        if df is not None:
            for i, orig_col in enumerate(df.columns, 1):
                masked_col = mask_map.get(orig_col, orig_col)
                ctype = ct.get(orig_col,
                                ct.get(masked_col, 'categorical'))
                sql_type = _pandas_dtype_to_sql(df[orig_col].dtype, ctype)
                desc = descs.get(orig_col, '') or '_(설명 미입력)_'
                L(f"| {i} | `[{masked_col}]` | {ctype} | `{sql_type}` | {desc} |")
        L("")

        L("**CREATE TABLE SQL** — 그대로 실행")
        L("")
        L("```sql")
        if df is not None:
            L(f"IF OBJECT_ID(N'{tname}', N'U') IS NULL")
            L("BEGIN")
            L(f"    CREATE TABLE [{tname}] (")
            col_lines = []
            for orig_col in df.columns:
                masked_col = mask_map.get(orig_col, orig_col)
                ctype = ct.get(orig_col,
                                ct.get(masked_col, 'categorical'))
                sql_type = _pandas_dtype_to_sql(df[orig_col].dtype, ctype)
                col_lines.append(f"        [{masked_col}] {sql_type}")
            L(",\n".join(col_lines))
            L("    );")
            L("END;")
        L("```")
        L("")

    # 2. 관계 — PK 자동 판정 후 PK+FK 또는 INDEX-only
    L("---")
    L("")
    L("## 2. 테이블 간 컬럼 관계")
    L("")
    if not active_rels:
        L("_활성화된 관계가 없습니다._")
        L("")
    else:
        # ── 관계 분류: PK-able vs index-only ──
        # pk_columns: {(file_idx, orig_col)} — PK로 만들 컬럼
        # fk_list: [{'child_idx', 'child_col', 'parent_idx', 'parent_col', 'rel'}]
        # idx_only_rels: [r] — 양쪽 모두 PK 불가능한 관계
        pk_columns = set()
        fk_list = []
        idx_only_rels = []
        for r in active_rels:
            from_idx = r['from_file_idx']
            to_idx = r['to_file_idx']
            from_fd = files_data[from_idx]
            to_fd = files_data[to_idx]
            # 합성 결과 기준 PK 판정 (원본만 unique여도 합성 후 중복이면 PK 미적용)
            from_pk = _is_pk_candidate_synth(from_fd, r['from_col'])
            to_pk = _is_pk_candidate_synth(to_fd, r['to_col'])

            if from_pk and to_pk:
                # 양쪽 모두 unique — 행 수 적은 쪽을 parent로
                if len(from_fd.df) <= len(to_fd.df):
                    parent_idx, parent_col = from_idx, r['from_col']
                    child_idx, child_col = to_idx, r['to_col']
                else:
                    parent_idx, parent_col = to_idx, r['to_col']
                    child_idx, child_col = from_idx, r['from_col']
            elif from_pk:
                parent_idx, parent_col = from_idx, r['from_col']
                child_idx, child_col = to_idx, r['to_col']
            elif to_pk:
                parent_idx, parent_col = to_idx, r['to_col']
                child_idx, child_col = from_idx, r['from_col']
            else:
                idx_only_rels.append(r)
                continue

            pk_columns.add((parent_idx, parent_col))
            fk_list.append({
                'child_idx': child_idx, 'child_col': child_col,
                'parent_idx': parent_idx, 'parent_col': parent_col,
                'rel': r,
            })

        # ── 분석 요약 ──
        n_pk = len(pk_columns)
        n_fk = len(fk_list)
        n_idx = len(idx_only_rels)
        L(f"**관계 분석 결과**: PK 생성 {n_pk}개 컬럼, "
          f"FK 생성 {n_fk}개, INDEX-only 관계 {n_idx}개")
        L("")
        L("**관계 요약 표**")
        L("")
        L("| # | 참조 컬럼 | 관계 | 대상 컬럼 | 신뢰도 | 적용 방식 |")
        L("|---|---|---|---|---|---|")
        for i, r in enumerate(active_rels, 1):
            from_t = table_names[r['from_file_idx']]
            to_t = table_names[r['to_file_idx']]
            from_fd = files_data[r['from_file_idx']]
            to_fd = files_data[r['to_file_idx']]
            from_col_masked = from_fd.column_mask_map.get(r['from_col'], r['from_col'])
            to_col_masked = to_fd.column_mask_map.get(r['to_col'], r['to_col'])
            # 어떤 방식인지 판단
            fk_entry = next((fk for fk in fk_list if fk['rel'] is r), None)
            if fk_entry:
                if fk_entry['parent_idx'] == r['from_file_idx']:
                    mode = "PK(좌) + FK(우)"
                    arrow = "←"
                else:
                    mode = "FK(좌) + PK(우)"
                    arrow = "→"
            else:
                mode = "INDEX-only (PK 불가)"
                arrow = "↔"
            L(f"| {i} | `[{from_t}].[{from_col_masked}]` | {arrow} "
              f"| `[{to_t}].[{to_col_masked}]` "
              f"| {r['confidence']:.2f} | {mode} |")
        L("")

        # ── (a) PRIMARY KEY 생성 ──
        if pk_columns:
            L("### 2-A. PRIMARY KEY 생성 (parent 컬럼)")
            L("")
            L("> 아래 컬럼들은 원본 데이터에서 NULL 없음 + 모든 값 고유로 확인되어 ")
            L("> PRIMARY KEY 적용이 가능합니다. **이 SQL을 먼저 실행하세요.**")
            L("")
            L("```sql")
            for (parent_idx, parent_col) in sorted(pk_columns,
                    key=lambda x: (x[0], str(x[1]))):
                parent_t = table_names[parent_idx]
                parent_fd = files_data[parent_idx]
                parent_col_masked = parent_fd.column_mask_map.get(
                    parent_col, parent_col)
                ctype = parent_fd.col_types.get(parent_col,
                    parent_fd.col_types.get(parent_col_masked, 'categorical'))
                sql_type = _pandas_dtype_to_sql(
                    parent_fd.df[parent_col].dtype, ctype)
                pk_name = f"PK_{parent_t}"
                L(f"-- [{parent_t}].[{parent_col_masked}]: "
                  f"NULL 없음 + unique → PRIMARY KEY")
                L(f"ALTER TABLE [{parent_t}] "
                  f"ALTER COLUMN [{parent_col_masked}] {sql_type} NOT NULL;")
                L(f"IF NOT EXISTS (SELECT 1 FROM sys.key_constraints "
                  f"WHERE name = N'{pk_name}' "
                  f"AND parent_object_id = OBJECT_ID(N'{parent_t}'))")
                L(f"    ALTER TABLE [{parent_t}] ADD CONSTRAINT [{pk_name}] "
                  f"PRIMARY KEY ([{parent_col_masked}]);")
                L("")
            L("```")
            L("")

        # ── (b) INDEX + FK 생성 ──
        if fk_list:
            L("### 2-B. FOREIGN KEY 생성 (child 측 INDEX → FK 순서)")
            L("")
            L("> child 컬럼에 INDEX를 만든 뒤 위에서 만든 PK를 참조하는 FK를 추가합니다.")
            L("")
            L("```sql")
            for fk in fk_list:
                child_idx_ = fk['child_idx']
                child_col = fk['child_col']
                parent_idx_ = fk['parent_idx']
                parent_col = fk['parent_col']
                child_t = table_names[child_idx_]
                parent_t = table_names[parent_idx_]
                child_fd = files_data[child_idx_]
                parent_fd = files_data[parent_idx_]
                child_col_masked = child_fd.column_mask_map.get(
                    child_col, child_col)
                parent_col_masked = parent_fd.column_mask_map.get(
                    parent_col, parent_col)
                idx_name = f"idx_{child_t}_{child_col_masked}"
                fk_name = f"FK_{child_t}_{child_col_masked}"
                L(f"-- [{child_t}].[{child_col_masked}]  →  "
                  f"[{parent_t}].[{parent_col_masked}]")
                L(f"IF NOT EXISTS (SELECT 1 FROM sys.indexes "
                  f"WHERE name = N'{idx_name}' "
                  f"AND object_id = OBJECT_ID(N'{child_t}'))")
                L(f"    CREATE INDEX [{idx_name}] "
                  f"ON [{child_t}]([{child_col_masked}]);")
                L(f"IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys "
                  f"WHERE name = N'{fk_name}' "
                  f"AND parent_object_id = OBJECT_ID(N'{child_t}'))")
                L(f"    ALTER TABLE [{child_t}] ADD CONSTRAINT [{fk_name}] "
                  f"FOREIGN KEY ([{child_col_masked}]) "
                  f"REFERENCES [{parent_t}]([{parent_col_masked}]);")
                L("")
            L("```")
            L("")

        # ── (c) INDEX-only 관계 ──
        if idx_only_rels:
            L("### 2-C. INDEX-only 관계 (PK 적용 불가)")
            L("")
            L("> 양쪽 모두 NULL 또는 중복이 있어 PK/FK 적용이 불가능합니다. ")
            L("> JOIN 성능을 위해 양쪽에 INDEX만 생성합니다.")
            L("")
            L("```sql")
            for r in idx_only_rels:
                from_t = table_names[r['from_file_idx']]
                to_t = table_names[r['to_file_idx']]
                from_fd = files_data[r['from_file_idx']]
                to_fd = files_data[r['to_file_idx']]
                from_col_masked = from_fd.column_mask_map.get(
                    r['from_col'], r['from_col'])
                to_col_masked = to_fd.column_mask_map.get(
                    r['to_col'], r['to_col'])
                from_idx_name = f"idx_{from_t}_{from_col_masked}"
                to_idx_name = f"idx_{to_t}_{to_col_masked}"
                L(f"-- [{from_t}].[{from_col_masked}]  ↔  "
                  f"[{to_t}].[{to_col_masked}]  (PK 불가)")
                L(f"IF NOT EXISTS (SELECT 1 FROM sys.indexes "
                  f"WHERE name = N'{from_idx_name}' "
                  f"AND object_id = OBJECT_ID(N'{from_t}'))")
                L(f"    CREATE INDEX [{from_idx_name}] "
                  f"ON [{from_t}]([{from_col_masked}]);")
                L(f"IF NOT EXISTS (SELECT 1 FROM sys.indexes "
                  f"WHERE name = N'{to_idx_name}' "
                  f"AND object_id = OBJECT_ID(N'{to_t}'))")
                L(f"    CREATE INDEX [{to_idx_name}] "
                  f"ON [{to_t}]([{to_col_masked}]);")
                L("")
            L("```")
            L("")
        L("> 위 SQL을 **2-A → 2-B → 2-C 순서대로** 실행하세요. ")
        L("> 순서를 바꾸면 ALTER COLUMN이 실패할 수 있습니다 ")
        L("> (INDEX가 먼저 걸리면 NOT NULL 변경 불가).")
        L("")

    # 3. 합성 데이터 → 테이블 매핑 (참고)
    L("---")
    L("")
    L("## 3. 합성 데이터 파일 ↔ 테이블 매핑")
    L("")
    L("| # | 합성 데이터 파일 | 적재 대상 테이블 |")
    L("|---|---|---|")
    for idx, fd in enumerate(files_data):
        xl = (os.path.basename(fd.output_xlsx)
              if fd.output_xlsx else '<합성데이터.xlsx>')
        L(f"| {idx+1} | `{xl}` | `[{table_names[idx]}]` |")
    L("")
    L("> 위 매핑대로 `.xlsx`의 데이터를 각 테이블에 적재하면 됩니다 "
      "(`BULK INSERT`, `INSERT`, 또는 pandas+SQLAlchemy 등 환경에 맞는 방법 사용).")
    L("")
    L("---")
    L("")
    L("_원본 컬럼명 및 원본 데이터는 본 문서에 포함되지 않습니다 (보안)._  ")
    L("_프로젝트명/컬럼 관계 정보는 데이터 추적성을 위해 보존되며, 원본 PII는 포함되지 않습니다._")

    content = "\n".join(lines)
    with open(save_path, 'w', encoding='utf-8') as f:
        f.write(content)
    return save_path


def generate_table_and_data_py(employee_id, files_data, relationships,
                                save_path):
    """`테이블생성.md` 와 짝을 이루는 임시 실행 스크립트(`tableanddata.py`) 생성.

    실행 순서:
      1. data_processor.dbconn() 으로 MSSQL 연결
      2. 각 테이블 CREATE TABLE
      3. 합성 데이터 xlsx 파일 → executemany INSERT
      4. PK / FK / INDEX 생성 (데이터가 들어간 뒤에 NOT NULL 변환이 가능하므로
         INSERT 다음에 배치)

    `save_path` 와 같은 디렉토리에 합성 xlsx 파일들이 있다고 가정.
    """
    table_names = {idx: _table_name_for_file(employee_id, fd)
                   for idx, fd in enumerate(files_data)}
    active_rels = [r for r in (relationships or []) if r.get('enabled')]

    # ── CREATE TABLE SQL 리스트 (테이블당 1개 블록) ──
    create_sqls = []
    for idx, fd in enumerate(files_data):
        tname = table_names[idx]
        if fd.df is None:
            continue
        ct = fd.col_types
        mask_map = fd.column_mask_map or {}
        col_lines = []
        for orig_col in fd.df.columns:
            masked_col = mask_map.get(orig_col, orig_col)
            ctype = ct.get(orig_col, ct.get(masked_col, 'categorical'))
            sql_type = _pandas_dtype_to_sql(fd.df[orig_col].dtype, ctype)
            col_lines.append(f"        [{masked_col}] {sql_type}")
        cols_sql = ",\n".join(col_lines)
        sql = (
            f"IF OBJECT_ID(N'{tname}', N'U') IS NULL\n"
            f"BEGIN\n"
            f"    CREATE TABLE [{tname}] (\n"
            f"{cols_sql}\n"
            f"    );\n"
            f"END;"
        )
        create_sqls.append(sql)

    # ── PK / FK / INDEX SQL 리스트 — generate_table_creation_md 와 동일 규칙 ──
    pk_columns = set()
    fk_list = []
    idx_only_rels = []
    for r in active_rels:
        from_idx = r['from_file_idx']
        to_idx = r['to_file_idx']
        from_fd = files_data[from_idx]
        to_fd = files_data[to_idx]
        from_pk = _is_pk_candidate_synth(from_fd, r['from_col'])
        to_pk = _is_pk_candidate_synth(to_fd, r['to_col'])
        if from_pk and to_pk:
            if len(from_fd.df) <= len(to_fd.df):
                parent_idx, parent_col = from_idx, r['from_col']
                child_idx, child_col = to_idx, r['to_col']
            else:
                parent_idx, parent_col = to_idx, r['to_col']
                child_idx, child_col = from_idx, r['from_col']
        elif from_pk:
            parent_idx, parent_col = from_idx, r['from_col']
            child_idx, child_col = to_idx, r['to_col']
        elif to_pk:
            parent_idx, parent_col = to_idx, r['to_col']
            child_idx, child_col = from_idx, r['from_col']
        else:
            idx_only_rels.append(r)
            continue
        pk_columns.add((parent_idx, parent_col))
        fk_list.append({
            'child_idx': child_idx, 'child_col': child_col,
            'parent_idx': parent_idx, 'parent_col': parent_col,
        })

    pk_sqls = []
    for (parent_idx, parent_col) in sorted(pk_columns,
            key=lambda x: (x[0], str(x[1]))):
        parent_t = table_names[parent_idx]
        parent_fd = files_data[parent_idx]
        parent_col_masked = parent_fd.column_mask_map.get(
            parent_col, parent_col)
        ctype = parent_fd.col_types.get(parent_col,
            parent_fd.col_types.get(parent_col_masked, 'categorical'))
        sql_type = _pandas_dtype_to_sql(
            parent_fd.df[parent_col].dtype, ctype)
        pk_name = f"PK_{parent_t}"
        pk_sqls.append(
            f"ALTER TABLE [{parent_t}] "
            f"ALTER COLUMN [{parent_col_masked}] {sql_type} NOT NULL;"
        )
        pk_sqls.append(
            f"IF NOT EXISTS (SELECT 1 FROM sys.key_constraints "
            f"WHERE name = N'{pk_name}' "
            f"AND parent_object_id = OBJECT_ID(N'{parent_t}'))\n"
            f"    ALTER TABLE [{parent_t}] ADD CONSTRAINT [{pk_name}] "
            f"PRIMARY KEY ([{parent_col_masked}]);"
        )

    fk_sqls = []
    for fk in fk_list:
        child_t = table_names[fk['child_idx']]
        parent_t = table_names[fk['parent_idx']]
        child_fd = files_data[fk['child_idx']]
        parent_fd = files_data[fk['parent_idx']]
        child_col_masked = child_fd.column_mask_map.get(
            fk['child_col'], fk['child_col'])
        parent_col_masked = parent_fd.column_mask_map.get(
            fk['parent_col'], fk['parent_col'])
        idx_name = f"idx_{child_t}_{child_col_masked}"
        fk_name = f"FK_{child_t}_{child_col_masked}"
        fk_sqls.append(
            f"IF NOT EXISTS (SELECT 1 FROM sys.indexes "
            f"WHERE name = N'{idx_name}' "
            f"AND object_id = OBJECT_ID(N'{child_t}'))\n"
            f"    CREATE INDEX [{idx_name}] "
            f"ON [{child_t}]([{child_col_masked}]);"
        )
        fk_sqls.append(
            f"IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys "
            f"WHERE name = N'{fk_name}' "
            f"AND parent_object_id = OBJECT_ID(N'{child_t}'))\n"
            f"    ALTER TABLE [{child_t}] ADD CONSTRAINT [{fk_name}] "
            f"FOREIGN KEY ([{child_col_masked}]) "
            f"REFERENCES [{parent_t}]([{parent_col_masked}]);"
        )

    idx_only_sqls = []
    for r in idx_only_rels:
        from_t = table_names[r['from_file_idx']]
        to_t = table_names[r['to_file_idx']]
        from_fd = files_data[r['from_file_idx']]
        to_fd = files_data[r['to_file_idx']]
        from_col_masked = from_fd.column_mask_map.get(
            r['from_col'], r['from_col'])
        to_col_masked = to_fd.column_mask_map.get(
            r['to_col'], r['to_col'])
        from_idx_name = f"idx_{from_t}_{from_col_masked}"
        to_idx_name = f"idx_{to_t}_{to_col_masked}"
        idx_only_sqls.append(
            f"IF NOT EXISTS (SELECT 1 FROM sys.indexes "
            f"WHERE name = N'{from_idx_name}' "
            f"AND object_id = OBJECT_ID(N'{from_t}'))\n"
            f"    CREATE INDEX [{from_idx_name}] "
            f"ON [{from_t}]([{from_col_masked}]);"
        )
        idx_only_sqls.append(
            f"IF NOT EXISTS (SELECT 1 FROM sys.indexes "
            f"WHERE name = N'{to_idx_name}' "
            f"AND object_id = OBJECT_ID(N'{to_t}'))\n"
            f"    CREATE INDEX [{to_idx_name}] "
            f"ON [{to_t}]([{to_col_masked}]);"
        )

    # ── 데이터 파일 → 테이블 매핑 ──
    data_files = []
    for idx, fd in enumerate(files_data):
        if not fd.output_xlsx:
            continue
        data_files.append((os.path.basename(fd.output_xlsx),
                            table_names[idx]))

    # ── Python 스크립트 본문 생성 ──
    def _q(lst):
        """SQL 리스트를 Python 리터럴(triple-quoted 문자열의 list)로 직렬화."""
        if not lst:
            return "[]"
        items = []
        for s in lst:
            # triple-quoted 문자열 안에서 충돌나지 않도록 r-string 사용
            items.append('    r"""' + s + '""",')
        return "[\n" + "\n".join(items) + "\n]"

    py_lines = [
        '"""tableanddata.py — 임시 실행 스크립트 (테이블 생성 + 데이터 INSERT)',
        '',
        '이 스크립트는 `테이블생성.md` 와 짝을 이루며, MD 안의 SQL 과 합성 데이터를',
        '한 번에 DB 에 반영합니다. 1회 실행 후 삭제해도 무방한 임시 파일입니다.',
        '',
        '사전 조건:',
        '  - 같은 폴더에 `data_processor.py` 가 작성되어 있어 `dbconn()` 을 import 가능',
        '  - 같은 폴더에 합성 데이터 xlsx 파일들이 존재',
        '  - .env 또는 환경변수로 DB 접속 정보 설정',
        '',
        '실행:',
        '    python tableanddata.py',
        '"""',
        'import os',
        'import sys',
        'import pandas as pd',
        '',
        'HERE = os.path.dirname(os.path.abspath(__file__))',
        'if HERE not in sys.path:',
        '    sys.path.insert(0, HERE)',
        '',
        'from data_processor import dbconn  # noqa: E402',
        '',
        f'PROJECT = {employee_id!r}',
        '',
        '# ──────────────────────────────────────────────',
        '# 1. CREATE TABLE SQL',
        '# ──────────────────────────────────────────────',
        f'CREATE_TABLE_SQLS = {_q(create_sqls)}',
        '',
        '# ──────────────────────────────────────────────',
        '# 2. PRIMARY KEY SQL (INSERT 후 실행)',
        '# ──────────────────────────────────────────────',
        f'PK_SQLS = {_q(pk_sqls)}',
        '',
        '# ──────────────────────────────────────────────',
        '# 3. FOREIGN KEY + INDEX SQL',
        '# ──────────────────────────────────────────────',
        f'FK_SQLS = {_q(fk_sqls)}',
        '',
        '# ──────────────────────────────────────────────',
        '# 4. INDEX-only 관계 (PK 적용 불가 케이스)',
        '# ──────────────────────────────────────────────',
        f'INDEX_ONLY_SQLS = {_q(idx_only_sqls)}',
        '',
        '# ──────────────────────────────────────────────',
        '# 5. 합성 데이터 파일 → 테이블 매핑',
        '# ──────────────────────────────────────────────',
        f'DATA_FILES = {data_files!r}',
        '',
        '',
        'def _run_batch(cur, sqls, label):',
        '    """SQL 리스트를 순차 실행. 오류 시 컨텍스트와 함께 재발생."""',
        '    for i, sql in enumerate(sqls, 1):',
        '        try:',
        '            cur.execute(sql)',
        '        except Exception as e:',
        '            raise RuntimeError(',
        '                f"[{label} #{i}] SQL 실행 실패: {e}\\n--- SQL ---\\n{sql}"',
        '            ) from e',
        '',
        '',
        'def _insert_xlsx(cur, file_name, table_name):',
        '    """xlsx 파일 → executemany INSERT."""',
        '    path = os.path.join(HERE, file_name)',
        '    if not os.path.exists(path):',
        '        raise FileNotFoundError(f"합성 데이터 파일 없음: {path}")',
        '    df = pd.read_excel(path)',
        '    if df.empty:',
        '        print(f"  [skip] {file_name} — 빈 파일")',
        '        return 0',
        '    # NaN/NaT → None (DB NULL)',
        '    df = df.astype(object).where(df.notna(), None)',
        '    cols = list(df.columns)',
        '    col_names = ",".join(f"[{c}]" for c in cols)',
        '    placeholders = ",".join("?" * len(cols))',
        '    sql = f"INSERT INTO [{table_name}] ({col_names}) VALUES ({placeholders})"',
        '    rows = list(df.itertuples(index=False, name=None))',
        '    try:',
        '        cur.fast_executemany = True',
        '    except Exception:',
        '        pass',
        '    cur.executemany(sql, rows)',
        '    return len(rows)',
        '',
        '',
        'def main():',
        '    print(f"[tableanddata] project={PROJECT}")',
        '    conn = dbconn()',
        '    if conn is None:',
        '        raise RuntimeError(',
        '            "dbconn() 이 None 을 반환했습니다. .env / 환경변수 확인 필요.")',
        '    cur = conn.cursor()',
        '    try:',
        '        print(f"[1/4] 테이블 생성 ({len(CREATE_TABLE_SQLS)}개)...")',
        '        _run_batch(cur, CREATE_TABLE_SQLS, "CREATE")',
        '        conn.commit()',
        '',
        '        print(f"[2/4] 데이터 INSERT ({len(DATA_FILES)}개 파일)...")',
        '        total = 0',
        '        for file_name, table_name in DATA_FILES:',
        '            n = _insert_xlsx(cur, file_name, table_name)',
        '            print(f"  ✅ {file_name} → [{table_name}] : {n} rows")',
        '            total += n',
        '        conn.commit()',
        '        print(f"      총 {total} rows INSERT 완료")',
        '',
        '        print(f"[3/4] PRIMARY KEY 적용 ({len(PK_SQLS)} statements)...")',
        '        _run_batch(cur, PK_SQLS, "PK")',
        '        conn.commit()',
        '',
        '        print(f"[4/4] FK + INDEX ({len(FK_SQLS) + len(INDEX_ONLY_SQLS)} statements)...")',
        '        _run_batch(cur, FK_SQLS, "FK")',
        '        _run_batch(cur, INDEX_ONLY_SQLS, "INDEX")',
        '        conn.commit()',
        '',
        '        print("🎉 완료 — 본 스크립트는 삭제해도 됩니다.")',
        '    except Exception:',
        '        conn.rollback()',
        '        raise',
        '    finally:',
        '        cur.close()',
        '        conn.close()',
        '',
        '',
        'if __name__ == "__main__":',
        '    main()',
        '',
    ]
    content = "\n".join(py_lines)
    with open(save_path, 'w', encoding='utf-8') as f:
        f.write(content)
    return save_path


def generate_data_processor_md(employee_id, files_data, save_path,
                                relationships=None):
    """`data_processor.py` 코드 작성용 별도 MD 생성.

    LLM(Claude/Gemini)이 이 파일 한 개만 입력받아도 외부망–내부망 인터페이스
    레이어 역할을 하는 `data_processor.py`를 작성할 수 있도록,
    4개 필수 함수 명세 + 테이블 구조 + 컬럼 설명 + 관계 정보를 모두 포함한다.
    """
    lines = []
    L = lines.append

    table_names = {}
    for idx, fd in enumerate(files_data):
        table_names[idx] = _table_name_for_file(employee_id, fd)

    active_rels = [r for r in (relationships or []) if r.get('enabled')]

    # 설명 미입력 컬럼 집계 (경고 섹션 노출 여부 결정)
    missing_desc = []
    for idx, fd in enumerate(files_data):
        if fd.df is None:
            continue
        descs = fd.col_descriptions or {}
        mask_map = fd.column_mask_map or {}
        for orig_col in fd.df.columns:
            if not (descs.get(orig_col) or '').strip():
                masked_col = mask_map.get(orig_col, orig_col)
                missing_desc.append((table_names[idx], str(masked_col)))

    # ── 헤더 ──
    L(f"# `data_processor.py` 작성 지시서 (프로젝트명: `{employee_id}`)")
    L("")
    L(f"- 생성 시각: {datetime.now().isoformat(timespec='seconds')}")
    L(f"- 대상 테이블 수: **{len(files_data)}개**")
    L(f"- 활성 관계 수: **{len(active_rels)}개**")
    L("")
    L("---")
    L("")

    # ── 작성 지시 (핵심) ──
    L("## 📌 작성 지시")
    L("")
    L("`data_processor.py`는 **외부망–내부망 인터페이스 레이어**입니다.")
    L("`app.py` / `chart.py` / `ML.py`는 본 모듈이 제공하는 **표준 컬럼명 상수**와")
    L("**`load_table(table_name)`** 만 사용해야 하며, 내부망에서는 본 파일의")
    L("`COLUMN_MAPPING` dict의 **키만** 실제 컬럼명으로 갱신하면 다운스트림 코드가")
    L("그대로 동작합니다.")
    L("")
    L("> 📌 **작성 중에는 자유롭게 수정 가능합니다.** read-only 전환은 §체크리스트가")
    L("> **전부 완료된 뒤 마지막 단계**에서 진행합니다 (아래 §🔒 참고).")
    L("")
    L("**파이프라인 (4단계)**")
    L("")
    L("1. **DB 연결**       — `dbconn()`")
    L("2. **로우 데이터 조회** — `fetch_raw(table_name)` : `SELECT * FROM [table]` → DataFrame")
    L("3. **매핑 생성**     — `build_mapping(table_name)` : `{원본 컬럼명: 표준 컬럼명 상수}` dict")
    L("4. **표준화된 df 반환** — `load_table(table_name)` : 위 셋을 묶어 rename 완료된 df 반환")
    L("")
    L("**다운스트림(app.py/chart.py/ML.py) 호출 규약**")
    L("")
    L("```python")
    L("from data_processor import (")
    L("    load_table,")
    L("    TBL_CUSTOMER,          # 테이블명도 표준 상수만 사용")
    L("    COL_CUSTOMER_NAME,")
    L("    COL_ORDER_DATE,")
    L(")")
    L("df = load_table(TBL_CUSTOMER)        # ⚠ load_table(\"<문자열>\") 직접 호출 금지")
    L("df[COL_CUSTOMER_NAME]                # ⚠ 한글/마스킹/원본 컬럼명 리터럴 금지")
    L("```")
    L("")
    L("> 🔑 **테이블명도 컬럼명과 동일하게 디커플링**됩니다. 다운스트림 코드에는 ")
    L("> `\"testd_aaaa\"` / `\"customer_master\"` 같은 **실제 DB 테이블명 문자열**이 ")
    L("> 단 한 번도 등장하면 안 됩니다 — 오직 `TBL_XXX` 상수만.")
    L("")
    L("---")
    L("")

    # ── 필수 4개 함수 명세 ──
    L("## 🧩 필수 4개 함수 명세")
    L("")
    L("아래 시그니처를 **그대로** 구현하세요 (이름/인수/리턴 타입 변경 금지).")
    L("")
    L("```python")
    L("def dbconn() -> \"pyodbc.Connection\":")
    L("    \"\"\"MSSQL 연결. 환경변수")
    L("    (DB_SERVER / DB_PORT / DB_NAME / DB_USER / DB_PASSWORD / DB_DRIVER) 기반.")
    L("    \"\"\"")
    L("")
    L("def fetch_raw(table_name: str) -> pd.DataFrame:")
    L("    \"\"\"SELECT * FROM [table_name] — 원본 컬럼명 그대로 반환.")
    L("")
    L("    ⚠ 메모리 절약 필수: `pd.read_sql(..., chunksize=...)` 로 chunk 단위 읽기 +")
    L("    chunk 마다 dtype 다운캐스팅 후 concat (구체 절차는 §💾 참고).")
    L("    \"\"\"")
    L("")
    L("def build_mapping(table_name: str) -> dict:")
    L("    \"\"\"COLUMN_MAPPING[table_name]을 반환 — {원본 컬럼명: 표준 컬럼명 상수}.\"\"\"")
    L("")
    L("def load_table(table_name: str) -> pd.DataFrame:")
    L("    \"\"\"fetch_raw 후 build_mapping으로 rename — 다운스트림이 쓰는 유일한 진입점.\"\"\"")
    L("    df = fetch_raw(table_name)")
    L("    return df.rename(columns=build_mapping(table_name))")
    L("```")
    L("")
    L("**`table_name` 인수 규약 — 테이블명 디커플링**")
    L("")
    L("- 다운스트림이 넘기는 `table_name` 은 **표준 테이블 상수의 값(영문 snake_case)**")
    L("  입니다. 예: `TBL_CUSTOMER = \"customer\"` 일 때 `load_table(TBL_CUSTOMER)`.")
    L("- `fetch_raw` 내부에서 **`TABLE_MAPPING[table_name]`** 으로 실제 DB 테이블명을")
    L("  lookup 한 뒤 그 결과로만 `SELECT` 합니다. `table_name` 을 SQL 에 직접 ")
    L("  format 하지 마세요.")
    L("- `build_mapping` / `COLUMN_MAPPING` 의 outer 키도 동일하게 표준 테이블 상수입니다.")
    L(f"- `테이블생성.md` 가 만든 실제 DB 테이블명(`{employee_id}_<논리테이블명>`)은")
    L("  **`TABLE_MAPPING` 의 *값*** 으로만 등장하며, 내부망 반입 시에는 이 값만")
    L("  실제 운영 테이블명으로 교체합니다.")
    L("")
    L("**`dbconn()` 구현 힌트**: 본 프로젝트 동봉 코드의 `_db_connect()`와")
    L("동일한 환경변수/접속 문자열을 사용하세요 (DRIVER + SERVER+PORT + DATABASE +")
    L("UID + PWD + TrustServerCertificate=yes).")
    L("")
    L("---")
    L("")

    # ── 메모리 효율 가이드 ──
    L("## 💾 메모리 효율 처리 (필수)")
    L("")
    L("실무 테이블은 수십 ~ 수백만 행이 될 수 있으므로 `fetch_raw` 는 **반드시**")
    L("아래 절차로 메모리 부담을 낮춰 구현하세요. 단일 `pd.read_sql(...)` 호출로")
    L("테이블 전체를 한 번에 메모리에 올리는 구현은 금지합니다.")
    L("")
    L("**1) chunk 단위 읽기**")
    L("")
    L("- `pd.read_sql(..., chunksize=50_000)` 로 iterator 를 받아 chunk 마다 처리")
    L("- 권장 기본값 `CHUNK_SIZE = 50_000` (모듈 상단 상수로 노출, 필요 시 조정)")
    L("- chunk 마다 즉시 다운캐스팅(아래 2) → 리스트에 누적 → 마지막에 단 1회만 `pd.concat`")
    L("")
    L("**2) dtype 다운캐스팅**")
    L("")
    L("- `int64` → `pd.to_numeric(..., downcast='integer')`")
    L("- `float64` → `pd.to_numeric(..., downcast='float')`")
    L("- 카디널리티 낮은 문자열(object, 고유값 / 전체행 ≤ 0.5) → `astype('category')`")
    L("- 다운캐스팅은 chunk 가 메모리에 작은 동안 수행해야 효과가 큼 (concat 후가 아님)")
    L("")
    L("**3) 커서 / 연결 자원 해제**")
    L("")
    L("- `with dbconn() as conn:` 컨텍스트 매니저로 연결을 명시적으로 닫음")
    L("- `chunks` 리스트는 `pd.concat` 직후 `del chunks; gc.collect()` 로 해제")
    L("- 다운스트림에서 `load_table` 결과를 다 쓴 뒤에도 변수 재바인딩으로 GC 유도 권장")
    L("")
    L("**4) 컬럼 절약 (옵션)**")
    L("")
    L("- 다운스트림이 일부 컬럼만 쓴다는 사실이 분명하면 `SELECT *` 대신")
    L("  `SELECT [c1],[c2]` 로 좁혀도 무방 — 단 `COLUMN_MAPPING` 의 키와 정확히 매칭")
    L("- 이 경우 `fetch_raw` 의 별도 파라미터로 받지 말고, 별도 헬퍼")
    L("  (`fetch_raw_columns(table_name, cols)`) 추가는 금지 — 4개 함수 외 신규 함수 추가 금지")
    L("- 컬럼 좁히기가 필요하면 `fetch_raw` 본체 내부에서 `COLUMN_MAPPING[table_name].keys()`")
    L("  를 SELECT 절에 사용해 항상 일관되게 적재")
    L("")
    L("**5) 금지 사항**")
    L("")
    L("- `pd.read_sql(..., chunksize=None)` 로 일괄 로드 (대용량에서 OOM 발생)")
    L("- chunk 마다 `pd.concat` 호출 (O(N²) 비용 — 마지막 1회만)")
    L("- `df.copy()` 의 무의미한 반복 호출")
    L("- 다운캐스팅 없이 `object` dtype 유지 (한국어 문자열 컬럼은 거의 항상 category 가능)")
    L("")
    L("---")
    L("")

    # ── ⚠ 경고: 설명 미입력 (조건부) ──
    if missing_desc:
        L("## ⚠ 경고: 설명 미입력 컬럼")
        L("")
        L(f"설명이 입력되지 않은 컬럼이 **{len(missing_desc)}개** 있습니다.")
        L("")
        L("LLM은 마스킹 코드(예: `cRG_77`)와 타입만으로 영문 snake_case 표준명을")
        L("**추정**한 뒤, 해당 상수 선언 옆에 반드시 다음 형식의 주석을 남기세요:")
        L("")
        L("```python")
        L("COL_XXX = \"xxx\"  # TODO: 사용자 확인 필요 — 원본 의미 불명")
        L("```")
        L("")
        L("**해당 컬럼 (테이블 / 마스킹 코드)**")
        L("")
        for tname, masked in missing_desc:
            L(f"- `[{tname}]` / `[{masked}]`")
        L("")
        L("---")
        L("")

    # ── 테이블 구조 + 컬럼 설명 ──
    L("## 📋 참고: 테이블 구조와 컬럼 설명")
    L("")
    L("아래 정보를 읽고 각 컬럼에 대해 적절한 영문 snake_case 표준 컬럼명을 결정하세요.")
    L("")

    for idx, fd in enumerate(files_data):
        tname = table_names[idx]
        L(f"### {idx+1}. `[{tname}]`")
        L("")
        L(f"- 실제 DB 테이블명 (마스킹): **`{tname}`** ← `TABLE_MAPPING` 의 *값*")
        L(f"- 표준 테이블 상수: **`TBL_???`** ← LLM 이 컬럼 설명을 보고 결정 ")
        L(f"  (영문 snake_case 값을 가진 모듈 상수, 예: `TBL_CUSTOMER = \"customer\"`)")
        L(f"- 행 수: {len(fd.df) if fd.df is not None else '?'}")
        L("")
        L("| # | 컬럼명(마스킹됨) | 타입 | 설명 |")
        L("|---|---|---|---|")

        df = fd.df
        ct = fd.col_types
        descs = fd.col_descriptions or {}
        mask_map = fd.column_mask_map or {}

        if df is not None:
            for i, orig_col in enumerate(df.columns, 1):
                masked_col = mask_map.get(orig_col, orig_col)
                ctype = ct.get(orig_col,
                                ct.get(masked_col, 'categorical'))
                raw_desc = (descs.get(orig_col, '') or '').strip()
                desc = raw_desc if raw_desc else '⚠ _(설명 미입력)_'
                L(f"| {i} | `[{masked_col}]` | {ctype} | {desc} |")
        L("")

    # ── 활성 관계 ──
    if active_rels:
        L("### 🔗 테이블 간 활성 관계")
        L("")
        L("아래 컬럼 페어에는 **같은 표준 상수**(또는 의미가 일관된 명명)를 부여하세요.")
        L("")
        L("| from 테이블.컬럼 | → | to 테이블.컬럼 | 신뢰도 |")
        L("|---|---|---|---|")
        mask_maps = [fd.column_mask_map or {} for fd in files_data]
        for r in active_rels:
            fi = r['from_file_idx']
            ti = r['to_file_idx']
            from_masked = mask_maps[fi].get(r['from_col'], r['from_col'])
            to_masked = mask_maps[ti].get(r['to_col'], r['to_col'])
            conf = r.get('confidence', 0)
            L(f"| `[{table_names[fi]}]`.`[{from_masked}]` | → "
              f"| `[{table_names[ti]}]`.`[{to_masked}]` | {conf} |")
        L("")

    L("---")
    L("")

    # ── 산출물 형식 예시 ──
    L("## 🐍 산출물 형식 예시 (LLM이 채울 자리)")
    L("")
    L("```python")
    L("# data_processor.py")
    L("# (체크리스트 완료 후 read-only 로 전환 — 아래 §🔒 참고)")
    L("import gc")
    L("import os")
    L("import pandas as pd")
    L("import pyodbc")
    L("")
    L("# fetch_raw 의 chunk 크기 — 메모리 ↔ 속도 trade-off (필요 시 조정)")
    L("CHUNK_SIZE = 50_000")
    L("")
    L("# ───── 표준 테이블명 상수 (불변) ─────")
    L("#   다운스트림은 이 상수만 import 해서 load_table() 인수로 넘김.")
    L("TBL_CUSTOMER = \"customer\"")
    L("TBL_ORDER = \"order\"")
    L("# ... 위 §참고의 모든 테이블에 대해 한 줄씩")
    L("")
    L("# ───── 표준 컬럼명 상수 (불변) ─────")
    L("COL_CUSTOMER_NAME = \"customer_name\"")
    L("COL_ORDER_DATE = \"order_date\"")
    L("# ... 위 §참고의 모든 테이블·모든 컬럼에 대해 한 줄씩")
    L("")
    L("# ───── 표준 테이블 상수 → 실제 DB 테이블명 매핑 ─────")
    L("#   내부망 반입 시 이 dict의 *값(오른쪽)* 만 실제 운영 테이블명으로 갱신.")
    L("#   키(상수)와 함수 본체는 절대 변경 금지.")
    L("TABLE_MAPPING = {")
    L(f"    TBL_CUSTOMER: \"{employee_id}_<논리테이블1>\",")
    L(f"    TBL_ORDER:    \"{employee_id}_<논리테이블2>\",")
    L("    # ...")
    L("}")
    L("")
    L("# ───── 원본/DB 컬럼명 → 표준 컬럼명 매핑 ─────")
    L("#   outer 키는 표준 테이블 상수(TBL_XXX) — 절대 변경 금지.")
    L("#   inner 키(왼쪽)는 실제 DB 컬럼명 — 내부망 반입 시 이 값만 갱신.")
    L("#   inner 값(오른쪽)은 표준 컬럼 상수(COL_XXX) — 절대 변경 금지.")
    L("COLUMN_MAPPING = {")
    L("    TBL_CUSTOMER: {")
    L("        \"<원본/DB 컬럼명>\": COL_CUSTOMER_NAME,")
    L("        # ...")
    L("    },")
    L("    # ...")
    L("}")
    L("")
    L("")
    L("def dbconn() -> pyodbc.Connection:")
    L("    \"\"\"MSSQL 연결.\"\"\"")
    L("    driver = os.getenv(\"DB_DRIVER\", \"ODBC Driver 17 for SQL Server\")")
    L("    conn_str = (")
    L("        f\"DRIVER={{{driver}}};\"")
    L("        f\"SERVER={os.getenv('DB_SERVER')},{os.getenv('DB_PORT')};\"")
    L("        f\"DATABASE={os.getenv('DB_NAME')};\"")
    L("        f\"UID={os.getenv('DB_USER')};\"")
    L("        f\"PWD={os.getenv('DB_PASSWORD')};\"")
    L("        \"TrustServerCertificate=yes;\"")
    L("    )")
    L("    return pyodbc.connect(conn_str)")
    L("")
    L("")
    L("def _downcast(df: pd.DataFrame) -> pd.DataFrame:")
    L("    \"\"\"chunk 단위 dtype 다운캐스팅 — 메모리 절약.\"\"\"")
    L("    for c in df.select_dtypes(include='integer').columns:")
    L("        df[c] = pd.to_numeric(df[c], downcast='integer')")
    L("    for c in df.select_dtypes(include='float').columns:")
    L("        df[c] = pd.to_numeric(df[c], downcast='float')")
    L("    for c in df.select_dtypes(include='object').columns:")
    L("        # 카디널리티 낮은 문자열 → category 로 변환")
    L("        if df[c].nunique(dropna=False) / max(len(df), 1) <= 0.5:")
    L("            df[c] = df[c].astype('category')")
    L("    return df")
    L("")
    L("")
    L("def fetch_raw(table_name: str) -> pd.DataFrame:")
    L("    \"\"\"표준 테이블 상수(TBL_XXX) → TABLE_MAPPING 으로 실제 DB 테이블명 lookup 후")
    L("    SELECT. chunk 단위 로딩 + dtype 다운캐스팅.\"\"\"")
    L("    if table_name not in TABLE_MAPPING:")
    L("        raise KeyError(")
    L("            f\"TABLE_MAPPING 에 등록되지 않은 테이블 상수: {table_name!r}\")")
    L("    real_table = TABLE_MAPPING[table_name]")
    L("    chunks = []")
    L("    with dbconn() as conn:")
    L("        sql = f\"SELECT * FROM [{real_table}]\"")
    L("        for chunk in pd.read_sql(sql, conn, chunksize=CHUNK_SIZE):")
    L("            chunks.append(_downcast(chunk))")
    L("    if not chunks:")
    L("        return pd.DataFrame()")
    L("    df = pd.concat(chunks, ignore_index=True)")
    L("    del chunks")
    L("    gc.collect()")
    L("    return df")
    L("")
    L("")
    L("def build_mapping(table_name: str) -> dict:")
    L("    \"\"\"{원본 컬럼명: 표준 컬럼명 상수} 반환.\"\"\"")
    L("    return COLUMN_MAPPING.get(table_name, {})")
    L("")
    L("")
    L("def load_table(table_name: str) -> pd.DataFrame:")
    L("    \"\"\"원본 → 표준 컬럼명으로 rename된 df 반환 — 다운스트림 진입점.\"\"\"")
    L("    df = fetch_raw(table_name)")
    L("    return df.rename(columns=build_mapping(table_name))")
    L("```")
    L("")
    L("---")
    L("")

    # ── 체크리스트 ──
    L("## ✅ 체크리스트")
    L("")
    L("- [ ] 모든 테이블에 대해 **표준 테이블 상수 `TBL_XXX`** 정의됨 (영문 snake_case)")
    L("- [ ] `TABLE_MAPPING` 이 `{TBL_XXX: \"<실제 DB 테이블명>\"}` 형태로 작성됨")
    L("- [ ] 모든 컬럼에 대해 **표준 컬럼 상수 `COL_XXX`** 정의됨")
    L("- [ ] `COLUMN_MAPPING` 의 **outer 키는 모두 `TBL_XXX` 상수** (문자열 리터럴 금지)")
    L("- [ ] `dbconn`, `fetch_raw`, `build_mapping`, `load_table` 4개 함수 모두 구현됨")
    L("- [ ] `fetch_raw` 가 `TABLE_MAPPING[table_name]` 으로 실제 명 lookup 후 SQL 실행")
    L("- [ ] `load_table(table_name)` 의 반환 df 컬럼명이 모두 표준 상수 값과 일치함")
    L("- [ ] 빈 설명 컬럼에 `# TODO: 사용자 확인 필요` 주석 부착됨 (해당되면)")
    L("- [ ] `app.py` / `chart.py` / `ML.py` 안에 ")
    L("      **테이블명·컬럼명 문자열 리터럴이 단 한 곳도 없음** (오직 `TBL_XXX`/`COL_XXX`)")
    L("- [ ] **메모리 효율**: `fetch_raw` 가 chunk 단위 읽기 + dtype 다운캐스팅 구현됨 ")
    L("      (§💾 참고 — 단일 `pd.read_sql` 일괄 로드 금지)")
    L("- [ ] 위 항목 모두 통과 후, 아래 §🔒 절차에 따라 파일을 read-only 로 전환")
    L("")
    L("---")
    L("")

    # ── read-only 전환 프로세스 ──
    L("## 🔒 read-only 전환 (체크리스트 완료 후 실행)")
    L("")
    L("**위 §체크리스트의 모든 항목이 체크된 것을 확인한 뒤**, 아래 절차로")
    L("`data_processor.py` 를 읽기 전용으로 전환합니다.")
    L("")
    L("**내부망 반입 시 갱신 가능한 부분 (이 두 곳만)**")
    L("")
    L("1. `TABLE_MAPPING` 의 **값** — 표준 테이블 상수에 매핑되는 실제 DB 테이블명")
    L("2. `COLUMN_MAPPING` 의 각 inner dict 의 **키** — 실제 DB 컬럼명")
    L("")
    L("그 외 — `TBL_XXX` / `COL_XXX` 상수명·값, `COLUMN_MAPPING` 의 outer 키 ")
    L("(= `TBL_XXX` 참조), 모든 함수 본체·시그니처 — 는 절대 변경 금지.")
    L("")
    L("### 1) 전환 직전 자기 점검")
    L("")
    L("- [ ] 모든 체크리스트가 ✅ 체크됨")
    L("- [ ] `python -c \"import data_processor as d; "
          "d.load_table(d.TBL_CUSTOMER)\"` 로 1회 정상 동작 확인 ")
    L("      (`TBL_CUSTOMER` 자리에 §참고의 임의 표준 테이블 상수)")
    L("- [ ] git 등 버전 관리에 커밋된 상태 (실수 시 복구용)")
    L("")
    L("### 2) OS 별 전환 명령")
    L("")
    L("**Windows (PowerShell / cmd)**")
    L("")
    L("```powershell")
    L("attrib +R data_processor.py")
    L("# 또는 GUI: 파일 우클릭 → 속성 → \"읽기 전용\" 체크")
    L("```")
    L("")
    L("**Linux / macOS**")
    L("")
    L("```bash")
    L("chmod 444 data_processor.py")
    L("```")
    L("")
    L("### 3) 전환 후 확인")
    L("")
    L("- 파일을 편집기로 열어 시험 저장 시 \"읽기 전용\" 경고가 떠야 정상")
    L("- 내부망 반입 후 테이블명/컬럼명이 달라져 수정이 필요한 경우:")
    L("  1. 일시적으로 `attrib -R` / `chmod 644` 로 잠금 해제")
    L("  2. **오직 `TABLE_MAPPING` 의 값 + `COLUMN_MAPPING` 의 inner 키만** 수정")
    L("     (상수명·상수값, outer 키, 함수 본체는 절대 건드리지 않음)")
    L("  3. 즉시 다시 read-only 로 재전환")
    L("")

    content = "\n".join(lines)
    with open(save_path, 'w', encoding='utf-8') as f:
        f.write(content)
    return save_path


# ══════════════════════════════════════════════════════════════
# 파일 상태 객체
# ══════════════════════════════════════════════════════════════

class FileTabBar(tk.Frame):
    """ttk.Notebook 대체 — Windows 네이티브 테마에서도 색이 먹는 커스텀 탭바.

    - 상단: tk.Button 행 (탭). 선택=녹색, 비선택=파란색.
    - 본문: 선택된 탭에 대응하는 Frame만 표시.
    - ttk.Notebook의 add/select/tabs/forget/tab API와 부분 호환.
    """

    COLOR_SELECTED = "#27ae60"      # 녹색
    COLOR_UNSELECTED = "#2980b9"    # 파랑
    COLOR_HOVER = "#5dade2"         # 밝은 파랑
    COLOR_BAR_BG = "#ecf0f1"

    def __init__(self, master, **kwargs):
        """탭 버튼 행(bar) + 본문 영역(body)으로 구성된 컨테이너 초기화."""
        super().__init__(master, **kwargs)
        self._tabs = []          # [(button, frame), ...]
        self._selected_idx = -1
        self._last_bar_w = 0     # 마지막 reflow 시 bar 너비 (무한 루프 방지)
        self.bar = tk.Frame(self, bg=self.COLOR_BAR_BG)
        self.bar.pack(side=tk.TOP, fill=tk.X)
        # bar 너비가 바뀌면 탭 버튼을 다시 줄바꿈 배치 (반응형)
        self.bar.bind("<Configure>", self._reflow)
        self.body = tk.Frame(self)
        self.body.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    def new_tab(self, text=""):
        """새 탭을 만들고 컨텐츠를 담을 Frame을 반환.

        호출자는 반환된 Frame을 부모로 child 위젯을 배치.
        """
        idx = len(self._tabs)
        btn = tk.Button(self.bar, text=text,
                        bg=self.COLOR_UNSELECTED, fg="white",
                        activebackground=self.COLOR_HOVER,
                        activeforeground="white",
                        font=("맑은 고딕", 10, "bold"),
                        relief="raised", bd=2, padx=14, pady=6,
                        cursor="hand2",
                        command=lambda i=idx: self.select(i))
        frame = ttk.Frame(self.body, padding=3)
        self._tabs.append((btn, frame))
        if self._selected_idx == -1:
            self.select(idx)
        # grid flow 배치 — 즉시 + idle 시 한 번 더 (최초엔 bar 너비 미확정)
        self._reflow()
        self.after_idle(self._reflow)
        return frame

    def _reflow(self, event=None):
        """탭 버튼을 bar 너비에 맞춰 여러 줄로 자동 줄바꿈(flow) 배치.

        가로 한 줄(side=LEFT) 배치 시 파일이 많으면 화면 밖으로 잘리던 문제를
        웹 프론트엔드의 flex-wrap 처럼 줄바꿈으로 해결.
        """
        if not self._tabs:
            self._last_bar_w = 0
            return
        bar_w = self.bar.winfo_width()
        if bar_w <= 1:
            return  # 아직 레이아웃 전 — Configure 이벤트에서 다시 호출됨
        # Configure 이벤트는 너비 변화가 없으면 무시 (재배치로 인한 무한 루프 방지)
        if event is not None and bar_w == self._last_bar_w:
            return
        self._last_bar_w = bar_w
        r = c = 0
        used = 0
        for btn, _frame in self._tabs:
            bw = btn.winfo_reqwidth() + 4
            if c > 0 and used + bw > bar_w:
                r += 1
                c = 0
                used = 0
            btn.grid(row=r, column=c, sticky="w", padx=(0, 2), pady=(0, 2))
            used += bw
            c += 1

    def select(self, idx):
        """탭 선택 (인덱스 또는 정수 변환 가능한 문자열)."""
        if isinstance(idx, str):
            try:
                idx = int(idx)
            except ValueError:
                return
        if idx < 0 or idx >= len(self._tabs):
            return
        for i, (btn, frame) in enumerate(self._tabs):
            if i == idx:
                btn.config(bg=self.COLOR_SELECTED, relief="sunken")
                frame.pack(in_=self.body, fill=tk.BOTH, expand=True)
            else:
                btn.config(bg=self.COLOR_UNSELECTED, relief="raised")
                frame.pack_forget()
        self._selected_idx = idx

    def tabs(self):
        """현재 등록된 탭 인덱스 리스트 반환 (ttk.Notebook.tabs 호환)."""
        return list(range(len(self._tabs)))

    def tab(self, idx, **kwargs):
        """탭 옵션 변경 — text만 지원."""
        if isinstance(idx, str):
            try:
                idx = int(idx)
            except ValueError:
                return
        if 0 <= idx < len(self._tabs) and 'text' in kwargs:
            self._tabs[idx][0].config(text=kwargs['text'])
            # 라벨(파일명/확정아이콘) 변경으로 버튼 폭이 달라질 수 있어 재배치
            self._last_bar_w = 0
            self._reflow()

    def forget(self, idx):
        """탭 제거."""
        if isinstance(idx, str):
            try:
                idx = int(idx)
            except ValueError:
                return
        if 0 <= idx < len(self._tabs):
            btn, frame = self._tabs[idx]
            btn.destroy()
            frame.destroy()
            del self._tabs[idx]
            for new_i, (b, _) in enumerate(self._tabs):
                b.config(command=lambda i=new_i: self.select(i))
            if self._selected_idx == idx:
                self._selected_idx = -1
                if self._tabs:
                    self.select(0)
            elif self._selected_idx > idx:
                self._selected_idx -= 1
            self._last_bar_w = 0
            self._reflow()

    def clear(self):
        """모든 탭 제거."""
        for btn, frame in self._tabs:
            btn.destroy()
            frame.destroy()
        self._tabs = []
        self._selected_idx = -1
        self._last_bar_w = 0


class FileData:
    """각 엑셀 파일별 상태 (파일별 탭 UI에서 자기 위젯과 입력값을 보존)."""

    def __init__(self, path):
        """엑셀 파일 한 개의 모든 상태(데이터/사용자입력/위젯/출력)를 보관.

        Args:
            path: 원본 엑셀 파일 절대 경로.
        """
        self.path = path
        self.file_name = os.path.basename(path)
        self.sheet_name = ''
        self.all_sheets = []
        self.df = None
        self.info = None
        self.col_types = {}
        self.original_columns = []
        self.pii_auto_detected = set()
        self.id_suspected = set()

        # 사용자 입력값 (위젯의 현재 상태를 항상 동기화)
        self.rename_values = {}      # {orig_col: new_name}
        self.mapping_values = {}     # {col: {orig_val: fake_val}}
        self.col_modes = {}          # {col: 'masked'/'keep'/'fake'}
        self.col_descriptions = {}   # {col: text}

        self.analyzed = False
        self.confirmed = False

        # 사용자 지정 테이블명 (영문 소문자) — UI에서 파일별로 입력
        self.table_name = ''

        # 관계 그룹 정보 — 관계 확정 시 채워짐 (해당 컬럼은 가짜값 read-only)
        self.linked_cols = {}        # {col: group_id}

        # 출력
        self.output_xlsx = ''
        self.column_mask_map = {}    # orig → masked
        self.shuffle_order = None
        # 합성 완료 후 final 데이터에서 unique+non-null인 컬럼 (마스킹된 이름).
        # MD의 PK 판정에 사용 — 원본만 보고 PK 적용하면 행 재샘플링으로 인한
        # 중복으로 INSERT 시 실패하므로 합성 결과로 다시 검증.
        self.synth_unique_cols = set()
        # OTHER 통합 메타데이터 — {col: {k_anon, n_bucketed_originals,
        # total_pre_bucket_rows, original_values: [{value, pre_bucket_freq,
        # fake_code}, ...]}}. 변환키 JSON에 기록되어 복원/필터링에 사용.
        self.other_buckets = {}

        # ── 파일별 탭 위젯 참조 (탭 빌드 시 채워짐) ──
        self.tab_frame = None            # 내부 노트북의 이 파일 탭 Frame
        self.rename_inner = None         # 컬럼명 변경 위젯 컨테이너
        self.input_inner = None          # 가짜값 입력 위젯 컨테이너
        self.confirm_btn = None          # 이 파일 확정 버튼
        self.confirm_status_lbl = None   # 이 파일 확정 상태 라벨
        self.analysis_text = None        # 이 파일 분석 요약 텍스트

        # 위젯 dict (탭 빌드 후 채워짐)
        self.rename_entries = {}         # {col: Entry}
        self.input_entry_map = {}        # {col: [(orig_val, Entry_or_None, fixed_val), ...]}
        self.mode_vars = {}              # {col: StringVar}
        self.desc_entries = {}           # {col: Entry}


# ══════════════════════════════════════════════════════════════
# GUI
# ══════════════════════════════════════════════════════════════

class SynthesizeAppV2:
    """v2 메인 GUI — 다중 엑셀 파일 일괄 마스킹 + 관계 보존 + MD 생성.

    워크플로:
        ❶ 파일 추가 + 사번 → ❷ 전체 분석 → ❸ 관계 설정/확정
        → ❹ 파일별 컬럼 편집/확정 → ❺ 실행 + MD 생성

    Tk root 하나에 모든 위젯을 build하고, FileData 리스트로 파일별 상태를 보관.
    """

    def __init__(self, root):
        """GUI 상태 변수 초기화 + 위젯 build + 첫 단계 표시.

        Args:
            root: tk.Tk 인스턴스.
        """
        self.root = root
        self.root.title("합성 데이터 자동 생성기 v2  —  다중 엑셀 + 관계 보존  |  삼성중공업 생산 DT센터")

        # ── 화면 크기에 맞춰 반응형으로 창 크기 결정 (13~15인치 노트북 대응) ──
        # 큰 화면에서는 기존 크기, 작은 화면에서는 화면에 맞게 축소 + 중앙 배치.
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        win_w = min(1280, max(900, sw - 60))
        win_h = min(1120, max(560, sh - 100))     # 작업표시줄 + 여유 (세로 잘림 방지)
        self._win_w = win_w
        self._win_h = win_h
        pos_x = max(0, (sw - win_w) // 2)
        pos_y = max(0, min((sh - win_h) // 4, 24))
        self.root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        # 최소 크기를 낮춰 작은 화면에서도 줄여서 쓸 수 있게 함 (이전 1200x980 → 잘림 유발)
        self.root.minsize(min(880, win_w), min(540, win_h))

        try:
            ico_path = _resource_path("synth_ico.ico")
            if os.path.exists(ico_path):
                self.root.iconbitmap(ico_path)
        except Exception:
            pass

        # ── 다중 파일 상태 ──
        self.files = []                # list[FileData]
        self.relationships = []        # 자동 감지된 관계 리스트
        self.relationships_confirmed = False  # 관계 확정 게이트

        # ── 공통 설정 ──
        self.employee_id_var = tk.StringVar()
        self.save_dir = tk.StringVar()
        self.k_anon_var = tk.IntVar(value=5)

        # ── 진행 상태 ──
        self._current_step = 0
        self._blink_state = True
        self._blink_id = None

        # ── 관계 UI 위젯 ──
        self.rel_check_vars = {}       # {idx: BooleanVar}
        self.rel_check_widgets = {}    # {idx: Checkbutton} — 잠금/해제용

        # ── 노트북 탭 인덱스 (탭 활성/비활성 제어용) ──
        self.tab_rel_idx = 0
        self.tab_cols_idx = 1

        self._build_ui()
        self._set_step(0)

        # 프로젝트명/저장경로 변경 시 자동으로 run 버튼 상태 갱신
        try:
            self.employee_id_var.trace_add('write',
                lambda *a: self._update_run_state())
            self.save_dir.trace_add('write',
                lambda *a: self._update_run_state())
        except (AttributeError, tk.TclError):
            pass

    # ──────────────────────────────────────────────────────────
    # UI 구축
    # ──────────────────────────────────────────────────────────

    def _build_ui(self):
        """전체 GUI 위젯 트리 생성.

        구조:
            - 로고 배너 (있는 경우)
            - 단계 인디케이터 (5단계)
            - 상단: ① 파일 목록 + ② 공통 설정 (사번/k-익명성/분석버튼)
            - 노트북:
                · ③ 관계 설정/확정 탭
                · ④ 컬럼 편집 탭 (파일별 sub-notebook)
            - 하단: ⑤ 저장 설정 + ⑥ 실행 로그
        """
        style = ttk.Style()
        style.configure("Sub.TLabel", font=("맑은 고딕", 9), foreground="#666")
        style.configure("ColHeader.TLabel", font=("맑은 고딕", 10, "bold"), foreground="#333")
        style.configure("Run.TButton", font=("맑은 고딕", 11, "bold"))
        style.configure("Auto.TButton", font=("맑은 고딕", 9))
        style.configure("Section.TLabel", font=("맑은 고딕", 10, "bold"), foreground="#2c3e50")
        # 테이블명 Entry는 ttk가 아닌 tk.Entry로 만들어 bg를 직접 지정 (vista 테마 호환성)

        # 세로가 좁은 화면(13인치 등)에서는 섹션 패딩을 줄여 본문 공간 확보
        _pad = 8 if getattr(self, '_win_h', 1000) >= 850 else 4

        # ── 파일별 sub-tab은 FileTabBar(커스텀 tk.Button 기반)로 색상 강제 적용 ──
        # ttk.Notebook은 Windows 네이티브 테마가 탭 배경을 가로채서 색이 안 먹음.

        # ── 로고 ── (세로가 좁은 화면에서는 생략해 본문 공간 확보)
        self._logo_image = None
        try:
            logo_path = _resource_path("logo.png")
            if (os.path.exists(logo_path) and HAS_PIL
                    and getattr(self, '_win_h', 1000) >= 900):
                img = Image.open(logo_path)
                target_w = max(600, getattr(self, '_win_w', 1280) - 10)
                ratio = target_w / img.width
                target_h = int(img.height * ratio)
                img = img.resize((target_w, target_h), Image.LANCZOS)
                self._logo_image = ImageTk.PhotoImage(img)
                logo_lbl = tk.Label(self.root, image=self._logo_image, bg="#2d4a7a")
                logo_lbl.pack(fill=tk.X, padx=5, pady=(5, 0))
        except Exception:
            pass

        # ── 단계 인디케이터 ──
        self._build_step_indicator(self.root)

        # ── 하단 (실행 + 저장) 먼저 pack ──
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5)

        sec_save = ttk.LabelFrame(bottom_frame, text="  ⑤ 저장 설정  ", padding=_pad)
        sec_save.pack(fill=tk.X, pady=(2, 2))

        rs1 = ttk.Frame(sec_save); rs1.pack(fill=tk.X)
        ttk.Label(rs1, text="저장 폴더:", width=10).pack(side=tk.LEFT)
        ttk.Entry(rs1, textvariable=self.save_dir, font=("Consolas", 9)).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        ttk.Button(rs1, text="폴더 선택...", command=self._browse_dir).pack(side=tk.LEFT)

        sec_run = ttk.LabelFrame(bottom_frame, text="  ⑥ 실행 로그  ", padding=_pad)
        sec_run.pack(fill=tk.BOTH, expand=True, pady=(0, 2))

        bf = ttk.Frame(sec_run); bf.pack(fill=tk.X, pady=(0, 5))
        self.run_btn = ttk.Button(bf, text="▶  전체 합성 데이터 생성 실행 (+ MD)",
                                  command=self._run, style="Run.TButton",
                                  state='disabled')
        self.run_btn.pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(bf, mode='determinate', maximum=100, length=200)
        self.progress.pack(side=tk.LEFT, padx=(12, 0))
        self.status_lbl = ttk.Label(bf, text="", style="Sub.TLabel")
        self.status_lbl.pack(side=tk.LEFT, padx=(12, 0))

        # 세로가 좁은 화면에서는 로그 높이를 줄여 본문(노트북) 공간을 확보
        _wh = getattr(self, '_win_h', 1000)
        _log_h = 7 if _wh >= 950 else (5 if _wh >= 830 else 3)
        self.log_text = scrolledtext.ScrolledText(sec_run, height=_log_h,
                                                   font=("Consolas", 9),
                                                   state=tk.DISABLED, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # ── 상단 (파일 목록 + 공통 설정) ──
        top_frame = ttk.Frame(self.root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=(5, 0))

        # ── ① 엑셀 파일 추가 ──
        sec_files = ttk.LabelFrame(top_frame, text="  ① 엑셀 파일 추가 (여러 개 가능)  ", padding=_pad)
        sec_files.pack(fill=tk.X, pady=(0, 5))

        rfb = ttk.Frame(sec_files); rfb.pack(fill=tk.X)
        ttk.Button(rfb, text="➕ 파일 추가", command=self._add_files).pack(side=tk.LEFT)
        ttk.Button(rfb, text="➖ 선택 제거", command=self._remove_selected_file).pack(
            side=tk.LEFT, padx=(6, 0))
        ttk.Button(rfb, text="🗑 전체 초기화", command=self._reset_all).pack(
            side=tk.LEFT, padx=(6, 0))
        ttk.Label(rfb, text="  파일 목록 (클릭하여 편집할 파일 선택):",
                  style="Sub.TLabel").pack(side=tk.LEFT, padx=(15, 0))

        list_frame = ttk.Frame(sec_files); list_frame.pack(fill=tk.X, pady=(4, 0))
        _list_h = 4 if getattr(self, '_win_h', 1000) >= 880 else 3
        self.file_listbox = tk.Listbox(list_frame, height=_list_h, font=("Consolas", 9),
                                       selectmode=tk.SINGLE, activestyle='dotbox')
        self.file_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        list_sb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                                 command=self.file_listbox.yview)
        list_sb.pack(side=tk.LEFT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=list_sb.set)
        # 더블클릭: 해당 파일 탭으로 이동 (컬럼 편집 탭이 활성일 때)
        self.file_listbox.bind('<Double-Button-1>', self._on_file_double_click)

        # ── ② 공통 설정 (사번 + 행수 + k-익명성) ──
        sec_common = ttk.LabelFrame(top_frame, text="  ② 공통 설정  ", padding=_pad)
        sec_common.pack(fill=tk.X, pady=(0, 5))

        rc1 = ttk.Frame(sec_common); rc1.pack(fill=tk.X)
        ttk.Label(rc1, text="📁 프로젝트명:", width=12,
                  font=("맑은 고딕", 9, "bold"),
                  foreground="#c0392b").pack(side=tk.LEFT)
        vcmd_proj = (self.root.register(self._validate_project_name), '%P')
        emp_entry = ttk.Entry(rc1, textvariable=self.employee_id_var,
                              width=18, font=("Consolas", 10),
                              validate='key', validatecommand=vcmd_proj)
        emp_entry.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(rc1, text="(영문 소문자만)",
                  style="Sub.TLabel").pack(side=tk.LEFT)

        # 파일별 테이블명 입력 영역 (프로젝트명 옆) — _refresh_table_names_frame()로 갱신
        ttk.Separator(rc1, orient=tk.VERTICAL).pack(
            side=tk.LEFT, fill=tk.Y, padx=(12, 8))
        ttk.Label(rc1, text="📋 파일별 테이블명:", width=16,
                  font=("맑은 고딕", 9, "bold"),
                  foreground="#2c3e50").pack(side=tk.LEFT)
        self.table_names_frame = ttk.Frame(rc1)
        self.table_names_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.table_names_empty_lbl = ttk.Label(self.table_names_frame,
            text="(파일 추가 후 자동 표시)", style="Sub.TLabel")
        self.table_names_empty_lbl.pack(side=tk.LEFT)

        # 생성 행 수 입력 제거 — 항상 원본과 동일 행 수로 합성
        rc2 = ttk.Frame(sec_common); rc2.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(rc2, text="🛡️ k-익명성:", width=10).pack(side=tk.LEFT)
        ttk.Spinbox(rc2, textvariable=self.k_anon_var, from_=1, to=50,
                    increment=1, width=7).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(rc2, text="(빈도 < k 카테고리 → 'OTHER' 통합)",
                  style="Sub.TLabel").pack(side=tk.LEFT)
        ttk.Label(rc2, text="  🔒 외부공유 마스킹 모드 — 컬럼명/카테고리=무의미 코드, PII=가명, 수치·날짜=통계 보존",
                  style="Sub.TLabel").pack(side=tk.LEFT, padx=(15, 0))

        rc3 = ttk.Frame(sec_common); rc3.pack(fill=tk.X, pady=(7, 0))
        self.analyze_all_btn = ttk.Button(rc3, text="📊  전체 파일 분석",
                                           command=self._analyze_all_files)
        self.analyze_all_btn.pack(side=tk.LEFT)
        ttk.Label(rc3, text="  ← 모든 파일을 한 번에 로드/분석 후 컬럼 자동 마스킹",
                  style="Sub.TLabel").pack(side=tk.LEFT, padx=(8, 0))

        self.detect_rel_btn = ttk.Button(rc3, text="🔗 컬럼 관계 자동 감지",
                                          command=self._detect_relationships,
                                          state='disabled')
        self.detect_rel_btn.pack(side=tk.LEFT, padx=(20, 0))

        # ── 노트북 (탭): 관계설정(③) 먼저 → 컬럼편집(④)  ──
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 0))

        # ═══ 탭 ③ 관계 설정/확정 ═══
        tab_rel = ttk.Frame(self.notebook, padding=3)
        self.notebook.add(tab_rel, text="  ③ 관계 설정/확정 (먼저)  ")

        rel_top = ttk.Frame(tab_rel)
        rel_top.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(rel_top,
                  text="파일 간 컬럼 관계를 자동 감지/추천합니다. 체크된 관계는 같은 가짜값을 공유하여 관계가 보존됩니다.",
                  style="Section.TLabel").pack(side=tk.LEFT)

        rel_btns = ttk.Frame(tab_rel)
        rel_btns.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(rel_btns, text="🔗 관계 재감지",
                   command=self._detect_relationships,
                   style="Auto.TButton").pack(side=tk.LEFT)
        ttk.Button(rel_btns, text="➕ 수동 관계 추가",
                   command=self._add_manual_relationship,
                   style="Auto.TButton").pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(rel_btns, text="☑️ 추천만 선택",
                   command=lambda: self._toggle_relationships(only_suggested=True),
                   style="Auto.TButton").pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(rel_btns, text="✅ 모두 선택",
                   command=lambda: self._toggle_relationships(check_all=True),
                   style="Auto.TButton").pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(rel_btns, text="⬜ 모두 해제",
                   command=lambda: self._toggle_relationships(check_all=False),
                   style="Auto.TButton").pack(side=tk.LEFT, padx=(6, 0))

        rel_mid = ttk.Frame(tab_rel)
        rel_mid.pack(fill=tk.BOTH, expand=True)

        self.rel_canvas = tk.Canvas(rel_mid, highlightthickness=0)
        rel_sb = ttk.Scrollbar(rel_mid, orient=tk.VERTICAL,
                                command=self.rel_canvas.yview)
        self.rel_inner = ttk.Frame(self.rel_canvas)
        self.rel_inner.bind("<Configure>",
            lambda e: self.rel_canvas.configure(scrollregion=self.rel_canvas.bbox("all")))
        self._rel_cw = self.rel_canvas.create_window((0, 0),
                                                      window=self.rel_inner,
                                                      anchor="nw")
        self.rel_canvas.configure(yscrollcommand=rel_sb.set)
        self.rel_canvas.bind("<Configure>",
            lambda e: self.rel_canvas.itemconfig(self._rel_cw, width=e.width))
        self.rel_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rel_sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.rel_empty_lbl = ttk.Label(self.rel_inner,
            text="\n\n   ◀ '📊 전체 파일 분석' 후 자동으로 관계가 여기 표시됩니다.\n",
            style="Sub.TLabel")
        self.rel_empty_lbl.pack(anchor='w')

        self._bind_mousewheel(self.rel_canvas)

        # 관계 탭 하단: 관계 확정 + 수정 버튼
        # side=BOTTOM 으로 먼저 하단 공간을 확보 → 관계 목록(expand)이 위를 채워도
        # 작은 화면에서 버튼이 화면 밖으로 밀리지 않음.
        rel_bottom = ttk.Frame(tab_rel)
        rel_bottom.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))
        ttk.Separator(tab_rel, orient=tk.HORIZONTAL).pack(
            side=tk.BOTTOM, fill=tk.X)
        self.confirm_rel_btn = tk.Button(rel_bottom,
            text="  ✅  관계 확정 → 컬럼 편집으로 이동  ",
            font=("맑은 고딕", 11, "bold"), bg="#27ae60", fg="white",
            activebackground="#219a52", activeforeground="white",
            relief="raised", bd=2, padx=20, pady=6,
            command=self._confirm_relationships,
            state='disabled')
        self.confirm_rel_btn.pack(side=tk.LEFT, padx=(20, 0), pady=(5, 5))
        self.unlock_rel_btn = ttk.Button(rel_bottom,
            text="🔧 관계 수정",
            command=self._unlock_relationships,
            state='disabled')
        self.unlock_rel_btn.pack(side=tk.LEFT, padx=(10, 0), pady=(5, 5))
        self.rel_status_lbl = tk.Label(rel_bottom, text="",
            font=("맑은 고딕", 9, "bold"), fg="#666",
            bg=self.root.cget('bg'))
        self.rel_status_lbl.pack(side=tk.LEFT, padx=(20, 0))

        # ═══ 탭 ④ 컬럼 편집 (파일별 sub-notebook) ═══
        tab_cols = ttk.Frame(self.notebook, padding=3)
        self.notebook.add(tab_cols, text="  ④ 컬럼 편집 (관계 확정 후)  ")

        self.cols_top_lbl = ttk.Label(tab_cols,
            text="◀ 먼저 '③ 관계 설정/확정' 탭에서 관계를 확정한 뒤 진행하세요. "
                 "관계 그룹의 컬럼은 자동으로 같은 가짜값이 적용되어 잠금 표시(🔗)됩니다.",
            style="Section.TLabel")
        self.cols_top_lbl.pack(fill=tk.X, padx=2, pady=(0, 3))

        # 파일별 내부 탭바 (관계 확정 후 빌드) — 선택=녹색 / 비선택=파란색
        self.col_inner_notebook = FileTabBar(tab_cols)
        self.col_inner_notebook.pack(fill=tk.BOTH, expand=True)

        # 초기: 컬럼편집 탭 비활성
        self.notebook.tab(self.tab_cols_idx, state='disabled')

    # ──────────────────────────────────────────────────────────
    # 단계 표시 바
    # ──────────────────────────────────────────────────────────

    def _build_step_indicator(self, parent):
        """상단 진행 단계 표시 바를 만든다 (5단계 라벨 + 화살표 + 힌트)."""
        step_frame = tk.Frame(parent, bg="#f0f4f8", relief="ridge", bd=1)
        _ipady = 5 if getattr(self, '_win_h', 1000) >= 850 else 1
        step_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 3), ipady=_ipady, padx=5)

        inner = tk.Frame(step_frame, bg="#f0f4f8")
        inner.pack(anchor='center')

        steps = [
            "❶ 파일 추가 + 프로젝트명",
            "❷ 전체 분석",
            "❸ 관계 설정/확정",
            "❹ 컬럼 편집/확정",
            "❺ 실행 + MD",
        ]

        self.step_labels = []
        self.step_arrows = []
        for i, t in enumerate(steps):
            if i > 0:
                arrow = tk.Label(inner, text="  ➤  ",
                                  font=("맑은 고딕", 12, "bold"),
                                  bg="#f0f4f8", fg="#bbb")
                arrow.pack(side=tk.LEFT)
                self.step_arrows.append(arrow)
            lbl = tk.Label(inner, text=f"  {t}  ",
                            font=("맑은 고딕", 9, "bold"),
                            bg="#ddd", fg="#999", padx=10, pady=3,
                            relief="groove", bd=1)
            lbl.pack(side=tk.LEFT, padx=2)
            self.step_labels.append(lbl)

        # 힌트는 단계 라벨 아래 별도 줄에 배치 — 작은 화면 가로 넘침 방지
        self.step_hint = tk.Label(step_frame, text="", font=("맑은 고딕", 9, "bold"),
                                   bg="#f0f4f8", fg="#e74c3c")
        self.step_hint.pack(side=tk.TOP, pady=(2, 0))

    def _set_step(self, step_num):
        """현재 진행 단계를 갱신 — 라벨 색상/화살표 색상/힌트 텍스트 적용.

        Args:
            step_num: 0~4(현재 단계) 또는 5(완료).
        """
        self._current_step = step_num
        hints = [
            "◀ 엑셀 파일을 추가하고 프로젝트명/파일별 테이블명을 입력하세요",
            "◀ '📊 전체 파일 분석' 버튼을 클릭하세요",
            "◀ 자동 감지된 관계를 확인하고 '✅ 관계 확정 → 컬럼 편집' 버튼을 누르세요",
            "◀ 각 파일 탭에서 컬럼을 확인하고 '✅ 이 파일 확정' 버튼을 누르세요",
            "◀ '▶ 전체 합성 데이터 생성 실행' 버튼을 클릭하세요",
        ]
        for i, lbl in enumerate(self.step_labels):
            if i < step_num:
                lbl.config(bg="#27ae60", fg="white")
            elif i == step_num:
                lbl.config(bg="#3498db", fg="white")
            else:
                lbl.config(bg="#ddd", fg="#999")
        for i, arrow in enumerate(self.step_arrows):
            if i < step_num:
                arrow.config(fg="#27ae60")
            elif i == step_num:
                arrow.config(fg="#3498db")
            else:
                arrow.config(fg="#bbb")
        if step_num < len(hints):
            self.step_hint.config(text=hints[step_num])
        else:
            self.step_hint.config(text="✅ 완료!")
        self._start_blink()

    def _start_blink(self):
        """현재 단계 라벨/힌트 깜빡임 타이머를 시작 (기존 타이머 취소)."""
        if self._blink_id:
            self.root.after_cancel(self._blink_id)
        self._blink_state = True
        self._do_blink()

    def _do_blink(self):
        """650ms 주기로 현재 단계 라벨과 힌트 색상을 토글하는 콜백."""
        step = self._current_step
        if step < len(self.step_labels):
            lbl = self.step_labels[step]
            if self._blink_state:
                lbl.config(bg="#3498db", fg="white")
            else:
                lbl.config(bg="#1a6fb5", fg="#b8daef")
            self._blink_state = not self._blink_state
        if self.step_hint.cget('text'):
            cur_fg = self.step_hint.cget('fg')
            self.step_hint.config(fg="#e74c3c" if cur_fg != "#e74c3c" else "#f5b7b1")
        self._blink_id = self.root.after(650, self._do_blink)

    # ──────────────────────────────────────────────────────────
    # 파일 목록 관리
    # ──────────────────────────────────────────────────────────

    def _add_files(self):
        """엑셀 파일을 선택해 self.files에 FileData로 추가 + 시트 이름 미리 로드.

        이미 등록된 경로는 중복 추가하지 않음.
        """
        paths = filedialog.askopenfilenames(
            title="엑셀 파일 선택 (여러 개 선택 가능)",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("All", "*.*")])
        if not paths:
            return
        added = 0
        existing_paths = {fd.path for fd in self.files}
        for p in paths:
            if p in existing_paths:
                continue
            fd = FileData(p)
            try:
                sheets, _ = get_sheet_names(p)
                fd.all_sheets = sheets
                fd.sheet_name = sheets[0] if sheets else ''
            except Exception as e:
                messagebox.showwarning("시트 로드 실패",
                                       f"{os.path.basename(p)}:\n{e}")
                continue
            self.files.append(fd)
            added += 1

        if not added:
            return

        if not self.save_dir.get().strip():
            self.save_dir.set(os.path.dirname(self.files[0].path))

        self._refresh_file_listbox()
        self.status_lbl.config(text=f"파일 {len(self.files)}개 등록됨 (이번에 +{added})")
        if self._current_step < 1:
            self._set_step(1)

    def _refresh_file_listbox(self):
        """파일 Listbox 표시 갱신 — 분석/확정 상태 아이콘 + 시트명 + 행수 포함."""
        self.file_listbox.delete(0, tk.END)
        for i, fd in enumerate(self.files):
            mark = "✅" if fd.confirmed else ("🔍" if fd.analyzed else "⬜")
            sheet = f"  [{fd.sheet_name}]" if fd.sheet_name else ""
            row_info = f"  ({len(fd.df)}행)" if fd.df is not None else ""
            self.file_listbox.insert(tk.END,
                f"  {mark}  {i+1:2d}. {fd.file_name}{sheet}{row_info}")
        self._refresh_table_names_frame()

    # 테이블명 입력란 한 줄에 배치할 최대 개수 (초과 시 다음 줄로 자동 줄바꿈)
    _TABLE_CELLS_PER_ROW = 6

    def _refresh_table_names_frame(self):
        """프로젝트명 옆 파일별 테이블명 입력 영역을 self.files 기준으로 재구성.

        각 파일마다 (파일번호 + 작은 Entry)를 grid로 배치하되,
        _TABLE_CELLS_PER_ROW 개를 넘으면 다음 줄로 줄바꿈한다.
        (가로 한 줄 배치 시 파일이 많으면 창 밖으로 잘려 보이던 문제 해결)
        영문 소문자만 입력 가능.
        """
        # 기존 위젯 제거
        for w in self.table_names_frame.winfo_children():
            w.destroy()
        if not self.files:
            self.table_names_empty_lbl = ttk.Label(self.table_names_frame,
                text="(파일 추가 후 자동 표시)", style="Sub.TLabel")
            self.table_names_empty_lbl.grid(row=0, column=0, sticky='w')
            return
        vcmd = (self.root.register(self._validate_table_name), '%P')
        per_row = self._TABLE_CELLS_PER_ROW
        for i, fd in enumerate(self.files):
            r, c = divmod(i, per_row)
            cell = ttk.Frame(self.table_names_frame)
            cell.grid(row=r, column=c, sticky='w', padx=(0, 6), pady=(0, 2))
            ttk.Label(cell, text=f"#{i+1}",
                      font=("맑은 고딕", 9, "bold"),
                      foreground="#2c3e50").pack(side=tk.LEFT, padx=(0, 2))
            var = tk.StringVar(value=fd.table_name)
            fd._table_name_var = var
            # tk.Entry(기본 Tk 위젯) — Windows ttk 테마에서도 bg 지정이 확실히 적용됨
            entry = tk.Entry(cell, textvariable=var, width=14,
                             font=("Consolas", 10),
                             validate='key', validatecommand=vcmd,
                             relief='solid', bd=1, bg='white',
                             highlightthickness=1,
                             highlightbackground='#bbb',
                             highlightcolor='#2c3e50')
            entry.pack(side=tk.LEFT)
            fd._db_check_entry = entry

            # 입력이 바뀌면 상태 초기화(흰 배경) — 포커스 빠지면 다시 검사
            def _on_write(*_a, f=fd, v=var, e=entry):
                f.table_name = v.get().strip()
                f._db_check_status = None
                try:
                    e.configure(bg='white')
                except tk.TclError:
                    pass
                self._update_run_state()
            var.trace_add('write', _on_write)

            # 위젯 재구성으로 인해 상태가 사라지지 않도록 저장된 상태 복원
            self._apply_db_check_style(fd)

            # 포커스 떠날 때 DB 존재 검사 — 백그라운드 스레드로 UI 비차단
            entry.bind('<FocusOut>',
                lambda _e, f=fd: self._check_table_in_db_async(f))

    def _remove_selected_file(self):
        """Listbox에서 선택한 파일을 제거하고 관계/컬럼 탭 상태도 초기화.

        파일 인덱스가 바뀌므로 기존 관계 정보는 모두 폐기.
        """
        sel = self.file_listbox.curselection()
        if not sel:
            messagebox.showinfo("알림", "목록에서 제거할 파일을 선택하세요.")
            return
        idx = sel[0]
        del self.files[idx]
        # 관계는 인덱스가 바뀌므로 초기화 + 컬럼 편집 탭도 초기화
        self.relationships = []
        self.relationships_confirmed = False
        self._refresh_relationship_widgets()
        self._clear_col_inner_tabs()
        self.notebook.tab(self.tab_cols_idx, state='disabled')
        self.confirm_rel_btn.config(state='disabled' if not self.files else 'normal')
        self.unlock_rel_btn.config(state='disabled')
        self._refresh_file_listbox()
        self._update_run_state()

    def _reset_all(self):
        """전체 초기화 — 파일/관계/컬럼탭/로그/진행상태를 처음 상태로 되돌림."""
        if self.files and not messagebox.askyesno("확인",
                "모든 파일과 설정을 초기화하시겠습니까?"):
            return
        self._clear_col_inner_tabs()
        self.files = []
        self.relationships = []
        self.relationships_confirmed = False
        self._refresh_file_listbox()
        self._refresh_relationship_widgets()
        self.notebook.tab(self.tab_cols_idx, state='disabled')
        self.notebook.select(self.tab_rel_idx)
        self.confirm_rel_btn.config(state='disabled')
        self.unlock_rel_btn.config(state='disabled')
        self.rel_status_lbl.config(text="", fg="#666")
        self.cols_top_lbl.config(
            text="◀ 먼저 '③ 관계 설정/확정' 탭에서 관계를 확정한 뒤 진행하세요.")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.progress['value'] = 0
        self.run_btn.config(state='disabled')
        self.detect_rel_btn.config(state='disabled')
        self.status_lbl.config(text="")
        self._set_step(0)

    def _on_file_double_click(self, event=None):
        """Listbox 더블클릭 → 컬럼 편집 탭의 해당 파일 sub-tab으로 이동."""
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self.files):
            return
        fd = self.files[idx]
        if not self.relationships_confirmed or fd.tab_frame is None:
            return
        try:
            tab_id = self.col_inner_notebook.tabs()[idx]
            self.col_inner_notebook.select(tab_id)
            self.notebook.select(self.tab_cols_idx)
        except (tk.TclError, IndexError):
            pass

    def _clear_col_inner_tabs(self):
        """컬럼 편집 탭의 모든 sub-tab을 제거."""
        self.col_inner_notebook.clear()
        for fd in self.files:
            fd.tab_frame = None
            fd.rename_inner = None
            fd.input_inner = None
            fd.confirm_btn = None
            fd.confirm_status_lbl = None
            fd.analysis_text = None
            fd.rename_entries = {}
            fd.input_entry_map = {}
            fd.mode_vars = {}
            fd.desc_entries = {}

    # ──────────────────────────────────────────────────────────
    # 전체 분석
    # ──────────────────────────────────────────────────────────

    def _analyze_all_files(self):
        """모든 등록된 파일을 일괄 분석.

        - Excel 로드
        - 컬럼 타입 추론 (numerical/datetime/categorical)
        - PII / ID 컬럼 자동 감지
        - 파일 간 충돌 없는 unique 마스킹 코드 생성
        - 카테고리 가짜값 매핑 사전 생성
        - 분석 완료 후 관계 자동 감지 + 관계 설정 탭으로 자동 전환
        """
        if not self.files:
            messagebox.showwarning("경고", "먼저 엑셀 파일을 추가하세요.")
            return
        if not self.employee_id_var.get().strip():
            if not messagebox.askyesno("프로젝트명 미입력",
                    "프로젝트명이 비어 있습니다. 그래도 분석을 진행하시겠습니까?\n"
                    "(실행 전에는 프로젝트명이 반드시 필요합니다.)"):
                return

        # 중복 실행 방지
        if getattr(self, '_analyzing', False):
            return
        self._analyzing = True
        try:
            self.analyze_all_btn.config(state='disabled')
        except tk.TclError:
            pass
        self.status_lbl.config(text="전체 파일 분석 준비 중...")
        self.progress['value'] = 0
        self.root.update_idletasks()

        # ── 이미 분석된 파일의 마스킹 코드를 수집 (전역 충돌 방지) — 메인 스레드 ──
        all_used_codes = set()
        for fd in self.files:
            if fd.analyzed and fd.column_mask_map:
                all_used_codes.update(str(v) for v in fd.column_mask_map.values())

        n_targets = len([fd for fd in self.files if not fd.analyzed])
        # 무거운 로드/분석/관계감지는 백그라운드 스레드에서 수행 (UI 멈춤 방지)
        threading.Thread(target=self._analyze_worker,
                         args=(all_used_codes, n_targets),
                         daemon=True).start()

    def _analyze_worker(self, all_used_codes, n_targets):
        """[백그라운드] 파일 로드 + 타입/PII/ID 추론 + 마스킹코드 + 매핑 + 관계감지.

        Tkinter 위젯에는 절대 접근하지 않는다 (스레드 안전).
        진행/완료 UI 갱신은 root.after로 메인 스레드에 위임한다.
        """
        errors = []
        loaded = 0
        done = 0
        try:
            for fd in self.files:
                if fd.analyzed:
                    loaded += 1
                    continue
                done += 1
                self.root.after(0,
                    lambda d=done, n=n_targets, nm=fd.file_name:
                        self._analyze_progress(d, n, nm))
                try:
                    (fd.df, fd.info), _engine = load_excel(
                        fd.path, fd.sheet_name or None)
                except Exception as e:
                    errors.append(f"{fd.file_name}:\n{e}")
                    continue
                fd.original_columns = list(fd.df.columns)

                fd.col_types = {}
                for col in fd.df.columns:
                    fd.col_types[col] = auto_detect_column_type(fd.df[col])

                fd.pii_auto_detected = detect_pii_columns(fd.df, fd.col_types)
                fd.id_suspected = detect_id_columns(fd.df, fd.col_types)

                seed = abs(hash(fd.file_name)) % 999999
                fd.column_mask_map = generate_unique_masked_column_names(
                    list(fd.df.columns), excluded=all_used_codes, seed=seed)
                all_used_codes.update(fd.column_mask_map.values())
                fd.rename_values = dict(fd.column_mask_map)

                fd.col_modes = {}
                fd.col_descriptions = {}
                for col in fd.df.columns:
                    if fd.col_types.get(col) == 'categorical':
                        fd.col_modes[col] = (
                            'fake' if col in fd.pii_auto_detected else 'masked')

                fd.mapping_values = self._compute_auto_mappings(fd)
                fd.analyzed = True
                loaded += 1

            # 관계 자동 감지 (순수 계산) — 백그라운드에서 수행
            if len([f for f in self.files if f.analyzed]) >= 2:
                self.root.after(0, lambda: self.status_lbl.config(
                    text="컬럼 관계 자동 감지 중..."))
                relationships = detect_relationships(self.files)
            else:
                relationships = []
        except Exception as e:
            errors.append(f"분석 중 오류: {e}")
            relationships = []

        self.root.after(0,
            lambda: self._analyze_finish(loaded, errors, relationships))

    def _analyze_progress(self, done, total, name):
        """[메인] 분석 진행률 표시."""
        try:
            self.progress['value'] = int(done / max(total, 1) * 100)
            self.status_lbl.config(text=f"분석 중... {done}/{total}  ({name})")
        except tk.TclError:
            pass

    def _analyze_finish(self, loaded, errors, relationships):
        """[메인] 분석 완료 후 UI 갱신 (위젯 빌드/관계탭/탭전환 등)."""
        self._analyzing = False
        try:
            self.analyze_all_btn.config(state='normal')
        except tk.TclError:
            pass
        self.progress['value'] = 0
        for err in errors:
            messagebox.showerror("로드 오류", err)

        self._refresh_file_listbox()
        self.detect_rel_btn.config(state='normal')

        # 관계 확정 상태 초기화 (재분석 시)
        self.relationships_confirmed = False
        self._clear_col_inner_tabs()
        self.notebook.tab(self.tab_cols_idx, state='disabled')

        # ── 관계 자동 감지 결과 반영 ──
        if len([f for f in self.files if f.analyzed]) >= 2:
            self.relationships = relationships
            self._refresh_relationship_widgets()
            n_sugg = sum(1 for r in self.relationships if r.get('enabled'))
            self.rel_status_lbl.config(
                text=f"{len(self.relationships)}개 발견 (추천 {n_sugg}개)",
                fg="#666")
            self.confirm_rel_btn.config(state='normal')
            self.unlock_rel_btn.config(state='disabled')
        else:
            # 파일 1개라면 관계 확정 단계를 스킵 가능하도록 버튼만 enable
            self.relationships = []
            self._refresh_relationship_widgets()
            self.rel_status_lbl.config(
                text="(파일 1개 — 관계 없음, 바로 확정 가능)", fg="#666")
            self.confirm_rel_btn.config(state='normal')

        # ── 관계가 전혀 감지되지 않으면 확정 버튼을 숨기고 자동으로 다음 단계로 ──
        if not self.relationships:
            try:
                self.confirm_rel_btn.pack_forget()
            except tk.TclError:
                pass
            self.rel_status_lbl.config(
                text="관계 없음 — 바로 컬럼 편집으로 진행합니다.",
                fg="#27ae60")
            # 컬럼 편집 단계로 자동 진입 (관계 확정 게이트 통과)
            self._confirm_relationships(auto_skip=True)
            self.status_lbl.config(
                text=f"분석 완료 — {loaded}/{len(self.files)}개. "
                     f"관계 없음 → 컬럼 편집 단계로 이동")
            return

        # 관계가 있을 때만 관계 탭 보여주기
        try:
            self.confirm_rel_btn.pack(side=tk.LEFT, padx=(20, 0), pady=(5, 5))
        except tk.TclError:
            pass
        self.notebook.select(self.tab_rel_idx)

        self.status_lbl.config(
            text=f"분석 완료 — {loaded}/{len(self.files)}개. 관계 확인 후 확정하세요.")
        self._set_step(2)

    def _compute_auto_mappings(self, fd):
        """가짜 데이터 자동 매핑 생성. {col: {orig_val: fake_val}}"""
        mappings = {}
        str_cols = [c for c in fd.df.columns
                    if (fd.df[c].dtype == object or pd.api.types.is_string_dtype(fd.df[c]))
                    and fd.col_types.get(c) == 'categorical']
        for col in str_cols:
            series = fd.df[col].dropna()
            if len(series) == 0:
                continue
            unique_vals = sorted([str(v) for v in series.unique() if v is not None])
            n = len(unique_vals)
            seed = abs(hash((fd.file_name, str(col)))) % 999999

            is_pii = col in fd.pii_auto_detected
            is_id = col in fd.id_suspected
            looks_like_name = series.astype(str).str.match(
                r'^[가-힣]{2,4}$').mean() > 0.7
            looks_like_company = str(col).lower() in [
                'client', '고객', '발주', '선주', 'company', '업체', '회사', '회사명']

            if is_id:
                orig_prefix = _extract_id_prefix(unique_vals)
                fake_prefix = _pick_different_prefix(orig_prefix)
                pool = [fake_prefix + str(i + 1).zfill(4) for i in range(n)]
            elif is_pii and looks_like_name:
                pool = generate_fake_persons(n, seed=seed)
            elif is_pii and looks_like_company:
                pool = generate_fake_companies(n, seed=seed)
            elif is_pii:
                pool = generate_fake_persons(n, seed=seed)
            else:
                col_idx = list(fd.df.columns).index(col)
                col_alpha = chr(ord('A') + (col_idx % 26))
                code_map = generate_masked_category_codes(unique_vals, col_alpha)
                pool = [code_map[v] for v in unique_vals]

            mappings[col] = {v: pool[i] if i < len(pool) else f"CAT_X_{i+1:03d}"
                              for i, v in enumerate(unique_vals)}
        return mappings

    # ──────────────────────────────────────────────────────────
    # 선택 파일을 편집 영역에 로드 / 저장
    # ──────────────────────────────────────────────────────────

    # ──────────────────────────────────────────────────────────
    # 파일별 sub-tab 빌드 (관계 확정 후 호출됨)
    # ──────────────────────────────────────────────────────────

    def _build_file_tab(self, fd, tab_idx):
        """각 파일을 위한 sub-tab을 col_inner_notebook에 추가하고 위젯을 빌드."""
        tab_label = self._file_tab_label(fd, tab_idx)
        tab = self.col_inner_notebook.new_tab(tab_label)
        fd.tab_frame = tab

        # 분석 요약 텍스트 (세로 좁은 화면은 2줄로 축소)
        _sum_h = 3 if getattr(self, '_win_h', 1000) >= 850 else 2
        fd.analysis_text = tk.Text(tab, height=_sum_h, font=("Consolas", 9),
                                    bg="#f5f5f0", state=tk.DISABLED, wrap=tk.WORD)
        fd.analysis_text.pack(fill=tk.X, pady=(0, 4))

        # 확정 버튼 — 요약 바로 아래(스크롤 영역 위)에 고정 배치.
        # 컬럼이 많거나 창이 작아도 항상 보이도록 상단에 둔다.
        confirm_frame = ttk.Frame(tab)
        confirm_frame.pack(fill=tk.X, pady=(0, 4), padx=2)
        fd.confirm_btn = tk.Button(confirm_frame,
            text="  ✅  이 파일 변환 계획 확정  ",
            font=("맑은 고딕", 11, "bold"), bg="#27ae60", fg="white",
            activebackground="#219a52", activeforeground="white",
            relief="raised", bd=2, padx=20, pady=6,
            command=lambda f=fd: self._confirm_current_file(f))
        fd.confirm_btn.pack(side=tk.LEFT, padx=(0, 10))
        fd.confirm_status_lbl = tk.Label(confirm_frame, text="",
            font=("맑은 고딕", 9, "bold"), fg="#e67e22")
        fd.confirm_status_lbl.pack(side=tk.LEFT)
        ttk.Separator(tab, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 4))

        # 스크롤 영역
        mid = ttk.Frame(tab)
        mid.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(mid, highlightthickness=0)
        sb = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>",
            lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        cw = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.bind("<Configure>",
            lambda e, c=canvas, w=cw: c.itemconfig(w, width=e.width))
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self._bind_mousewheel(canvas)

        # 컬럼명 변경 + 설명 영역
        sec_rename = ttk.LabelFrame(inner,
            text="  ④ 컬럼명 변경 / 설명 입력 (가짜값은 자동 생성됨)  ", padding=8)
        sec_rename.pack(fill=tk.X, pady=(0, 5), padx=2)
        rt = ttk.Frame(sec_rename); rt.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(rt,
                  text="새 이름을 비워두면 원본 유지. 마스킹 코드 그대로가 안전합니다.",
                  style="Sub.TLabel").pack(side=tk.LEFT)
        ttk.Button(rt, text="✅ 컬럼명 적용",
                   command=lambda f=fd: self._apply_column_rename(f),
                   style="Auto.TButton").pack(side=tk.RIGHT)

        # AI 채팅 복붙 도우미 — 컬럼명 CSV 복사 + AI 답변 일괄 붙여넣기
        # (자동 연동 아님 — Claude/ChatGPT 채팅창에 사람이 직접 복붙하는 용도)
        ai_bar = ttk.Frame(sec_rename); ai_bar.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(ai_bar, text="💬 AI 채팅 복붙 도우미:",
                  style="Sub.TLabel").pack(side=tk.LEFT)
        ttk.Button(ai_bar, text="📋 컬럼+질문 프롬프트 복사",
                   command=lambda f=fd: self._copy_columns_csv(f),
                   style="Auto.TButton").pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(ai_bar, text="📥 AI 답변 설명 일괄 붙여넣기",
                   command=lambda f=fd: self._paste_descriptions_dialog(f),
                   style="Auto.TButton").pack(side=tk.LEFT, padx=(6, 0))

        fd.rename_inner = ttk.Frame(sec_rename)
        fd.rename_inner.pack(fill=tk.X)

        # 가짜값(원본 데이터) 입력 영역은 표시하지 않음 — 대용량 데이터에서 UI 멈춤 방지.
        # fd.mapping_values는 분석 단계의 _compute_auto_mappings에서 이미 자동 채워졌고,
        # 합성 단계에서 그대로 사용된다.
        fd.input_inner = None

        # 컨텐츠 채우기
        self._populate_file_tab(fd)

    def _file_tab_label(self, fd, idx):
        """파일별 sub-tab의 라벨 텍스트 — 확정 여부 아이콘 + 번호 + 파일명."""
        mark = "✅" if fd.confirmed else "⬜"
        return f"  {mark} {idx+1}. {fd.file_name}  "

    def _update_file_tab_label(self, fd):
        """이 파일의 sub-tab 라벨 갱신 (확정 상태 반영)."""
        if fd.tab_frame is None:
            return
        try:
            idx = self.files.index(fd)
            tab_id = self.col_inner_notebook.tabs()[idx]
            self.col_inner_notebook.tab(tab_id, text=self._file_tab_label(fd, idx))
        except (ValueError, tk.TclError, IndexError):
            pass

    def _populate_file_tab(self, fd):
        """파일 sub-tab의 분석요약/rename/input 위젯을 채우거나 갱신."""
        if fd.analysis_text is None:
            return

        # 분석 요약
        type_counts = {'numerical': 0, 'datetime': 0, 'categorical': 0}
        for t in fd.col_types.values():
            type_counts[t] = type_counts.get(t, 0) + 1
        icons = {'numerical': '📐수치', 'datetime': '📅날짜', 'categorical': '🏷️범주'}
        col_list = "  |  ".join(
            [f"{icons[t]} {str(c)}" for c, t in fd.col_types.items()])
        pii_line = ""
        if fd.pii_auto_detected:
            pii_line = f"\n🔒 PII (잠금): {', '.join(str(c) for c in fd.pii_auto_detected)}"
        id_line = ""
        if fd.id_suspected:
            id_line = f"\n🆔 ID 의심: {', '.join(str(c) for c in fd.id_suspected)}"
        link_line = ""
        if fd.linked_cols:
            link_line = ("\n🔗 관계 그룹 컬럼 (자동 공유 매핑/잠금): "
                          + ", ".join(f"{c}(G{g+1})"
                                       for c, g in fd.linked_cols.items()))
        summary = (
            f"파일: {fd.file_name}  |  시트: {fd.sheet_name}  |  "
            f"{len(fd.df)}행 × {len(fd.df.columns)}열  "
            f"(수치 {type_counts['numerical']}, 날짜 {type_counts['datetime']}, "
            f"범주 {type_counts['categorical']})\n"
            f"{col_list}{pii_line}{id_line}{link_line}")
        fd.analysis_text.config(state=tk.NORMAL)
        fd.analysis_text.delete("1.0", tk.END)
        fd.analysis_text.insert(tk.END, summary)
        fd.analysis_text.config(state=tk.DISABLED)

        self._build_rename_widgets(fd)
        # _build_input_widgets는 대용량 데이터에서 UI 멈춤을 유발하므로 호출하지 않음.
        # 가짜값 매핑은 _compute_auto_mappings에서 이미 생성됨.

        if fd.confirmed:
            fd.confirm_btn.config(state='disabled', bg="#95a5a6",
                                   text="  ✅  확정 완료  ")
            fd.confirm_status_lbl.config(
                text="✅ 이 파일이 확정되었습니다. 다른 파일 탭으로 이동하세요.",
                fg="#27ae60")
        else:
            fd.confirm_btn.config(state='normal', bg="#27ae60",
                                   text="  ✅  이 파일 변환 계획 확정  ")
            fd.confirm_status_lbl.config(
                text="▲ 컬럼명/가짜값을 확인한 뒤 위 버튼을 눌러 이 파일을 확정하세요.",
                fg="#e67e22")

    def _attach_placeholder(self, entry, placeholder_text):
        """ttk.Entry에 placeholder 동작 부착 — 빈 상태일 때 회색 안내 문구 표시.

        entry._is_placeholder 플래그로 현재 표시 중인 텍스트가 사용자 입력인지
        placeholder인지 구분 — _save_current_edits에서 placeholder 저장 방지.
        """
        # 이미 값이 있으면 placeholder 비활성
        if entry.get():
            entry._is_placeholder = False
            return

        def _show():
            entry.delete(0, tk.END)
            entry.insert(0, placeholder_text)
            try:
                entry.configure(foreground="#999")
            except tk.TclError:
                pass
            entry._is_placeholder = True

        def _hide():
            if getattr(entry, '_is_placeholder', False):
                entry.delete(0, tk.END)
                try:
                    entry.configure(foreground="black")
                except tk.TclError:
                    pass
                entry._is_placeholder = False

        def _on_focus_in(_e):
            _hide()

        def _on_focus_out(_e):
            if not entry.get():
                _show()

        _show()
        entry.bind("<FocusIn>", _on_focus_in)
        entry.bind("<FocusOut>", _on_focus_out)

    def _save_current_edits(self, fd):
        """이 파일 sub-tab의 위젯 값을 FileData에 저장."""
        if not fd.analyzed:
            return
        # 컬럼명 변경 값
        for orig, entry in fd.rename_entries.items():
            try:
                v = entry.get().strip()
            except tk.TclError:
                continue
            if v:
                fd.rename_values[orig] = v
            else:
                fd.rename_values.pop(orig, None)
        # 가짜값 매핑 (관계 그룹 컬럼은 fixed로 저장돼있어 entry가 None)
        for col, entries in fd.input_entry_map.items():
            mapping = {}
            for val, entry, fixed in entries:
                if fixed is not None:
                    mapping[val] = fixed
                elif entry is not None:
                    try:
                        f = entry.get().strip()
                    except tk.TclError:
                        f = ''
                    if f:
                        mapping[val] = f
            if mapping:
                fd.mapping_values[col] = mapping
        # 컬럼 모드
        for col, mvar in fd.mode_vars.items():
            try:
                fd.col_modes[col] = mvar.get()
            except tk.TclError:
                pass
        # 컬럼 설명 — placeholder 텍스트는 빈 값으로 저장
        for col, entry in fd.desc_entries.items():
            try:
                val = entry.get().strip()
            except tk.TclError:
                continue
            if getattr(entry, '_is_placeholder', False):
                val = ''
            fd.col_descriptions[col] = val

    def _build_rename_widgets(self, fd):
        """fd.rename_inner에 컬럼명 변경 위젯을 빌드."""
        for w in fd.rename_inner.winfo_children():
            w.destroy()
        fd.rename_entries = {}
        fd.mode_vars = {}
        fd.desc_entries = {}

        header = ttk.Frame(fd.rename_inner)
        header.pack(fill=tk.X, padx=5, pady=(2, 4))
        ttk.Label(header, text="No.", width=4,
                  font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="기존 컬럼명", width=16,
                  font=("맑은 고딕", 9, "bold"),
                  anchor='w').pack(side=tk.LEFT, padx=(3, 0))
        ttk.Label(header, text="→", width=2,
                  font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="마스킹 코드", width=14,
                  font=("맑은 고딕", 9, "bold"),
                  anchor='w').pack(side=tk.LEFT)
        ttk.Label(header, text="타입", width=12,
                  font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="처리방식", width=18,
                  font=("맑은 고딕", 9, "bold"),
                  anchor='w').pack(side=tk.LEFT, padx=(3, 0))
        ttk.Label(header, text="설명 (MD/외부 AI용)",
                  font=("맑은 고딕", 9, "bold"),
                  anchor='w').pack(side=tk.LEFT, padx=(3, 0))
        ttk.Separator(fd.rename_inner, orient=tk.HORIZONTAL).pack(
            fill=tk.X, padx=5)

        icons = {'numerical': '📐', 'datetime': '📅', 'categorical': '🏷️'}
        for idx, col in enumerate(fd.df.columns):
            row = ttk.Frame(fd.rename_inner)
            row.pack(fill=tk.X, padx=5, pady=1)

            ttk.Label(row, text=f"{idx+1}", width=4, font=("Consolas", 9),
                      foreground="#999").pack(side=tk.LEFT)
            ttk.Label(row, text=str(col)[:18], width=16, anchor='w',
                      font=("Consolas", 9)).pack(side=tk.LEFT, padx=(3, 0))
            ttk.Label(row, text="→", width=2).pack(side=tk.LEFT)

            entry = ttk.Entry(row, font=("Consolas", 9), width=14)
            entry.pack(side=tk.LEFT)
            cur = fd.rename_values.get(col, fd.column_mask_map.get(col, ''))
            entry.insert(0, cur)
            fd.rename_entries[col] = entry

            col_type = fd.col_types.get(col, 'categorical')
            icon = icons.get(col_type, '🏷️')
            is_pii = col in fd.pii_auto_detected
            is_id = col in fd.id_suspected
            is_linked = col in fd.linked_cols
            badges = []
            if is_linked:
                badges.append(f"🔗G{fd.linked_cols[col]+1}")
            if is_pii:
                badges.append("🔒")
            elif is_id:
                badges.append("🆔")
            badge_str = " " + " ".join(badges) if badges else ""
            if is_linked:
                type_fg = "#8e44ad"
            elif is_pii:
                type_fg = "#c0392b"
            elif is_id:
                type_fg = "#2980b9"
            else:
                type_fg = "#888"
            ttk.Label(row, text=f"{icon}{col_type}{badge_str}", width=13,
                      font=("맑은 고딕", 8),
                      foreground=type_fg).pack(side=tk.LEFT, padx=(3, 0))

            mode_frame = ttk.Frame(row, width=160)
            mode_frame.pack(side=tk.LEFT, padx=(3, 0))
            mode_frame.pack_propagate(False)
            if col_type == 'categorical':
                default_mode = fd.col_modes.get(col,
                    'fake' if is_pii else 'masked')
                mvar = tk.StringVar(value=default_mode)
                fd.mode_vars[col] = mvar
                if is_linked:
                    ttk.Label(mode_frame, text="🔗 관계공유",
                              font=("맑은 고딕", 8, "italic"),
                              foreground="#8e44ad").pack(side=tk.LEFT)
                elif is_pii:
                    ttk.Label(mode_frame, text="🔒 PII가명",
                              font=("맑은 고딕", 8, "italic"),
                              foreground="#c0392b").pack(side=tk.LEFT)
                else:
                    for label, val in [("마스킹", 'masked'), ("원본", 'keep')]:
                        rb = ttk.Radiobutton(mode_frame, text=label,
                                              variable=mvar, value=val)
                        rb.pack(side=tk.LEFT, padx=(0, 2))
            else:
                ttk.Label(mode_frame, text="(통계보존 자동)",
                          font=("맑은 고딕", 8, "italic"),
                          foreground="#888").pack(side=tk.LEFT)

            desc_entry = ttk.Entry(row, font=("맑은 고딕", 9))
            desc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(3, 0))
            saved_desc = fd.col_descriptions.get(col, '')
            if saved_desc:
                desc_entry.insert(0, saved_desc)
            self._attach_placeholder(desc_entry,
                "Sola Ai를 사용하여 해당컬럼을 외부로 반출시 설명할 문구를 작성해서 입력해주세요")
            fd.desc_entries[col] = desc_entry

    def _build_input_widgets(self, fd):
        """fd.input_inner에 가짜값 입력 위젯을 빌드.
        관계 그룹 컬럼은 read-only 라벨로 표시."""
        for w in fd.input_inner.winfo_children():
            w.destroy()
        fd.input_entry_map = {}

        str_cols = [c for c in fd.df.columns
                    if (fd.df[c].dtype == object
                        or pd.api.types.is_string_dtype(fd.df[c]))
                    and fd.col_types.get(c) == 'categorical']
        if not str_cols:
            ttk.Label(fd.input_inner,
                      text="  입력이 필요한 문자열 컬럼이 없습니다.",
                      style="Sub.TLabel").pack(anchor='w', pady=10)
            return

        for col_idx, col in enumerate(str_cols):
            series = fd.df[col].dropna()
            is_id = _is_id_col(series)
            is_linked = col in fd.linked_cols
            unique_vals = sorted([str(v) for v in series.unique() if v is not None])
            n = len(unique_vals)

            col_frame = ttk.Frame(fd.input_inner)
            col_frame.pack(fill=tk.X, pady=(8 if col_idx > 0 else 0, 2))
            title = f"📋 {col}  ({n}개)"
            ttk.Label(col_frame, text=title,
                      style="ColHeader.TLabel").pack(side=tk.LEFT)
            if is_linked:
                ttk.Label(col_frame,
                          text=f"  🔗 관계 그룹 G{fd.linked_cols[col]+1} — 공유 매핑/잠금",
                          font=("맑은 고딕", 9, "italic"),
                          foreground="#8e44ad").pack(side=tk.LEFT, padx=(8, 0))
            elif is_id:
                ttk.Label(col_frame, text="  (ID형 — 자동 재생성)",
                          style="Sub.TLabel").pack(side=tk.LEFT, padx=(8, 0))
            else:
                ttk.Button(col_frame, text="🔄 자동 채우기",
                           command=lambda f=fd, c=col: self._auto_fill_column(f, c),
                           style="Auto.TButton").pack(side=tk.RIGHT)

            ttk.Separator(fd.input_inner, orient=tk.HORIZONTAL).pack(
                fill=tk.X, pady=1)

            entries = []
            saved_mapping = fd.mapping_values.get(col, {})

            if is_linked:
                # 관계 그룹 컬럼: 모든 값을 read-only 라벨로 표시 (공유 매핑 잠금)
                for val in unique_vals:
                    row = ttk.Frame(fd.input_inner)
                    row.pack(fill=tk.X, padx=(20, 0), pady=1)
                    ttk.Label(row, text=f"{str(val)[:30]}", width=25,
                              anchor='w',
                              font=("Consolas", 9)).pack(side=tk.LEFT)
                    ttk.Label(row, text="→", width=3).pack(side=tk.LEFT)
                    fake_val = saved_mapping.get(val, '?')
                    ttk.Label(row, text=fake_val, font=("Consolas", 9, "bold"),
                              foreground="#8e44ad").pack(side=tk.LEFT)
                    entries.append((val, None, fake_val))
            elif is_id:
                orig_prefix = _extract_id_prefix(unique_vals)
                fake_prefix = _pick_different_prefix(orig_prefix)
                for i, val in enumerate(unique_vals):
                    row = ttk.Frame(fd.input_inner)
                    row.pack(fill=tk.X, padx=(20, 0), pady=1)
                    ttk.Label(row, text=f"{val}", width=25, anchor='w',
                              font=("Consolas", 9)).pack(side=tk.LEFT)
                    ttk.Label(row, text="→", width=3).pack(side=tk.LEFT)
                    fake_val = saved_mapping.get(val,
                        fake_prefix + str(i + 1).zfill(4))
                    ttk.Label(row, text=fake_val, font=("Consolas", 9),
                              foreground="#2266aa").pack(side=tk.LEFT)
                    entries.append((val, None, fake_val))
            else:
                for val in unique_vals:
                    row = ttk.Frame(fd.input_inner)
                    row.pack(fill=tk.X, padx=(20, 0), pady=1)
                    ttk.Label(row, text=f"{str(val)[:30]}", width=25,
                              anchor='w',
                              font=("Consolas", 9)).pack(side=tk.LEFT)
                    ttk.Label(row, text="→", width=3).pack(side=tk.LEFT)
                    entry = ttk.Entry(row, font=("Consolas", 9), width=30)
                    entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                    if val in saved_mapping:
                        entry.insert(0, saved_mapping[val])
                    entries.append((val, entry, None))

            fd.input_entry_map[col] = entries

    def _auto_fill_column(self, fd, col):
        """단일 컬럼의 비어있는 가짜값 Entry를 자동 채우기.

        관계 그룹 컬럼은 공유 매핑이 이미 적용돼 있으므로 건너뜀.
        """
        if col not in fd.input_entry_map:
            return
        if col in fd.linked_cols:
            return  # 관계 그룹 컬럼은 read-only
        new_mapping = self._compute_auto_mappings(fd).get(col, {})
        for val, entry, fixed in fd.input_entry_map[col]:
            if entry is not None and not entry.get().strip():
                if val in new_mapping:
                    entry.delete(0, tk.END)
                    entry.insert(0, new_mapping[val])

    def _auto_fill_all(self, fd):
        """이 파일의 모든 비어있는 가짜값 Entry를 자동 채우기."""
        for col in fd.input_entry_map:
            self._auto_fill_column(fd, col)
        self.status_lbl.config(
            text=f"{fd.file_name}: 가짜값 자동 채우기 완료")

    # ── AI 채팅 복붙 도우미: 컬럼명 CSV 복사 / 답변 일괄 붙여넣기 ──────
    #   자동 연동(API) 아님 — Claude/ChatGPT 등 채팅창에 사람이 직접
    #   복사·붙여넣기 하는 수작업을 편하게 해주는 도우미일 뿐이다.

    def _build_ai_prompt(self, fd):
        """이 파일의 컬럼 목록 + AI 질문 프롬프트를 하나의 텍스트로 조립.

        AI가 '컬럼명: 설명' 형식으로 답하도록 유도하여, 그 답변을
        '📥 AI 답변 설명 일괄 붙여넣기'로 그대로 일괄 입력할 수 있게 한다.
        """
        cols = [str(c) for c in fd.df.columns]
        csv_line = ", ".join(cols)
        prompt = (
            "다음은 한 업무용 데이터 테이블의 컬럼 목록입니다.\n"
            "각 컬럼이 어떤 데이터를 담는지, 나중에 컬럼명을 보지 않고도\n"
            "내용을 알아볼 수 있도록 한국어로 명확하게 설명해 주세요.\n"
            "\n"
            "[작성 지침]\n"
            "- 형식은 \"컬럼명: 설명\" 이며, 각 컬럼 항목 사이는 반드시 백슬래시(\\) 로 구분하세요.\n"
            "  예) 사번: 직원 고유 식별 번호 \\ 이름: 직원 성명 \\ 부서: 소속 부서명\n"
            "- 줄바꿈을 쓰지 말고, 모든 항목을 백슬래시(\\)로 이어 한 줄(또는 한 단락)로 출력하세요.\n"
            "- 설명은 한 문장(대략 20~45자)으로, 무엇을 담는 값인지 구체적으로 쓰세요.\n"
            "- 단위·형식이 유추되면 함께 적으세요 (예: 금액(원), 날짜(YYYY-MM-DD), 비율(%)).\n"
            "- 번호 매기기·머리말·표·부가 설명 없이 \"컬럼명: 설명\" 항목만 출력하세요.\n"
            "- 컬럼 순서는 아래 목록 순서를 그대로 유지하세요.\n"
            "\n"
            f"[컬럼 목록 — 총 {len(cols)}개]\n"
            f"{csv_line}"
        )
        return prompt, csv_line, cols

    def _copy_columns_csv(self, fd):
        """이 파일의 컬럼 목록 + AI 질문 프롬프트를 한 번에 클립보드에 복사.

        사용자가 Claude/ChatGPT 채팅창에 그대로 붙여넣으면 되도록,
        '컬럼명: 설명' 형식 답변을 요청하는 프롬프트까지 포함한다.
        AI 답변은 다시 복사해 '📥 AI 답변 설명 일괄 붙여넣기'로 입력한다.
        """
        if fd.df is None:
            return
        prompt, csv_line, cols = self._build_ai_prompt(fd)
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(prompt)
            self.root.update_idletasks()
        except tk.TclError:
            messagebox.showerror("복사 실패", "클립보드에 접근할 수 없습니다.",
                                 parent=self.root)
            return
        self.status_lbl.config(
            text=f"{fd.file_name}: 컬럼 {len(cols)}개 + 질문 프롬프트 복사 완료 (클립보드)")
        # parent 지정 — 팝업이 메인 창 뒤로 숨어 '멈춘 것처럼' 보이는 현상 방지
        messagebox.showinfo("AI 질문 프롬프트 복사 완료",
            f"컬럼 {len(cols)}개와 질문 프롬프트를 함께 복사했습니다.\n"
            f"AI 채팅창(Claude/ChatGPT 등)에 그대로 붙여넣으세요.\n"
            f"AI가 '컬럼명: 설명' 형식으로 답하면, 그 답변을 복사해\n"
            f"'📥 AI 답변 설명 일괄 붙여넣기'로 입력하면 됩니다.\n\n"
            f"--- 미리보기 ---\n{prompt[:260]} ...",
            parent=self.root)

    def _parse_description_block(self, fd, text):
        """붙여넣은 텍스트 블록을 {원본컬럼: 설명} 매핑으로 파싱.

        두 가지 형식을 자동 감지한다:
          (A) '컬럼명: 설명' / '컬럼명, 설명' / '컬럼명<TAB>설명' → 이름 기반 매핑
          (B) 한 줄에 설명 1개 → 컬럼 순서대로 위치 기반 매핑
        선행 번호/불릿('1.', '2)', '-', '*', '•')은 제거한다.

        Returns:
            (mapping, mode, leftover)
              mode: 'name' / 'position' / 'empty'
              leftover: 매칭되지 않은(이름) 또는 컬럼 수 초과(위치) 줄 리스트
        """
        real_cols = list(fd.df.columns)
        norm_to_col = {_normalize_col_name(str(c)): c for c in real_cols}

        # 구분자 정규화: 줄바꿈 없이 백슬래시(\)로 이어 붙인 답변도 지원.
        # 백슬래시를 줄바꿈으로 치환해 동일한 라인 파싱 경로를 태운다.
        normalized = text.replace('\r', '').replace('\\', '\n')
        raw_lines = [ln.strip() for ln in normalized.split('\n')]
        lines = [ln for ln in raw_lines if ln]
        if not lines:
            return {}, 'empty', []

        def _strip_bullet(s):
            return re.sub(r'^\s*(?:[-*•]|\(?\d+[.)])\s*', '', s).strip()

        # (A) 이름 기반 매핑 시도
        name_mapping = {}
        name_hits = 0
        leftover = []
        for ln in lines:
            body = _strip_bullet(ln)
            parts = re.split(r'\s*[:\t,]\s*', body, maxsplit=1)
            if len(parts) == 2:
                left, right = parts[0].strip(), parts[1].strip()
                key = _normalize_col_name(left)
                if key in norm_to_col and right:
                    name_mapping[norm_to_col[key]] = right
                    name_hits += 1
                    continue
            leftover.append(ln)

        # 컬럼의 과반 이상이 이름으로 매칭되면 이름 기반 채택
        if name_hits >= max(1, len(real_cols) // 2):
            return name_mapping, 'name', leftover

        # (B) 위치 기반 매핑 — 줄 순서대로 컬럼에 대응
        stripped = [_strip_bullet(ln) for ln in lines]
        pos_mapping = {}
        for i, desc in enumerate(stripped):
            if i >= len(real_cols):
                break
            if desc:
                pos_mapping[real_cols[i]] = desc
        extra = stripped[len(real_cols):] if len(stripped) > len(real_cols) else []
        return pos_mapping, 'position', extra

    def _apply_descriptions(self, fd, mapping):
        """{원본컬럼: 설명} 매핑을 fd.col_descriptions와 화면 Entry에 반영.

        Returns: 실제로 적용된 컬럼 수
        """
        applied = 0
        for col, desc in mapping.items():
            if col not in fd.df.columns:
                continue
            fd.col_descriptions[col] = desc
            entry = fd.desc_entries.get(col)
            if entry is not None:
                try:
                    entry._is_placeholder = False
                    entry.delete(0, tk.END)
                    entry.insert(0, desc)
                    entry.configure(foreground="black")
                except tk.TclError:
                    pass
            applied += 1
        return applied

    def _paste_descriptions_dialog(self, fd):
        """AI 채팅 답변(컬럼 설명)을 한 번에 붙여넣어 설명란을 일괄 채우는 대화상자."""
        if fd.df is None or fd.analysis_text is None:
            messagebox.showwarning("안내",
                "먼저 이 파일을 분석한 뒤 사용하세요.", parent=self.root)
            return
        cols = list(fd.df.columns)

        dlg = tk.Toplevel(self.root)
        dlg.title("AI 답변 설명 일괄 붙여넣기")
        dlg.transient(self.root)

        # ── 메인 창 중앙에 강제 배치 + 화면 안으로 클램프 ──
        # (모달 창이 메인 창 뒤/화면 밖에 떠서 앱 전체가 '멈춘 것처럼' 잠기는 현상 방지)
        dw, dh = 660, 520
        self.root.update_idletasks()
        try:
            rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
            dx = max(0, min(rx + (rw - dw) // 2, sw - dw))
            dy = max(0, min(ry + (rh - dh) // 2, sh - dh - 40))
        except tk.TclError:
            dx, dy = 100, 80
        dlg.geometry(f"{dw}x{dh}+{dx}+{dy}")

        # 맨 앞으로 끌어올리기 (잠깐 topmost 후 해제)
        dlg.lift()
        dlg.attributes('-topmost', True)
        dlg.after(300, lambda: dlg.attributes('-topmost', False))
        dlg.bind('<Escape>', lambda _e: dlg.destroy())

        # grab은 창이 실제로 보인 뒤에 — 비가시 상태 grab으로 인한 입력 잠김 방지
        def _safe_grab():
            try:
                dlg.grab_set()
                dlg.focus_force()
            except tk.TclError:
                pass
        dlg.after(100, _safe_grab)

        info = (
            f"이 파일에는 컬럼이 {len(cols)}개 있습니다.\n"
            "AI 채팅창(Claude/ChatGPT 등)의 답변을 복사해 아래에 붙여넣고 [미리보기]를 누르세요.\n"
            "지원 형식:\n"
            "  • 컬럼명: 설명   ('컬럼명, 설명' / 탭 구분도 가능) — 이름으로 매칭\n"
            "  • 항목 구분: 줄바꿈 또는 백슬래시(\\) 모두 인식 (줄바꿈 못 하는 AI 대응)\n"
            "  • 한 줄에 설명 1개 — 컬럼 순서대로 매칭 (줄 수 = 컬럼 수 권장)\n"
            "  • 선행 번호(1. / 2) / - )는 자동 제거됩니다."
        )
        ttk.Label(dlg, text=info, justify='left',
                  font=("맑은 고딕", 9)).pack(fill=tk.X, padx=12, pady=(10, 6))

        txt = scrolledtext.ScrolledText(dlg, height=12, font=("맑은 고딕", 10),
                                        wrap=tk.WORD)
        txt.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 6))
        txt.focus_set()

        preview_lbl = ttk.Label(dlg, text="", justify='left',
                                font=("맑은 고딕", 9, "bold"),
                                foreground="#2980b9")
        preview_lbl.pack(fill=tk.X, padx=12, pady=(0, 4))

        state = {'mapping': {}, 'mode': None}

        def _do_preview():
            raw = txt.get("1.0", tk.END)
            mapping, mode, leftover = self._parse_description_block(fd, raw)
            state['mapping'] = mapping
            state['mode'] = mode
            mode_label = {'name': '이름 기반', 'position': '순서 기반',
                          'empty': '입력 없음'}.get(mode, mode)
            msg = (f"매칭 방식: {mode_label}  |  "
                   f"채워질 컬럼: {len(mapping)} / {len(cols)}")
            if leftover:
                msg += f"  |  매칭 안 된 줄: {len(leftover)}개"
            preview_lbl.config(text=msg)

        def _do_apply():
            if not state['mapping']:
                _do_preview()
            mapping = state['mapping']
            if not mapping:
                messagebox.showwarning("적용 불가",
                    "매칭된 설명이 없습니다. 입력 내용을 확인해 주세요.",
                    parent=dlg)
                return
            applied = self._apply_descriptions(fd, mapping)
            dlg.destroy()
            self.status_lbl.config(
                text=f"{fd.file_name}: 설명 {applied}개 일괄 입력 완료")
            messagebox.showinfo("적용 완료",
                f"{applied}개 컬럼 설명을 입력했습니다.", parent=self.root)

        btn_bar = ttk.Frame(dlg)
        btn_bar.pack(fill=tk.X, padx=12, pady=(0, 12))
        ttk.Button(btn_bar, text="🔍 미리보기",
                   command=_do_preview).pack(side=tk.LEFT)
        ttk.Button(btn_bar, text="✅ 적용",
                   command=_do_apply).pack(side=tk.RIGHT)
        ttk.Button(btn_bar, text="취소",
                   command=dlg.destroy).pack(side=tk.RIGHT, padx=(0, 6))

    def _apply_column_rename(self, fd):
        """컬럼명 변경 Entry 값을 fd.rename_values에 저장 + 중복 검사.

        실제 컬럼명 적용은 _confirm_current_file에서 일괄 처리되므로
        여기서는 임시 저장만 함.
        """
        self._save_current_edits(fd)
        new_names = [fd.rename_entries[c].get().strip()
                      or fd.column_mask_map.get(c, c)
                      for c in fd.rename_entries]
        if len(set(new_names)) != len(new_names):
            messagebox.showerror("오류", "컬럼명이 중복됩니다. 다시 확인해 주세요.")
            return
        messagebox.showinfo("적용 완료",
            "컬럼명이 임시 저장되었습니다.\n"
            "'✅ 이 파일 변환 계획 확정' 버튼을 눌러야 실제 적용됩니다.")

    def _confirm_current_file(self, fd):
        """이 파일의 변환 계획을 확정 — 컬럼명/가짜값 중복 검사 후 fd.confirmed=True.

        모든 파일이 확정되면 실행 단계로 진입.
        """
        self._save_current_edits(fd)

        # 컬럼명 중복 검사 (이 파일 내)
        new_names = []
        for orig in fd.df.columns:
            v = fd.rename_values.get(orig, '').strip()
            new_names.append(v if v else str(orig))
        if len(set(new_names)) != len(new_names):
            messagebox.showerror("오류",
                "이 파일 내에서 컬럼명이 중복됩니다. 다시 확인해 주세요.")
            return

        # 가짜값 중복 검사
        dup_issues = []
        for col, mp in fd.mapping_values.items():
            seen = {}
            for o, f in mp.items():
                if f in seen:
                    dup_issues.append(
                        f"  ⚠ '{col}': 가짜값 '{f}'이(가) "
                        f"'{seen[f]}'와 '{o}'에 중복됨")
                else:
                    seen[f] = o
        if dup_issues:
            msg = ("가짜값이 중복된 항목이 있습니다.\n"
                   "복원 시 구분이 불가능합니다.\n\n"
                   + "\n".join(dup_issues)
                   + "\n\n그래도 진행하시겠습니까?")
            if not messagebox.askyesno("매핑 중복 경고", msg, icon='warning'):
                return

        fd.confirmed = True
        fd.confirm_btn.config(state='disabled', bg="#95a5a6",
                                text="  ✅  확정 완료  ")
        fd.confirm_status_lbl.config(
            text="✅ 이 파일이 확정되었습니다. 다른 파일 탭으로 이동하세요.",
            fg="#27ae60")
        self._update_file_tab_label(fd)
        self._refresh_file_listbox()

        if all(f.confirmed for f in self.files):
            self._set_step(4)
        self._update_run_state()

    # ──────────────────────────────────────────────────────────
    # 관계 확정 / 수정 / 공유매핑 적용
    # ──────────────────────────────────────────────────────────

    def _confirm_relationships(self, auto_skip=False):
        """관계 체크박스 확정 → 공유 매핑 생성 → 컬럼편집 탭 활성화.

        Args:
            auto_skip: True면 관계 없음 케이스에서 자동 호출된 것이므로
                      재확정 확인 다이얼로그를 건너뜀.
        """
        if not self.files:
            messagebox.showwarning("경고", "먼저 파일을 추가하고 분석하세요.")
            return

        # 관계 확정 시 모든 파일의 기존 확정 상태를 초기화 (공유 매핑이 바뀌므로)
        any_previously_confirmed = any(f.confirmed for f in self.files)
        if any_previously_confirmed and not auto_skip:
            if not messagebox.askyesno("재확정 확인",
                    "관계를 재확정하면 이미 확정된 파일들의 확정 상태가 초기화되고\n"
                    "관계 그룹 컬럼의 가짜값이 새로 계산됩니다.\n\n계속하시겠습니까?"):
                return
            for fd in self.files:
                fd.confirmed = False
        elif any_previously_confirmed:
            for fd in self.files:
                fd.confirmed = False

        # 공유 매핑 적용
        self._apply_shared_mappings()

        # 관계 체크박스 잠금
        for cb in self.rel_check_widgets.values():
            try:
                cb.config(state='disabled')
            except tk.TclError:
                pass

        self.relationships_confirmed = True
        self.confirm_rel_btn.config(state='disabled', bg="#95a5a6",
                                      text="  ✅  관계 확정 완료  ")
        self.unlock_rel_btn.config(state='normal')

        n_active = sum(1 for r in self.relationships if r.get('enabled'))
        groups = build_relationship_groups(self.relationships)
        self.rel_status_lbl.config(
            text=f"✅ 관계 {n_active}개 확정 → 그룹 {len(groups)}개 공유 매핑 적용됨",
            fg="#27ae60")

        # 컬럼 편집 sub-tab을 새로 빌드
        self._clear_col_inner_tabs()
        for i, fd in enumerate(self.files):
            if fd.analyzed:
                self._build_file_tab(fd, i)

        # 컬럼 편집 탭 활성화 + 자동 전환
        self.notebook.tab(self.tab_cols_idx, state='normal')
        self.notebook.select(self.tab_cols_idx)
        self.cols_top_lbl.config(
            text=f"▶ 각 파일 탭에서 컬럼명/가짜값을 확인하고 '✅ 이 파일 확정' 버튼을 누르세요. "
                 f"🔗 관계 그룹 컬럼은 공유 매핑이 자동 적용되어 잠금 상태입니다.")

        self._refresh_file_listbox()
        self._set_step(3)
        self._update_run_state()

    def _apply_shared_mappings(self):
        """관계 그룹에 속한 컬럼들에 동일한 가짜값 매핑을 적용 + fd.linked_cols 갱신."""
        # 우선 모든 파일의 linked_cols 초기화
        for fd in self.files:
            fd.linked_cols = {}

        groups = build_relationship_groups(self.relationships)
        for gi, group in enumerate(groups):
            # 그룹 전체 unique 값 union
            all_vals = set()
            is_pii_in_group = False
            is_id_in_group = False
            looks_person = False
            looks_company = False
            for (fi, col) in group:
                if fi >= len(self.files):
                    continue
                fd = self.files[fi]
                if fd.df is None or col not in fd.df.columns:
                    continue
                vals = fd.df[col].dropna().astype(str).unique()
                all_vals.update(vals)
                if col in fd.pii_auto_detected:
                    is_pii_in_group = True
                if col in fd.id_suspected:
                    is_id_in_group = True
                series = fd.df[col].dropna().astype(str)
                if len(series) > 0:
                    if series.str.match(r'^[가-힣]{2,4}$').mean() > 0.7:
                        looks_person = True
                    if str(col).lower() in ['client', '고객', '발주', '선주',
                                              'company', '업체', '회사', '회사명']:
                        looks_company = True

            sorted_vals = sorted(all_vals)
            n = len(sorted_vals)
            if n == 0:
                continue
            seed = abs(hash(tuple(sorted(
                (int(fi), str(c)) for fi, c in group)))) % 999999

            if is_id_in_group:
                orig_prefix = _extract_id_prefix(sorted_vals)
                fake_prefix = _pick_different_prefix(orig_prefix)
                pool = [fake_prefix + str(i + 1).zfill(4) for i in range(n)]
            elif looks_person or (is_pii_in_group and not looks_company):
                pool = generate_fake_persons(n, seed=seed)
            elif looks_company:
                pool = generate_fake_companies(n, seed=seed)
            else:
                pool = [f"CAT_REL{gi+1}_{i+1:04d}" for i in range(n)]

            shared_map = {v: pool[i] for i, v in enumerate(sorted_vals)}

            # 각 멤버 컬럼의 mapping_values를 공유 매핑으로 갱신 + linked_cols 등록
            for (fi, col) in group:
                if fi >= len(self.files):
                    continue
                fd = self.files[fi]
                if fd.df is None or col not in fd.df.columns:
                    continue
                col_vals = fd.df[col].dropna().astype(str).unique()
                # 이 컬럼의 원본 값에 대응하는 부분만 추출
                fd.mapping_values[col] = {
                    v: shared_map[v] for v in col_vals if v in shared_map}
                fd.linked_cols[col] = gi

    def _unlock_relationships(self):
        """관계 잠금 해제 — 체크박스 다시 편집 가능, 확정 상태 리셋."""
        any_confirmed = any(f.confirmed for f in self.files)
        if any_confirmed:
            if not messagebox.askyesno("관계 수정 확인",
                    "관계를 수정하면 이미 확정된 파일의 확정 상태가 초기화됩니다.\n"
                    "(관계 그룹이 바뀌면 공유 매핑도 다시 계산되어야 함)\n\n"
                    "계속하시겠습니까?"):
                return

        self.relationships_confirmed = False
        for cb in self.rel_check_widgets.values():
            try:
                cb.config(state='normal')
            except tk.TclError:
                pass
        self.confirm_rel_btn.config(state='normal', bg="#27ae60",
                                      text="  ✅  관계 확정 → 컬럼 편집으로 이동  ")
        self.unlock_rel_btn.config(state='disabled')
        self.rel_status_lbl.config(text="관계 수정 가능", fg="#e67e22")

        # 컬럼 편집 탭 비활성화 + sub-tabs 제거
        self._clear_col_inner_tabs()
        self.notebook.tab(self.tab_cols_idx, state='disabled')
        for fd in self.files:
            fd.confirmed = False
            fd.linked_cols = {}
        self._refresh_file_listbox()
        self.cols_top_lbl.config(
            text="◀ 관계 설정 탭에서 다시 '관계 확정'을 눌러주세요.")
        self._set_step(2)
        self._update_run_state()

    # ──────────────────────────────────────────────────────────
    # 관계 감지/설정
    # ──────────────────────────────────────────────────────────

    def _detect_relationships(self):
        """파일 간 컬럼 관계를 자동 재감지 (이미 확정된 경우 잠금 해제 후).

        detect_relationships()를 호출해 self.relationships 갱신 + 관계 탭에 표시.
        """
        analyzed_files = [f for f in self.files if f.analyzed]
        if len(analyzed_files) < 2:
            messagebox.showinfo("알림",
                "관계를 감지하려면 2개 이상의 분석된 파일이 필요합니다.")
            return
        if self.relationships_confirmed:
            if not messagebox.askyesno("관계 재감지",
                    "관계가 이미 확정되어 있습니다. 재감지하면 잠금이 해제되고\n"
                    "확정된 파일들의 상태도 초기화됩니다. 진행하시겠습니까?"):
                return
            self._unlock_relationships()

        self.status_lbl.config(text="컬럼 관계 자동 감지 중...")
        self.root.update_idletasks()

        rels = detect_relationships(self.files)
        self.relationships = rels
        self._refresh_relationship_widgets()
        n_suggested = sum(1 for r in rels if r.get('enabled'))
        self.rel_status_lbl.config(
            text=f"{len(rels)}개 발견 (추천 {n_suggested}개)", fg="#666")
        self.status_lbl.config(
            text=f"관계 감지 완료 — 총 {len(rels)}개, 추천 {n_suggested}개")
        self.notebook.select(self.tab_rel_idx)
        self.confirm_rel_btn.config(state='normal')
        self._update_run_state()

    def _refresh_relationship_widgets(self):
        """관계 설정 탭의 체크박스 행을 self.relationships 기준으로 다시 그림.

        관계 확정 상태이면 체크박스는 disabled로 표시.
        """
        for w in self.rel_inner.winfo_children():
            w.destroy()
        self.rel_check_vars = {}
        self.rel_check_widgets = {}

        if not self.relationships:
            self.rel_empty_lbl = ttk.Label(self.rel_inner,
                text="\n\n   ◀ '📊 전체 파일 분석' 후 자동으로 관계가 여기 표시됩니다.\n",
                style="Sub.TLabel")
            self.rel_empty_lbl.pack(anchor='w')
            return

        # 헤더
        header = ttk.Frame(self.rel_inner)
        header.pack(fill=tk.X, padx=5, pady=(5, 4))
        ttk.Label(header, text="✓", width=3,
                  font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="참조 파일.컬럼", width=42, anchor='w',
                  font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(3, 0))
        ttk.Label(header, text="↔", width=2,
                  font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="대상 파일.컬럼", width=42, anchor='w',
                  font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(3, 0))
        ttk.Label(header, text="신뢰도", width=8,
                  font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text="근거", anchor='w',
                  font=("맑은 고딕", 9, "bold")).pack(side=tk.LEFT, padx=(3, 0))
        ttk.Separator(self.rel_inner, orient=tk.HORIZONTAL).pack(
            fill=tk.X, padx=5)

        cb_state = 'disabled' if self.relationships_confirmed else 'normal'
        for idx, r in enumerate(self.relationships):
            row = ttk.Frame(self.rel_inner)
            row.pack(fill=tk.X, padx=5, pady=1)

            cv = tk.BooleanVar(value=r.get('enabled', False))
            self.rel_check_vars[idx] = cv

            def _on_check(i=idx, var=cv):
                """체크박스 토글 → self.relationships[i]['enabled'] 동기화."""
                self.relationships[i]['enabled'] = var.get()

            cb = ttk.Checkbutton(row, variable=cv, command=_on_check,
                                  state=cb_state)
            cb.pack(side=tk.LEFT, padx=(0, 4))
            self.rel_check_widgets[idx] = cb

            from_file = self.files[r['from_file_idx']].file_name
            to_file = self.files[r['to_file_idx']].file_name
            ttk.Label(row,
                      text=f"{from_file[:18]}.{r['from_col']}",
                      width=42, anchor='w',
                      font=("Consolas", 9)).pack(side=tk.LEFT, padx=(3, 0))
            ttk.Label(row, text="↔", width=2).pack(side=tk.LEFT)
            ttk.Label(row,
                      text=f"{to_file[:18]}.{r['to_col']}",
                      width=42, anchor='w',
                      font=("Consolas", 9)).pack(side=tk.LEFT, padx=(3, 0))

            conf = r['confidence']
            conf_color = ("#27ae60" if conf >= 0.7
                          else ("#e67e22" if conf >= 0.5 else "#95a5a6"))
            ttk.Label(row, text=f"{conf:.2f}", width=8,
                      foreground=conf_color,
                      font=("Consolas", 9, "bold")).pack(side=tk.LEFT)
            ttk.Label(row, text=r.get('reason', ''),
                      foreground="#666",
                      font=("맑은 고딕", 9)).pack(side=tk.LEFT, padx=(3, 0))

    def _toggle_relationships(self, only_suggested=False, check_all=None):
        """관계 체크박스 일괄 토글 — 추천만/전체선택/전체해제 중 하나.

        Args:
            only_suggested: True면 confidence >= 0.45만 ON.
            check_all: True면 모두 ON, False면 모두 OFF (only_suggested보다 우선순위 낮음).
        """
        if not self.relationships:
            return
        for idx, r in enumerate(self.relationships):
            if only_suggested:
                # 추천(원래 enabled가 True였던 것들)만 ON
                new_val = (r.get('confidence', 0) >= 0.45)
            elif check_all is True:
                new_val = True
            elif check_all is False:
                new_val = False
            else:
                continue
            r['enabled'] = new_val
            if idx in self.rel_check_vars:
                self.rel_check_vars[idx].set(new_val)
        self._update_run_state()

    def _add_manual_relationship(self):
        """수동으로 관계를 추가하는 다이얼로그."""
        analyzed = [(i, f) for i, f in enumerate(self.files) if f.analyzed]
        if len(analyzed) < 2:
            messagebox.showinfo("알림", "분석된 파일이 2개 이상 필요합니다.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("수동 관계 추가")
        dlg.geometry("600x300")
        dlg.transient(self.root)
        dlg.grab_set()

        ttk.Label(dlg, text="관계로 묶을 두 파일의 컬럼을 선택하세요.",
                  style="Section.TLabel").pack(pady=(10, 5))

        # 파일1
        row1 = ttk.Frame(dlg); row1.pack(fill=tk.X, padx=20, pady=4)
        ttk.Label(row1, text="파일 A:", width=10).pack(side=tk.LEFT)
        file1_var = tk.StringVar()
        file1_cb = ttk.Combobox(row1, textvariable=file1_var, state='readonly',
                                  width=35,
                                  values=[f"{i+1}. {f.file_name}"
                                           for i, f in analyzed])
        file1_cb.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(row1, text="컬럼:").pack(side=tk.LEFT)
        col1_var = tk.StringVar()
        col1_cb = ttk.Combobox(row1, textvariable=col1_var, state='readonly',
                                 width=25)
        col1_cb.pack(side=tk.LEFT)

        # 파일2
        row2 = ttk.Frame(dlg); row2.pack(fill=tk.X, padx=20, pady=4)
        ttk.Label(row2, text="파일 B:", width=10).pack(side=tk.LEFT)
        file2_var = tk.StringVar()
        file2_cb = ttk.Combobox(row2, textvariable=file2_var, state='readonly',
                                  width=35,
                                  values=[f"{i+1}. {f.file_name}"
                                           for i, f in analyzed])
        file2_cb.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(row2, text="컬럼:").pack(side=tk.LEFT)
        col2_var = tk.StringVar()
        col2_cb = ttk.Combobox(row2, textvariable=col2_var, state='readonly',
                                 width=25)
        col2_cb.pack(side=tk.LEFT)

        def _update_cols(file_var, col_cb):
            """파일 콤보박스 선택 시 해당 파일의 컬럼 목록을 col_cb에 채움."""
            v = file_var.get()
            if not v:
                return
            try:
                idx = int(v.split('.')[0]) - 1
            except ValueError:
                return
            fd = self.files[idx]
            cols = [str(c) for c in fd.df.columns]
            col_cb.config(values=cols)
            if cols:
                col_cb.current(0)

        file1_cb.bind('<<ComboboxSelected>>',
                       lambda e: _update_cols(file1_var, col1_cb))
        file2_cb.bind('<<ComboboxSelected>>',
                       lambda e: _update_cols(file2_var, col2_cb))

        status_lbl = ttk.Label(dlg, text="", style="Sub.TLabel")
        status_lbl.pack(pady=4)

        def _on_add():
            """수동 관계 다이얼로그의 '추가' 클릭 핸들러 — 검증 후 self.relationships에 append."""
            v1 = file1_var.get(); v2 = file2_var.get()
            c1 = col1_var.get(); c2 = col2_var.get()
            if not all([v1, v2, c1, c2]):
                status_lbl.config(text="모든 항목을 선택하세요.", foreground="#c0392b")
                return
            i1 = int(v1.split('.')[0]) - 1
            i2 = int(v2.split('.')[0]) - 1
            if i1 == i2:
                status_lbl.config(text="다른 파일을 선택하세요.", foreground="#c0392b")
                return
            from_idx, to_idx = sorted([i1, i2])
            if from_idx == i2:
                c1, c2 = c2, c1
            # 중복 검사
            for r in self.relationships:
                if (r['from_file_idx'] == from_idx and r['to_file_idx'] == to_idx
                        and r['from_col'] == c1 and r['to_col'] == c2):
                    status_lbl.config(text="이미 등록된 관계입니다.",
                                       foreground="#e67e22")
                    return
            # 값 점수 계산
            f1 = self.files[from_idx]; f2 = self.files[to_idx]
            vals1 = f1.df[c1].dropna().unique()
            vals2 = f2.df[c2].dropna().unique()
            name_score = compute_name_similarity(c1, c2)
            value_score = compute_value_overlap(vals1, vals2)
            confidence = min(1.0, 0.5 * name_score + 0.5 * value_score + 0.1)
            self.relationships.append({
                'from_file_idx': from_idx, 'from_col': c1,
                'to_file_idx': to_idx, 'to_col': c2,
                'name_score': round(name_score, 3),
                'value_score': round(value_score, 3),
                'confidence': round(confidence, 3),
                'reason': '수동 추가',
                'enabled': True,
            })
            self.relationships.sort(key=lambda r: -r['confidence'])
            self._refresh_relationship_widgets()
            self._update_run_state()
            dlg.destroy()

        btn_row = ttk.Frame(dlg); btn_row.pack(pady=10)
        ttk.Button(btn_row, text="추가", command=_on_add).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="취소", command=dlg.destroy).pack(side=tk.LEFT, padx=4)

    def _update_run_state(self):
        """실행 버튼의 활성화 상태 갱신.

        모든 게이트(파일/관계확정/모든파일확정/프로젝트명/파일별테이블명/저장경로)가
        통과되어야 활성.
        """
        # 관계 확정 게이트: 파일이 1개면 파일 간 관계 자체가 없으므로
        # 관계 확정을 요구하지 않는다 (단일 파일도 실행 가능하도록).
        rel_gate = self.relationships_confirmed or len(self.files) <= 1
        ready = (
            len(self.files) > 0
            and rel_gate
            and all(f.confirmed for f in self.files)
            and bool(self.employee_id_var.get().strip())
            and all((f.table_name or '').strip() for f in self.files)
            and bool(self.save_dir.get().strip())
        )
        self.run_btn.config(state='normal' if ready else 'disabled')

    def _check_cross_file_column_duplicates(self):
        """파일 간 마스킹된 컬럼명 중복 검사.

        원복 프로그램에서 어느 파일의 컬럼인지 구분할 수 없는 충돌을 방지.
        Returns: {중복명: [(file_idx, orig_col), ...]} (중복만 포함)
        """
        name_to_origins = defaultdict(list)
        for i, fd in enumerate(self.files):
            if fd.df is None:
                continue
            for orig in fd.df.columns:
                new_name = fd.rename_values.get(orig, '').strip()
                if not new_name:
                    new_name = str(orig)
                name_to_origins[new_name].append((i, str(orig)))
        return {name: origins
                for name, origins in name_to_origins.items()
                if len(origins) > 1}

    # ──────────────────────────────────────────────────────────
    # 저장/유틸
    # ──────────────────────────────────────────────────────────

    def _apply_db_check_style(self, fd):
        """fd._db_check_status 값에 따라 테이블명 Entry 배경색을 적용 (tk.Entry).

        상태값:
            'ok'        → 녹색 (DB에 없음, 사용 가능)
            'collision' → 빨강 (DB에 이미 존재)
            'no_conn'   → 노랑 (DB 연결 실패)
            None        → 흰색 (검사 전/입력 중)
        """
        entry = getattr(fd, '_db_check_entry', None)
        if entry is None:
            return
        status = getattr(fd, '_db_check_status', None)
        bg = {
            'ok':        '#a8e6a3',   # 녹색
            'collision': '#f5a8a8',   # 빨강
            'no_conn':   '#f5e9a8',   # 노랑
        }.get(status, 'white')
        try:
            entry.configure(bg=bg)
        except tk.TclError:
            pass

    def _check_table_in_db_async(self, fd):
        """포커스 아웃 시 DB에 동일 테이블이 존재하는지 백그라운드 검사.

        결과는 테이블명 Entry 배경색으로 표시 (전체 분석 등 위젯 재구성 후에도
        fd._db_check_status에 저장되어 _apply_db_check_style로 복원됨):
            녹색  사용 가능 (DB에 없음)
            빨강  중복 (DB에 이미 존재)
            노랑  DB 연결 실패 (검사 생략)
        """
        user_table = (fd.table_name or '').strip()
        project = self.employee_id_var.get().strip()
        if not user_table or not project:
            fd._db_check_status = None
            self._apply_db_check_style(fd)
            return

        full_name = f"{project}_{user_table}"

        def _work():
            exists = _db_table_exists(full_name)
            if exists is True:
                fd._db_check_status = 'collision'
            elif exists is False:
                fd._db_check_status = 'ok'
            else:
                fd._db_check_status = 'no_conn'
            self.root.after(0, lambda: self._apply_db_check_style(fd))
        threading.Thread(target=_work, daemon=True).start()

    def _check_all_tables_in_db(self):
        """실행 직전 모든 파일의 테이블명을 DB와 비교 — 중복 목록 반환."""
        project = self.employee_id_var.get().strip()
        conn = _db_connect()
        if conn is None:
            return []  # DB 연결 불가 시 검사 생략
        try:
            collisions = []
            for fd in self.files:
                t = (fd.table_name or '').strip()
                if not t:
                    continue
                full = f"{project}_{t}" if project else t
                if _db_table_exists(full, conn=conn) is True:
                    collisions.append((fd.file_name, full))
            return collisions
        finally:
            try:
                conn.close()
            except Exception:
                pass

    def _validate_project_name(self, proposed):
        """프로젝트명 입력 검증 — 영문 소문자만 허용. 위반 시 경고 팝업."""
        if proposed == '' or re.match(r'^[a-z]+$', proposed):
            return True
        # 중복 팝업 방지를 위해 after_idle로 한 번만 표시
        if not getattr(self, '_proj_warn_pending', False):
            self._proj_warn_pending = True
            def _show():
                try:
                    messagebox.showwarning("입력 제한",
                        "영문 소문자만 가능합니다.")
                finally:
                    self._proj_warn_pending = False
            self.root.after(1, _show)
        return False

    def _validate_table_name(self, proposed):
        """파일별 테이블명 입력 검증 — 영문 소문자만 허용. 위반 시 경고 팝업."""
        if proposed == '' or re.match(r'^[a-z]+$', proposed):
            return True
        if not getattr(self, '_table_warn_pending', False):
            self._table_warn_pending = True
            def _show():
                try:
                    messagebox.showwarning("입력 제한",
                        "영문 소문자만 가능합니다.")
                finally:
                    self._table_warn_pending = False
            self.root.after(1, _show)
        return False

    def _browse_dir(self):
        """저장 폴더 선택 다이얼로그 → self.save_dir에 반영."""
        d = filedialog.askdirectory(title="저장 폴더 선택")
        if d:
            self.save_dir.set(d)

    def _check_other_and_prompt(self, files_with_other):
        """OTHER 통합 발생 시 사용자 선택 다이얼로그 — 'keep' 또는 'redo' 반환.

        Args:
            files_with_other: fd.other_buckets가 비어있지 않은 FileData 리스트.

        Returns:
            'keep'  — OTHER 그대로 유지 (변환키 JSON에 other_buckets 포함)
            'redo' — k값 낮춰 재합성 (이번 출력 모두 폐기)
        """
        msg_lines = [
            "⚠️ 다음 컬럼들에서 1:1 매핑(빈도 < k)이 많아",
            "OTHER로 통합되었습니다:",
            "",
        ]
        for fd in files_with_other:
            xlsx_name = (os.path.basename(fd.output_xlsx)
                         if fd.output_xlsx else fd.file_name)
            msg_lines.append(f"  • {xlsx_name}")
            for col, info in fd.other_buckets.items():
                msg_lines.append(
                    f"      - {col}: {info['n_bucketed_originals']}개 원본값 "
                    f"→ OTHER (k={info['k_anon']})")
        msg_lines.extend([
            "",
            "권장: 공통 설정의 🛡️ k-익명성 값을 낮추면",
            "OTHER 없이 합성 가능합니다 (예: k=2 또는 k=1).",
            "",
            "[예] OTHER 유지하고 진행 (변환키에 OTHER 정보 포함)",
            "[아니오] k값 낮추고 재합성 (이번 출력 모두 삭제)",
        ])
        full_msg = "\n".join(msg_lines)
        result = messagebox.askyesno("OTHER 통합 발생", full_msg,
                                      default='yes', icon='warning')
        return 'keep' if result else 'redo'

    def _cleanup_synth_outputs(self, save_dir, employee_id, files):
        """재합성 선택 시 이번 실행에서 생성한 출력 파일 일괄 삭제.

        파일별 (xlsx/parquet/csv/변환키.json/원본_*.parquet|csv) +
        전체 (MD 작업서/관계메타.json)를 모두 정리.
        """
        suffixes_per_file = ['.xlsx', '.parquet', '.csv', '_변환키.json']
        for fd in files:
            if not fd.output_xlsx:
                continue
            base = os.path.splitext(os.path.basename(fd.output_xlsx))[0]
            for suffix in suffixes_per_file:
                p = os.path.join(save_dir, base + suffix)
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except OSError:
                        pass
            # 원본 파일 — base는 '{prefix}_{safe_base}' 형식이라 prefix 제거
            if '_' in base:
                safe_base = base.split('_', 1)[1]
                for suffix in ['.parquet', '.csv']:
                    p = os.path.join(save_dir, f"원본_{safe_base}{suffix}")
                    if os.path.exists(p):
                        try:
                            os.remove(p)
                        except OSError:
                            pass
            # FileData 상태 리셋
            fd.output_xlsx = ''
            fd.other_buckets = {}
            fd.synth_unique_cols = set()
        # 전체 prefix 기반 파일들 (관계메타) — 프로젝트명 기반
        for suffix in ['_관계메타.json']:
            p = os.path.join(save_dir, f"{employee_id}{suffix}")
            if os.path.exists(p):
                try:
                    os.remove(p)
                except OSError:
                    pass
        # 고정 이름 파일 (테이블생성.md / tableanddata.py / data_processor작성지시.md)
        for fixed_name in ["테이블생성.md", "tableanddata.py",
                            "data_processor작성지시.md"]:
            p = os.path.join(save_dir, fixed_name)
            if os.path.exists(p):
                try:
                    os.remove(p)
                except OSError:
                    pass

    def _bind_mousewheel(self, canvas):
        """canvas에 마우스휠 스크롤 바인딩 — 마우스가 canvas 위에 있을 때만 동작.

        Enter/Leave로 글로벌 바인딩을 토글해 다른 canvas의 스크롤 이벤트와
        충돌하지 않도록 함.
        """
        def _on_wheel(event, c=canvas):
            c.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>",
            lambda e, c=canvas: c.bind_all("<MouseWheel>", _on_wheel))
        canvas.bind("<Leave>",
            lambda e, c=canvas: c.unbind_all("<MouseWheel>"))

    def _log(self, msg):
        """실행 로그 텍스트 위젯에 한 줄 추가 (스크롤 자동 이동)."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    # ──────────────────────────────────────────────────────────
    # 실행
    # ──────────────────────────────────────────────────────────

    def _run(self):
        """실행 버튼 핸들러 — 게이트(관계확정/사번/저장경로/약자) + 중복 컬럼명 검사
        후 사용자 최종 확인 다이얼로그 → 백그라운드 worker 스레드 시작.
        """
        # ── 최종 게이트 검사 ──
        # 파일이 2개 이상일 때만 관계 확정을 요구. 단일 파일은 관계가 없으므로 통과.
        if len(self.files) > 1 and not self.relationships_confirmed:
            messagebox.showwarning("경고",
                "먼저 '③ 관계 설정/확정' 탭에서 관계를 확정해 주세요.")
            return
        if not self.employee_id_var.get().strip():
            messagebox.showwarning("경고", "프로젝트명을 입력해 주세요.")
            return
        missing_tables = [f.file_name for f in self.files
                          if not (f.table_name or '').strip()]
        if missing_tables:
            messagebox.showwarning("경고",
                "다음 파일의 테이블명이 비어 있습니다 "
                "(영문 소문자):\n" + "\n".join(missing_tables))
            return
        # DB에 이미 존재하는 테이블명 검사
        collisions = self._check_all_tables_in_db()
        if collisions:
            msg = ("❌ 다음 테이블명이 이미 데이터베이스에 존재합니다.\n"
                   "다른 이름으로 수정해 주세요.\n\n")
            for fname, full in collisions:
                msg += f"  • {fname} → '{full}'\n"
            messagebox.showerror("테이블 중복 — 수정 필요", msg)
            return
        if not all(f.confirmed for f in self.files):
            unconfirmed = [f.file_name for f in self.files if not f.confirmed]
            messagebox.showwarning("경고",
                f"확정되지 않은 파일이 있습니다:\n" + "\n".join(unconfirmed))
            return
        if not self.save_dir.get().strip():
            messagebox.showwarning("경고", "저장 폴더를 선택해 주세요.")
            return

        # ── 파일 간 마스킹 컬럼명 중복 검사 (원복 충돌 방지) ──
        dups = self._check_cross_file_column_duplicates()
        if dups:
            msg = ("❌ 다른 파일과 같은 마스킹 컬럼명이 있습니다.\n"
                   "원복 프로그램에서 어느 파일의 컬럼인지 구분할 수 없습니다.\n"
                   "각 파일의 '④ 컬럼 편집' 탭에서 다르게 수정해 주세요.\n\n")
            for name, origins in list(dups.items())[:10]:
                files_str = ", ".join(
                    f"{self.files[fi].file_name}('{oc}')" for fi, oc in origins)
                msg += f"  • '{name}' → {files_str}\n"
            if len(dups) > 10:
                msg += f"  ... 외 {len(dups) - 10}건\n"
            messagebox.showerror("컬럼명 중복 — 수정 필요", msg)
            return

        # ── 최종 확인 ──
        active_rels = sum(1 for r in self.relationships if r.get('enabled'))
        confirm_msg = (
            f"전체 {len(self.files)}개 파일을 합성하고 "
            f"활성 관계 {active_rels}개를 적용합니다.\n"
            f"프로젝트명: {self.employee_id_var.get().strip()}\n\n"
            f"진행하시겠습니까?")
        if not messagebox.askyesno("최종 확인", confirm_msg):
            return

        self.run_btn.config(state='disabled')
        self.progress['value'] = 0
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        """백그라운드 스레드에서 _do_all_synth 실행 + 예외/완료 처리."""
        try:
            self._do_all_synth()
        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda m=err_msg: self._log(f"\n❌ 오류: {m}"))
            self.root.after(0, lambda m=err_msg: messagebox.showerror("오류", m))
        finally:
            self.root.after(0, lambda: self.progress.configure(value=100))

    def _do_all_synth(self):
        """전체 파일 합성 파이프라인 — 모든 파일 합성 + MD/관계메타 JSON 생성.

        실행 순서:
            1. 관계 그룹 멤버십 정리 (PK 보존 + FK 컬럼 식별용)
            2. 각 파일 _synth_one_file 호출 (rename → 매핑 → 합성 → 저장)
            3. 전체 완료 후 MD 작업서 + 관계메타 JSON 저장
        """
        L = lambda m: self.root.after(0, lambda msg=m: self._log(msg))
        S = lambda m: self.root.after(0,
            lambda msg=m: self.status_lbl.config(text=msg))
        P = lambda v: self.root.after(0,
            lambda val=v: self.progress.configure(value=val))

        employee_id = self.employee_id_var.get().strip()
        base_save_dir = self.save_dir.get().strip()
        if not employee_id:
            # 게이트에서 이미 막혔어야 하지만 안전장치
            raise ValueError("프로젝트명이 비어있습니다.")
        # 모든 산출물(xlsx/parquet/csv/변환키/MD/관계메타)을 result 하위 폴더에 모아 저장
        save_dir = os.path.join(base_save_dir, "result")
        os.makedirs(save_dir, exist_ok=True)

        try:
            k_anon_val = int(self.k_anon_var.get())
        except (tk.TclError, ValueError):
            k_anon_val = 5

        # 생성 행 수: 항상 원본과 동일하게 사용 (UI 입력 제거)
        nr_override = None

        L("=" * 60)
        L(f"  v2 다중 파일 합성 시작  |  프로젝트명: {employee_id}")
        L(f"  파일 수: {len(self.files)}  |  활성 관계: "
          f"{sum(1 for r in self.relationships if r.get('enabled'))}개")
        L("=" * 60)

        # ── 관계 그룹 멤버십 구성 ──
        # fd.mapping_values는 이미 _apply_shared_mappings로 공유 매핑이 적용된 상태.
        # 여기서는 멤버십(어느 컬럼이 어느 그룹에 속하는지)만 추적해
        # _synth_one_file에서 PII 처리(재샘플링 스킵)에 사용.
        L("\n[관계 분석] 공유 매핑 그룹 정리 (mapping_values 기준)...")
        groups = build_relationship_groups(self.relationships)
        L(f"  관계 그룹: {len(groups)}개")
        group_membership = {}  # {(file_idx, orig_col): group_idx}
        for gi, group in enumerate(groups):
            label_parts = []
            for (fi, col) in group:
                group_membership[(fi, col)] = gi
                fd = self.files[fi]
                if fd.df is not None and col in fd.df.columns:
                    label_parts.append(f"{fd.file_name}.{col}")
            n_vals = 0
            if group:
                fi0, c0 = group[0]
                fd0 = self.files[fi0]
                if c0 in fd0.mapping_values:
                    n_vals = len(fd0.mapping_values[c0])
            L(f"  🔗 그룹{gi+1}: {' ↔ '.join(label_parts)}  (공유 값 ~{n_vals}개)")

        # ── 각 파일별 합성 실행 ──
        n_files = len(self.files)
        per_file_progress = 90 / max(n_files, 1)

        for f_idx, fd in enumerate(self.files):
            base_progress = f_idx * per_file_progress
            L("\n" + "─" * 60)
            L(f"[{f_idx+1}/{n_files}] {fd.file_name} 합성 시작")
            L("─" * 60)
            S(f"[{f_idx+1}/{n_files}] {fd.file_name} 합성 중...")
            P(int(base_progress))

            self._synth_one_file(fd, f_idx, group_membership,
                                  employee_id, save_dir, employee_id,
                                  nr_override, k_anon_val, L, S, P,
                                  base_progress, per_file_progress)

        # ── Block D: OTHER 종합 점검 + 사용자 선택 다이얼로그 ──
        files_with_other = [fd for fd in self.files if fd.other_buckets]
        if files_with_other:
            L("\n" + "─" * 60)
            L(f"[OTHER 점검] {len(files_with_other)}개 파일에서 OTHER 통합 발생")
            S("OTHER 발생 — 사용자 선택 대기...")
            # UI 스레드에서 다이얼로그 표시 (worker는 백그라운드 스레드)
            decision_event = threading.Event()
            decision = [None]

            def _ask():
                """UI 스레드에서 다이얼로그 띄우고 결과를 [0]에 저장."""
                try:
                    decision[0] = self._check_other_and_prompt(files_with_other)
                finally:
                    decision_event.set()

            self.root.after(0, _ask)
            decision_event.wait()
            choice = decision[0] or 'keep'  # 다이얼로그 실패 시 안전 기본값

            if choice == 'redo':
                L("\n⚠️ 사용자가 재합성 선택 — 이번 출력 파일 모두 삭제")
                self._cleanup_synth_outputs(save_dir,
                                             employee_id, self.files)
                L("   k-익명성 값을 낮춘 뒤 ▶ 실행 버튼을 다시 클릭하세요.")
                P(100); S("재합성 대기 — k-익명성 값을 낮추세요")
                # 실행 버튼 재활성화
                self.root.after(0, lambda: self.run_btn.config(state='normal'))
                self.root.after(0, lambda: messagebox.showinfo(
                    "재합성 안내",
                    "이번 출력이 삭제되었습니다.\n\n"
                    "공통 설정의 🛡️ k-익명성 값을 더 낮게 조정한 뒤\n"
                    "▶ 실행 버튼을 다시 클릭하세요."))
                return  # MD/관계메타 생성 건너뛰고 종료

            # 'keep' 선택 — 변환키에 OTHER 포함된 상태로 정상 진행
            L("✅ 사용자가 OTHER 유지 선택 — 변환키에 OTHER 정보 포함됨")
            S("OTHER 유지 — MD 생성 계속")

        # ── MD 파일 생성 ──
        P(95); S("MD 작업서 생성..."); L("\n[MD 생성] 클로드코드용 테이블 생성 작업서...")
        md_path = os.path.join(save_dir, "테이블생성.md")
        try:
            generate_table_creation_md(employee_id, self.files,
                                        self.relationships, md_path)
            L(f"  ✅ MD 작업서: {md_path}")
        except Exception as e:
            L(f"  ❌ MD 생성 실패: {e}")

        # ── tableanddata.py 임시 실행 스크립트 생성 ──
        L("[스크립트 생성] tableanddata.py (테이블 생성 + 데이터 INSERT)...")
        td_py_path = os.path.join(save_dir, "tableanddata.py")
        try:
            generate_table_and_data_py(employee_id, self.files,
                                        self.relationships, td_py_path)
            L(f"  ✅ 실행 스크립트: {td_py_path}")
        except Exception as e:
            L(f"  ❌ tableanddata.py 생성 실패: {e}")

        # ── data_processor.py 작성 지시 MD (별도 파일) ──
        L("[MD 생성] data_processor.py 작성 지시서...")
        dp_md_path = os.path.join(save_dir, "data_processor작성지시.md")
        try:
            generate_data_processor_md(employee_id, self.files, dp_md_path,
                                        relationships=self.relationships)
            L(f"  ✅ data_processor 지시서: {dp_md_path}")
        except Exception as e:
            L(f"  ❌ data_processor 지시서 생성 실패: {e}")

        # ── 관계 메타데이터 JSON ──
        rel_json = {
            'generated_at': datetime.now().isoformat(),
            'employee_id': employee_id,
            'files': [
                {
                    'idx': i, 'file_name': fd.file_name,
                    'table_name': _table_name_for_file(employee_id, fd),
                    'sheet': fd.sheet_name, 'rows': len(fd.df),
                    'output_xlsx': os.path.basename(fd.output_xlsx),
                }
                for i, fd in enumerate(self.files)
            ],
            'relationships': self.relationships,
            'active_relationship_groups': [
                [{'file_idx': fi, 'col': c} for (fi, c) in g]
                for g in groups
            ],
        }
        rel_json_path = os.path.join(save_dir,
            f"{employee_id}_관계메타.json")
        try:
            with open(rel_json_path, 'w', encoding='utf-8') as f:
                json.dump(rel_json, f, ensure_ascii=False, indent=2,
                          default=str)
            L(f"  ✅ 관계 메타: {rel_json_path}")
        except Exception as e:
            L(f"  ❌ 관계 메타 저장 실패: {e}")

        P(100)
        L("\n" + "=" * 60)
        L(f"  🎉 전체 완료 — {n_files}개 파일 + MD 작업서 생성됨")
        L("=" * 60)
        S(f"완료 — {n_files}개 파일 합성 + MD 생성")
        self.root.after(0, lambda: self._set_step(5))
        self.root.after(0, lambda: messagebox.showinfo("완료",
            f"전체 합성 완료!\n\n"
            f"파일 수: {n_files}\n"
            f"MD 작업서: {os.path.basename(md_path)}\n"
            f"저장 위치: {save_dir}\n"
            f"(모든 산출물이 'result' 폴더 안에 생성되었습니다)"))

    def _synth_one_file(self, fd, f_idx, group_membership,
                         employee_id, save_dir, save_prefix,
                         nr_override, k_anon_val, L, S, P,
                         base_progress, per_file_progress):
        """단일 파일 합성 — v1의 _do_synth 로직을 단순화/이식.

        group_membership: {(file_idx, orig_col): group_idx} — 관계 그룹 멤버십.
        fd.mapping_values는 이미 _apply_shared_mappings로 공유 매핑이 적용된 상태.
        """
        df = fd.df
        ct = fd.col_types

        # 1) 컬럼명 변경 적용
        rename_map = {}
        for orig in df.columns:
            new_name = fd.rename_values.get(orig, '').strip()
            if new_name and new_name != str(orig):
                rename_map[orig] = new_name
        if rename_map:
            df = df.rename(columns=rename_map)
            new_ct = {}
            for c, t in ct.items():
                new_ct[rename_map.get(c, c)] = t
            ct = new_ct
            fd.pii_auto_detected = {rename_map.get(c, c)
                                      for c in fd.pii_auto_detected}
            fd.id_suspected = {rename_map.get(c, c) for c in fd.id_suspected}
            new_mapping = {rename_map.get(c, c): v
                            for c, v in fd.mapping_values.items()}
            fd.mapping_values = new_mapping
            new_modes = {rename_map.get(c, c): v
                          for c, v in fd.col_modes.items()}
            fd.col_modes = new_modes

        # 2) 매핑 수집 — fd.mapping_values를 그대로 사용 (관계 컬럼은 이미 공유 매핑)
        mapping_dict = {}
        for col, mp in fd.mapping_values.items():
            if col not in df.columns:
                continue
            mapping_dict[col] = dict(mp)

        L(f"  [1/7] 매핑 수집 완료 ({len(mapping_dict)}개 컬럼)")
        P(int(base_progress + per_file_progress * 0.1))

        # 3) 문자열 합성 — fd.mapping_values가 이미 공유 매핑이므로 그대로 적용
        df_text, desc_map = synthesize_text_columns(df, ct, mapping_dict)
        L(f"  [2/7] 문자열 합성 완료 (공유 매핑 보존)")
        P(int(base_progress + per_file_progress * 0.25))

        # 4) 카테고리 마스킹 후처리 (k-익명성)
        col_modes = {c: m for c, m in fd.col_modes.items() if c in df.columns}
        pii_cols_final = set(fd.pii_auto_detected)
        # ★ 관계 그룹 컬럼은 PII와 같은 처리 (pii_cols_final 추가)로 분류.
        # 일반 'masked' 모드는 빈도 기반 multinomial 재샘플링을 하므로 row 매핑이 깨짐
        # → 파일A의 '강호동→조우준' 매핑이 파일B에서 '강호동→임예나'로 어긋남.
        # PII로 분류하면 재샘플링을 건너뛰고 synthesize_text_columns의 row-level 매핑이 보존됨.
        rel_grouped_cols = set()
        for (gfi, gcol), _g_idx in group_membership.items():
            if gfi != f_idx:
                continue
            current_col = rename_map.get(gcol, gcol)
            if current_col in df.columns:
                pii_cols_final.add(current_col)
                rel_grouped_cols.add(current_col)
        L(f"  [관계 보존] 그룹 컬럼 {len(rel_grouped_cols)}개를 row-level 매핑 유지로 처리")

        # Block A — synthesize_categorical_masked 호출 전 가짜값 빈도 스냅샷.
        # (OTHER 통합 후에는 가짜값 → OTHER로 덮어써져 원본 추적이 불가하므로
        #  통합 전 빈도를 기록해두고, 호출 후 빈도<k_anon인 가짜값을 역추적.)
        pre_other_value_counts = {}
        for col in df_text.columns:
            if ct.get(col) != 'categorical':
                continue
            if col in pii_cols_final:
                continue
            if col_modes.get(col) == 'keep':
                continue
            pre_other_value_counts[col] = (
                df_text[col].dropna().astype(str).value_counts())

        df_text, exposed_cats, other_counts = synthesize_categorical_masked(
            df, df_text, ct, col_modes=col_modes, pii_cols=pii_cols_final,
            k_anon=k_anon_val)
        L(f"  [3/7] 카테고리 마스킹 (k={k_anon_val})")

        # Block B — OTHER로 통합된 원본값 메타데이터를 역추적해 fd.other_buckets에 저장.
        # 이후 변환키 JSON에 포함되어 복원/필터링 단계에서 사용됨.
        other_buckets = {}
        for col, n_other in other_counts.items():
            if n_other <= 0:
                continue
            freqs = pre_other_value_counts.get(col)
            if freqs is None:
                continue
            rare = freqs[freqs < int(k_anon_val)]
            if len(rare) == 0:
                continue
            col_mapping = mapping_dict.get(col, {})
            reverse = {str(fake): orig for orig, fake in col_mapping.items()}
            entries = []
            for fake_code, freq_count in rare.items():
                entries.append({
                    'value': reverse.get(str(fake_code), str(fake_code)),
                    'pre_bucket_freq': int(freq_count),
                    'fake_code': str(fake_code),
                })
            other_buckets[str(col)] = {
                'k_anon': int(k_anon_val),
                'n_bucketed_originals': len(entries),
                'total_pre_bucket_rows': int(rare.sum()),
                'original_values': entries,
            }
            # 사용자 알림 — 컬럼별 OTHER 발생 + 권장 사항
            L(f"    🛡️ {col}: {len(entries)}개 원본값이 OTHER로 통합 "
              f"(빈도<{k_anon_val}, 1:1 매핑 많음) — k값 낮추기 권장")
        fd.other_buckets = other_buckets

        P(int(base_progress + per_file_progress * 0.4))

        # 5) 상관관계/제약/수치합성
        cr = analyze_correlations(df, ct)
        cons = auto_detect_constraints(df, ct)
        nr = nr_override or len(df)
        syn_num = generate_numeric_datetime_masked(df, ct, cons, nr,
                                                     n_quantiles=64)
        L(f"  [4/7] 수치/날짜 합성 ({len(syn_num)}행) — 상관 쌍 "
          f"{len(cr.get('strong_pairs', []))}개")
        P(int(base_progress + per_file_progress * 0.6))

        # 6) 행 조합 — PK 후보 컬럼이 있으면 중복 방지 (replace=False + 행수 고정)
        # 원본 데이터(fd.df) 기준 PK 후보가 하나라도 있으면 dim 테이블로 간주.
        # replace=True로 행을 재샘플링하면 unique 컬럼에도 중복이 생겨
        # 이후 DB INSERT 시 PRIMARY KEY 위반이 발생함.
        has_pk_candidate = any(
            _is_pk_candidate(fd, c) for c in fd.df.columns)
        n = len(syn_num) if len(syn_num) > 0 else nr
        cat_cols = [c for c, t in ct.items()
                    if t == 'categorical' and c in df_text.columns]

        if has_pk_candidate:
            # dim 테이블 보호: 원본 행 수 유지 + replace=False (순서만 셔플)
            n = len(df_text)
            idx = np.random.permutation(len(df_text))
            L(f"  [5/7] 행 조합 — PK 후보 발견: 행 수 {n}로 고정, "
              f"replace=False (중복 방지)")
        elif cat_cols and len(df_text) > 0:
            key_cat = [c for c in cat_cols
                       if df_text[c].nunique() <= max(20, len(df_text) * 0.15)]
            if key_cat:
                freq_col = key_cat[0]
                freqs = df_text[freq_col].value_counts(normalize=True)
                weights = df_text[freq_col].map(freqs).fillna(
                    1.0 / len(df_text)).values
                weights = weights / weights.sum()
                idx = np.random.choice(len(df_text), size=n, replace=True,
                                        p=weights)
            else:
                idx = np.random.choice(len(df_text), size=n,
                                        replace=(n > len(df_text)))
            L(f"  [5/7] 행 조합 완료 ({n}행)")
        else:
            idx = np.random.choice(len(df_text), size=n,
                                    replace=(n > len(df_text)))
            L(f"  [5/7] 행 조합 완료 ({n}행)")

        final = df_text.iloc[idx].reset_index(drop=True)
        for c in syn_num.columns:
            # PK 후보 보호 모드에서 n != len(syn_num)일 수 있으므로 길이 맞춤
            sn = syn_num[c].values
            if len(sn) != n:
                # 수치 컬럼도 같은 idx로 잘라줌 — syn_num의 행 수가 n과 다르면 자르거나 반복
                if len(sn) >= n:
                    final[c] = sn[:n]
                else:
                    # syn_num이 더 짧으면 idx로 보간 (실제로는 거의 발생 안 함)
                    final[c] = np.resize(sn, n)
            else:
                final[c] = sn
        final = final[[c for c in df.columns if c in final.columns]]
        P(int(base_progress + per_file_progress * 0.7))

        # 7) 날짜 정리 + 셔플 + 저장
        date_cols = [c for c, t in ct.items() if t == 'datetime']
        for col in date_cols:
            if col in final.columns:
                try:
                    final[col] = _clean_datetime_column(final[col])
                except Exception:
                    pass

        original_order = list(final.columns)
        shuffled_order = shuffle_columns_safely(original_order, df, ct)
        final = final[shuffled_order]
        fd.shuffle_order = {'original': original_order,
                              'shuffled': shuffled_order}

        # 품질
        ov, scores = validate_quality(df, final, ct)
        L(f"  [6/7] 품질 검증 — 종합 {ov:.1%}")
        P(int(base_progress + per_file_progress * 0.85))

        # 저장 — 파일명은 <프로젝트명>_<테이블명> 형식
        # 테이블명은 게이트에서 영문 소문자로 검증되므로 그대로 사용 가능하지만
        # 윈도우 금지문자만 한 번 더 정제
        safe_base = re.sub(r'[<>:"/\\|?*]', '_',
                            (fd.table_name or '').strip()
                            or os.path.splitext(fd.file_name)[0])
        out_xl = os.path.join(save_dir,
            f"{save_prefix}_{safe_base}.xlsx")
        final.to_excel(out_xl, index=False)
        fd.output_xlsx = out_xl

        # ── 합성 결과의 unique 컬럼 기록 (MD의 PK 판정에 사용) ──
        # 원본 fd.df 기준이 아니라 실제 저장된 데이터를 검사 — 합성 과정의
        # 잔여 중복이 있다면 PK를 적용하지 않도록 보장.
        fd.synth_unique_cols = set()
        for col in final.columns:
            series = final[col]
            if len(series) == 0:
                continue
            if series.isna().any():
                continue
            if series.duplicated().any():
                continue
            fd.synth_unique_cols.add(str(col))
        if fd.synth_unique_cols:
            L(f"        + 합성 후 unique 컬럼 {len(fd.synth_unique_cols)}개 "
              f"(PK 후보로 MD에 표시)")

        # ── 원본 데이터 Parquet 백업 (사용자 요청) ──
        # 합성 데이터(마스킹된 .xlsx)는 외부 공유용이고, 사내 원본은 빠른 재로드를 위해
        # parquet으로 저장. fd.df 는 합성 처리 중 컬럼 rename 등이 적용되므로
        # 원본 파일을 다시 로드.
        try:
            (orig_df_full, _orig_info), _eng = load_excel(
                fd.path, fd.sheet_name or None)
            orig_pq = os.path.join(save_dir,
                f"원본_{safe_base}.parquet")
            orig_df_full.to_parquet(orig_pq, index=False)
            L(f"        + 원본 Parquet: {os.path.basename(orig_pq)} "
              f"({len(orig_df_full)}행)")
        except Exception as pq_err:
            # pyarrow/fastparquet 미설치 시 CSV 폴백
            try:
                if 'orig_df_full' not in locals():
                    (orig_df_full, _), _ = load_excel(
                        fd.path, fd.sheet_name or None)
                orig_csv = os.path.join(save_dir,
                    f"원본_{safe_base}.csv")
                orig_df_full.to_csv(orig_csv, index=False, encoding='utf-8-sig')
                L(f"        + 원본 CSV (Parquet 실패): "
                  f"{os.path.basename(orig_csv)}")
            except Exception as csv_err:
                L(f"        ⚠️ 원본 Parquet/CSV 저장 실패: "
                  f"{pq_err} / {csv_err}")

        # 변환키 JSON
        out_key = os.path.join(save_dir,
            f"{save_prefix}_{safe_base}_변환키.json")
        key_data = {
            'generated_at': datetime.now().isoformat(),
            'employee_id': employee_id,
            'source_file': fd.path,
            'sheet_name': fd.sheet_name,
            'original_columns': [str(c) for c in fd.original_columns],
            'current_columns': [str(c) for c in df.columns],
            'column_rename': {str(k): str(v) for k, v in rename_map.items()},
            'value_mapping': {str(c): mp for c, mp in mapping_dict.items()},
            # OTHER 통합 메타데이터 (복원/필터링용) — OTHER 미발생 시 빈 dict
            'other_buckets': fd.other_buckets,
            'shuffle_order': fd.shuffle_order,
            'quality_overall': ov,
            'quality_by_col': scores,
        }
        with open(out_key, 'w', encoding='utf-8') as f:
            json.dump(key_data, f, ensure_ascii=False, indent=2, default=str)

        # 컬럼 마스크 맵 갱신 (MD 생성에 사용)
        fd.column_mask_map = {str(o): str(rename_map.get(o, o))
                                for o in fd.original_columns}

        L(f"  [7/7] 저장 — {os.path.basename(out_xl)}")
        L(f"        + {os.path.basename(out_key)}")
        P(int(base_progress + per_file_progress))


# ══════════════════════════════════════════════════════════════
# 엔트리 포인트
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        os.chdir(os.path.dirname(sys.executable))
    root = tk.Tk()
    SynthesizeAppV2(root)
    root.mainloop()
