"""
make_sample_data.py
auto_synthesize_gui_dpv2.py 테스트용 샘플 엑셀 4종 생성기.

설계 의도 (앱 기능별 자극 포인트):
- 다중 파일 일괄 마스킹      → 4개 파일
- 컬럼 관계 자동 감지        → 사번 / 프로젝트코드 / 호선번호를 파일 간 공유
- PII 자동 감지(가명 처리)   → 이름, 연락처, 이메일, 회사명(발주사/선주), PM담당자
- ID 컬럼 감지              → 사번(E001..), 프로젝트코드(PRJ-001..), 작업ID, 호선번호
- 날짜/숫자/범주 혼합        → 입사일, 시작일/종료일, 예산, 진행률, 투입공수 등
- 시작<종료 부등식 제약      → 프로젝트 시작일 < 종료일
- null 보존                 → 일부 컬럼에 결측치 삽입

생성 위치: ./samples/*.xlsx
실행:  python make_sample_data.py
"""

import os
import numpy as np
import pandas as pd

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples")
os.makedirs(OUT_DIR, exist_ok=True)

rng = np.random.default_rng(20260612)

# ──────────────────────────────────────────────────────────────
# 공통 마스터 키 (파일 간 관계의 원천)
# ──────────────────────────────────────────────────────────────
N_EMP = 40          # 직원 수
N_PRJ = 8           # 프로젝트 수
N_SHIP = 12         # 호선 수
N_TASK = 60         # 작업 배정 수

emp_ids = [f"E{str(i+1).zfill(3)}" for i in range(N_EMP)]          # E001..E040
prj_codes = [f"PRJ-{str(i+1).zfill(3)}" for i in range(N_PRJ)]     # PRJ-001..PRJ-008
ship_nos = [f"H{2100 + i}" for i in range(N_SHIP)]                 # H2100..H2111

# 한글 이름 풀
SURNAMES = list("김이박최정강조윤장임한오서신권황안송")
GIVEN = ["민준", "서연", "도윤", "지우", "하준", "서윤", "예준", "지민",
         "현우", "수아", "건우", "지호", "유진", "준서", "민서", "지훈",
         "성민", "재현", "은지", "다은", "태경", "승현", "주원", "하늘"]

def rand_names(n, seed):
    r = np.random.default_rng(seed)
    out = set()
    res = []
    while len(res) < n:
        nm = SURNAMES[r.integers(len(SURNAMES))] + GIVEN[r.integers(len(GIVEN))]
        if nm not in out:
            out.add(nm)
            res.append(nm)
    return res

emp_names = rand_names(N_EMP, 1)

DEPTS = ["선체설계", "의장설계", "생산관리", "품질보증", "용접기술", "도장기술"]
RANKS = ["사원", "대리", "과장", "차장", "부장"]
COMPANIES = ["한바다해운", "동방조선해양", "그린오션라인", "태평양상선",
             "북극성쉬핑", "현대글로벌마린", "오리온로지스틱스", "제니스에너지"]
SHIP_TYPES = ["LNG운반선", "컨테이너선", "원유운반선", "벌크선", "FPSO", "셔틀탱커"]
TASK_TYPES = ["기본설계", "상세설계", "생산설계", "현장지원", "검사", "수정작업"]

def inject_null(series, ratio, seed):
    """series에 ratio 비율로 결측치 삽입 (null 보존 테스트)."""
    r = np.random.default_rng(seed)
    s = series.copy()
    n = len(s)
    k = int(n * ratio)
    if k > 0:
        idx = r.choice(n, size=k, replace=False)
        s.iloc[idx] = np.nan
    return s


# ──────────────────────────────────────────────────────────────
# 1) 직원정보.xlsx  (PII: 이름/연락처/이메일,  ID: 사번)
# ──────────────────────────────────────────────────────────────
df_emp = pd.DataFrame({
    "사번": emp_ids,
    "이름": emp_names,
    "부서": [DEPTS[rng.integers(len(DEPTS))] for _ in range(N_EMP)],
    "직급": [RANKS[rng.integers(len(RANKS))] for _ in range(N_EMP)],
    "입사일": pd.to_datetime("2010-01-01") + pd.to_timedelta(
        rng.integers(0, 5500, N_EMP), unit="D"),
    "연락처": [f"010-{rng.integers(1000,9999)}-{rng.integers(1000,9999)}"
              for _ in range(N_EMP)],
    "이메일": [f"{eid.lower()}@shi-sample.com" for eid in emp_ids],
    "기본급": rng.integers(3200, 8200, N_EMP) * 10000,
})
df_emp["직급"] = inject_null(df_emp["직급"], 0.05, 11)

# ──────────────────────────────────────────────────────────────
# 2) 프로젝트.xlsx  (PII: 발주사/PM담당자,  ID: 프로젝트코드,  시작<종료)
# ──────────────────────────────────────────────────────────────
start = pd.to_datetime("2023-01-01") + pd.to_timedelta(
    rng.integers(0, 700, N_PRJ), unit="D")
duration = rng.integers(180, 900, N_PRJ)
end = start + pd.to_timedelta(duration, unit="D")
df_prj = pd.DataFrame({
    "프로젝트코드": prj_codes,
    "프로젝트명": [f"{COMPANIES[i % len(COMPANIES)][:2]}-{SHIP_TYPES[rng.integers(len(SHIP_TYPES))]} 건조"
                for i in range(N_PRJ)],
    "발주사": [COMPANIES[rng.integers(len(COMPANIES))] for _ in range(N_PRJ)],
    "PM담당자": [emp_names[rng.integers(N_EMP)] for _ in range(N_PRJ)],   # 직원 이름과 관계
    "계약일": start - pd.to_timedelta(rng.integers(30, 120, N_PRJ), unit="D"),
    "시작일": start,
    "종료일": end,
    "계약금액": rng.integers(800, 5000, N_PRJ) * 100000000,  # 억 단위
    "진행상태": [["계획", "진행중", "완료"][rng.integers(3)] for _ in range(N_PRJ)],
})

# ──────────────────────────────────────────────────────────────
# 3) 작업배정.xlsx  (관계: 사번→직원, 프로젝트코드→프로젝트,  ID: 작업ID)
# ──────────────────────────────────────────────────────────────
df_task = pd.DataFrame({
    "작업ID": [f"T{str(i+1).zfill(4)}" for i in range(N_TASK)],
    "사번": [emp_ids[rng.integers(N_EMP)] for _ in range(N_TASK)],
    "프로젝트코드": [prj_codes[rng.integers(N_PRJ)] for _ in range(N_TASK)],
    "작업유형": [TASK_TYPES[rng.integers(len(TASK_TYPES))] for _ in range(N_TASK)],
    "투입공수(MH)": rng.integers(8, 400, N_TASK),
    "진행률": rng.integers(0, 101, N_TASK),               # 0~100 range 제약
    "배정일": pd.to_datetime("2024-01-01") + pd.to_timedelta(
        rng.integers(0, 500, N_TASK), unit="D"),
})
df_task["진행률"] = inject_null(df_task["진행률"], 0.08, 33)

# ──────────────────────────────────────────────────────────────
# 4) 호선정보.xlsx  (관계: 프로젝트코드→프로젝트,  PII: 선주,  ID: 호선번호)
# ──────────────────────────────────────────────────────────────
df_ship = pd.DataFrame({
    "호선번호": ship_nos,
    "프로젝트코드": [prj_codes[rng.integers(N_PRJ)] for _ in range(N_SHIP)],
    "선종": [SHIP_TYPES[rng.integers(len(SHIP_TYPES))] for _ in range(N_SHIP)],
    "선주": [COMPANIES[rng.integers(len(COMPANIES))] for _ in range(N_SHIP)],  # 발주사와 관계
    "전장(m)": np.round(rng.uniform(120, 400, N_SHIP), 1),
    "재화중량(DWT)": rng.integers(20000, 320000, N_SHIP),
    "인도예정일": pd.to_datetime("2025-06-01") + pd.to_timedelta(
        rng.integers(0, 900, N_SHIP), unit="D"),
})

# ──────────────────────────────────────────────────────────────
# 저장
# ──────────────────────────────────────────────────────────────
files = {
    "1_직원정보.xlsx": df_emp,
    "2_프로젝트.xlsx": df_prj,
    "3_작업배정.xlsx": df_task,
    "4_호선정보.xlsx": df_ship,
}

for fname, df in files.items():
    path = os.path.join(OUT_DIR, fname)
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"  생성: {path}  ({len(df)}행 x {len(df.columns)}열)")

print("\n[관계 설계 요약]")
print("  사번        : 1_직원정보 ↔ 3_작업배정")
print("  프로젝트코드 : 2_프로젝트 ↔ 3_작업배정 ↔ 4_호선정보 (3-파일 그룹)")
print("  회사명      : 2_프로젝트.발주사 ↔ 4_호선정보.선주 (값 중복 기반)")
print("  PM담당자/이름: 2_프로젝트.PM담당자 ↔ 1_직원정보.이름 (값 중복 기반)")
print(f"\n총 4개 파일 → {OUT_DIR}")
