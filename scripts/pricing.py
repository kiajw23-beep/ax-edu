"""신청명단의 과정명을 환급/비환급 가격표와 대조해 금액을 계산하는 공통 로직.

process_apply_list.py, generate_company_files.py, build_gaegang_sheet.py가
공통으로 사용한다. 금액 관련 규칙은 모두 사용자가 직접 확인해준 내용이다:

- 과정명이 환급과정 커리큘럼에 있으면 "환급"(과정유형: 사업주).
- 없고 비환급 가격표에만 있으면 "비환급"(과정유형: 일반).
- 양쪽에 모두 있으면 "환급"으로 표기하고 금액도 환급 쪽 금액을 쓴다. 이 경우
  "일반→환급 전환" 메모를 남긴다.
- 비환급 과정은 정부지원금이 없으므로 기업부담금 = 수강료 전액.
- 어느 쪽에서도 과정명을 찾지 못하면 금액을 비우고 확인필요로만 표시한다
  (추측해서 채우지 않는다).
"""
import re

import pandas as pd

COMPANY_SUFFIXES = ["주식회사", "(주)", "㈜", "유한회사", "유한책임회사", "합자회사", "합명회사"]


def clean_company_name(raw) -> str:
    name = str(raw).strip()
    for suffix in COMPANY_SUFFIXES:
        name = name.replace(suffix, "")
    return re.sub(r"\s+", " ", name).strip()


def normalize(text) -> str:
    return re.sub(r"\s+", " ", str(text)).strip()


def find_column(columns, keyword):
    return next((c for c in columns if isinstance(c, str) and keyword in c), None)


EXCLUDE_TEST_KEYWORDS = ["테스트", "test"]


def is_excluded_row(row) -> bool:
    """출처에 '프리패스'라는 단어가 들어있거나, 아이디/회사명에 '테스트'/'test'가
    들어간 신청 건이면 True."""
    source = normalize(row.get("출처", ""))
    if "프리패스" in source:
        return True
    id_text = str(row.get("아이디", "")).lower()
    company_text = str(row.get("회사명", "")).lower()
    for keyword in EXCLUDE_TEST_KEYWORDS:
        if keyword in id_text or keyword in company_text:
            return True
    return False


def filter_excluded_rows(df: pd.DataFrame):
    """출처='프리패스' 건, 아이디/회사명에 테스트 흔적이 있는 건을 제외한다.
    (신청명단을 쓰는 1·2·3단계 모두 공통 적용)

    반환값: (제외하고 남은 df, 제외된 건수)
    """
    mask = df.apply(is_excluded_row, axis=1)
    kept = df[~mask].reset_index(drop=True)
    return kept, int(mask.sum())


def load_price_table(path: str) -> dict:
    """과정명 -> {"수강료", "정부지원금", "기업부담금"} 사전을 만든다.
    같은 과정명이 서로 다른 금액으로 중복 등록돼 있으면 {"중복": True}로 표시한다."""
    df = pd.read_csv(path) if path.lower().endswith(".csv") else pd.read_excel(path)

    name_col = find_column(df.columns, "과정명")
    fee_col = find_column(df.columns, "수강료")
    gov_col = find_column(df.columns, "정부지원금")
    company_col = find_column(df.columns, "기업부담금")

    missing = [label for label, col in [("과정명", name_col), ("수강료", fee_col),
                                         ("정부지원금", gov_col), ("기업부담금", company_col)]
               if col is None]
    if missing:
        raise ValueError(f"가격표 파일({path})에 다음 칸을 찾을 수 없습니다: {missing}")

    table = {}
    for _, row in df.iterrows():
        name = normalize(row[name_col])
        if not name:
            continue
        record = {"수강료": row[fee_col], "정부지원금": row[gov_col], "기업부담금": row[company_col]}
        if name in table and not table[name].get("중복") and table[name]["수강료"] != record["수강료"]:
            table[name] = {"중복": True}
        elif name not in table:
            table[name] = record
    return table


def classify_course(course_name, refund_table: dict, general_table: dict) -> dict:
    """1인 기준 과정 분류·금액을 계산해 돌려준다.

    반환값: {"환급구분", "과정유형", "수강료", "정부지원금", "기업부담금", "확인필요"}
    가격을 알 수 없으면 "수강료"/"정부지원금"/"기업부담금"은 None이다.
    """
    course_norm = normalize(course_name)
    in_refund = course_norm in refund_table
    in_general = course_norm in general_table

    note = ""
    if in_refund:
        price = refund_table[course_norm]
        course_type = "사업주"
        refund_flag = "환급"
        if in_general:
            note = "일반과정에도 등록되어 있으나 환급과정 우선 적용(일반→환급 전환)"
    elif in_general:
        price = general_table[course_norm]
        course_type = "일반"
        refund_flag = "비환급"
    else:
        price = None
        course_type = "확인필요"
        refund_flag = "확인필요"
        note = "환급/비환급 커리큘럼 어디에도 과정명이 없음 - 금액 직접 확인 필요"

    if price is not None and price.get("중복"):
        price = None
        note = (note + " / " if note else "") + "같은 과정명이 가격표에 서로 다른 금액으로 중복 등록되어 금액 확인 필요"

    if price is None:
        tuition = gov = company_burden = None
    else:
        tuition = price["수강료"]
        if refund_flag == "환급":
            gov = price["정부지원금"]
            company_burden = price["기업부담금"]
        else:  # 비환급: 정부지원금 없음 -> 기업부담금 = 수강료 전액
            gov = 0
            company_burden = tuition

    return {
        "환급구분": refund_flag,
        "과정유형": course_type,
        "수강료": tuition,
        "정부지원금": gov,
        "기업부담금": company_burden,
        "확인필요": note,
    }
