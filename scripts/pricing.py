"""신청명단의 과정명을 환급/비환급과정 커리큘럼과 대조해 금액을 계산하는 공통 로직.

process_apply_list.py, generate_company_files.py, build_gaegang_sheet.py가
공통으로 사용한다. 금액 관련 규칙은 모두 사용자가 직접 확인해준 내용이다:

- 과정명이 환급과정 커리큘럼에 있으면 "환급"(과정유형: 사업주).
- 없고 비환급과정 커리큘럼에만 있으면 "비환급"(과정유형: 일반).
- 양쪽에 모두 있으면 항상 "환급"으로 표기하고 금액도 환급 쪽 금액을 쓴다. 이때
  신청명단의 "구분" 칸에 "일반"이 들어있어 실제로 비환급으로 신청한 경우에만
  "비환급 -> 환급으로 변경" 메모를 남긴다(원래부터 환급/사업주로 신청했으면
  메모를 남기지 않는다).
- 비환급 과정은 정부지원금이 없으므로 기업부담금 = 수강료 전액.
- 어느 쪽에서도 과정명을 찾지 못하면 금액을 비우고 확인필요로만 표시한다
  (추측해서 채우지 않는다).
"""
import re
import unicodedata
from datetime import date

import pandas as pd
from openpyxl.styles import Alignment

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


def exclusion_reason(row) -> str:
    """출처에 '프리패스'가 들어있거나, 신청과정명에 '카드결제'(띄어쓰기 무관)가
    들어있거나, 아이디/회사명에 '테스트'/'test'가 들어있거나, 신청상태가
    '취소'인 신청 건이면 그 사유를 돌려주고, 아니면 빈 문자열을 돌려준다."""
    source = normalize(row.get("출처", ""))
    if "프리패스" in source:
        return "프리패스"
    course_no_space = re.sub(r"\s+", "", str(row.get("신청과정명", "")))
    if "카드결제" in course_no_space:
        return "카드결제"
    id_text = str(row.get("아이디", "")).lower()
    company_text = str(row.get("회사명", "")).lower()
    for keyword in EXCLUDE_TEST_KEYWORDS:
        if keyword in id_text or keyword in company_text:
            return "테스트"
    if normalize(row.get("신청상태", "")) == "취소":
        return "취소"
    return ""


def is_excluded_row(row) -> bool:
    return bool(exclusion_reason(row))


def filter_excluded_rows(df: pd.DataFrame):
    """출처='프리패스' 건, 아이디/회사명에 테스트 흔적이 있는 건,
    신청과정명에 '카드결제'가 들어간 건을 제외한다.
    (신청명단을 쓰는 1·2·3단계 모두 공통 적용)

    반환값: (제외하고 남은 df, 제외된 건만 모은 df("제외사유" 칸 포함), 제외된 건수)
    """
    reasons = df.apply(exclusion_reason, axis=1)
    mask = reasons != ""
    kept = df[~mask].reset_index(drop=True)
    excluded = df[mask].reset_index(drop=True).copy()
    excluded["제외사유"] = reasons[mask].reset_index(drop=True)
    return kept, excluded, int(mask.sum())


def filter_by_date_range(df: pd.DataFrame, start_date=None, end_date=None):
    """교육시작일이 start_date~end_date(둘 다 포함) 범위 안에 있는 줄만 남긴다.
    start_date/end_date가 둘 다 비어 있으면 아무것도 거르지 않고 전체를 돌려준다
    (하나만 비어 있으면 그쪽 경계는 적용하지 않는다).

    반환값: (범위 안에 있는 df, 범위 밖이라 제외된 건수)
    """
    if not start_date and not end_date:
        return df.reset_index(drop=True), 0

    education_start = pd.to_datetime(df["교육시작일"], errors="coerce")
    mask = pd.Series(True, index=df.index)
    if start_date:
        mask &= education_start >= pd.to_datetime(start_date)
    if end_date:
        mask &= education_start <= pd.to_datetime(end_date)

    kept = df[mask].reset_index(drop=True)
    return kept, int((~mask).sum())


DUPLICATE_NOTE = "기존 명단에 같은 아이디·과정명 신청이 이미 있음 - 중복 신청 확인필요"


def added_note(is_added: bool) -> str:
    """추가 신청으로 들어온 줄이면 "YYYY-MM-DD 추가 신청" 메모를, 아니면 빈 문자열을 돌려준다."""
    return f"{date.today().isoformat()} 추가 신청" if is_added else ""


def read_apply_list(apply_path, required_columns, added_path=None) -> pd.DataFrame:
    """신청명단을 읽는다. added_path(추가 신청 파일)가 있으면 원본 뒤에 이어 붙인다.

    내부용 칸 2개를 붙여서 돌려준다(결과 파일에는 나가지 않음):
    - "_추가": 추가 신청 파일에서 온 줄이면 True
    - "_중복메모": 추가 줄이 원본에 이미 있는 아이디+과정명이면 DUPLICATE_NOTE
    필요한 칸이 없으면 ValueError를 발생시킨다(추측해서 넘어가지 않음)."""
    df = pd.read_excel(apply_path)
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        raise ValueError(f"신청명단 파일에 다음 칸이 없습니다. 원본 파일 형식을 확인해주세요: {missing}")
    df["_추가"] = False
    df["_중복메모"] = ""
    if not added_path:
        return df

    added = pd.read_excel(added_path)
    missing = [c for c in required_columns if c not in added.columns]
    if missing:
        raise ValueError(f"추가 신청 파일에 다음 칸이 없습니다. 원본 신청명단과 같은 칸이어야 합니다: {missing}")
    base_keys = {(normalize(i), normalize(c)) for i, c in zip(df["아이디"], df["신청과정명"])}
    added["_추가"] = True
    added["_중복메모"] = [DUPLICATE_NOTE if (normalize(i), normalize(c)) in base_keys else ""
                         for i, c in zip(added["아이디"], added["신청과정명"])]
    return pd.concat([df, added], ignore_index=True)


# 이 회사 신청 건은 1·2·3단계 결과에서 항상 빼고, 필요하면 별도 파일로 만든다(tel_list.py).
TEL_COMPANY_NAME = "도쿄일렉트론코리아"


def is_tel_row(row) -> bool:
    """회사명(법인형태·띄어쓰기 제거)이 '도쿄일렉트론코리아'와 정확히 같으면 True."""
    return re.sub(r"\s+", "", clean_company_name(row.get("회사명", ""))) == TEL_COMPANY_NAME


def split_tel_rows(df: pd.DataFrame):
    """반환값: (도쿄일렉트론코리아 건을 뺀 df, 도쿄일렉트론코리아 건만 모은 df)"""
    mask = df.apply(is_tel_row, axis=1) if len(df) else pd.Series(dtype=bool)
    return df[~mask].reset_index(drop=True), df[mask].reset_index(drop=True)


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


def classify_course(course_name, refund_table: dict, general_table: dict, application_type: str = "") -> dict:
    """1인 기준 과정 분류·금액을 계산해 돌려준다.

    application_type: 신청명단의 "구분" 칸 값. 여기에 "일반"이 들어있으면 학생이
    실제로 비환급(일반)으로 신청했다는 뜻이다. 과정명이 환급·비환급 양쪽 커리큘럼에
    모두 있으면 항상 "환급"으로 처리하되, 실제로 비환급으로 신청한 경우에만
    "비환급 -> 환급으로 변경" 메모를 남긴다(원래부터 환급/사업주로 신청했으면
    양쪽에 다 있어도 전환된 게 아니므로 메모를 남기지 않는다).

    반환값: {"환급구분", "과정유형", "수강료", "정부지원금", "기업부담금", "확인필요"}
    가격을 알 수 없으면 "수강료"/"정부지원금"/"기업부담금"은 None이다.
    """
    course_norm = normalize(course_name)
    in_refund = course_norm in refund_table
    in_general = course_norm in general_table
    applied_as_general = "일반" in normalize(application_type)

    note = ""
    if in_refund:
        price = refund_table[course_norm]
        course_type = "사업주"
        refund_flag = "환급"
        if in_general and applied_as_general:
            note = "비환급 -> 환급으로 변경"
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


CURRENCY_FORMAT = "#,##0"
TEXT_FORMAT = "@"
_NO_WRAP = Alignment(wrap_text=False)


def _display_width(value) -> float:
    """한글은 영문보다 넓게 보이므로, 동아시아 넓은 글자를 2칸으로 센 길이를 돌려준다."""
    width = 0
    for ch in str(value):
        width += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return width


def style_worksheet(ws, currency_columns=(), text_columns=()):
    """엑셀 결과 파일 공통 서식을 적용한다.
    - currency_columns: 천단위 콤마(#,##0)로 표시할 칸 이름들
    - text_columns: 텍스트 서식(@)으로 고정할 칸 이름들(예: 아이디)
    - 모든 칸: 텍스트 줄바꿈 끄기 + 글자 길이에 맞춰 열 너비 자동 조정
    """
    headers = [c.value for c in ws[1]]
    header_idx = {h: i + 1 for i, h in enumerate(headers) if h}

    for col_name in currency_columns:
        idx = header_idx.get(col_name)
        if not idx:
            continue
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=idx).number_format = CURRENCY_FORMAT

    for col_name in text_columns:
        idx = header_idx.get(col_name)
        if not idx:
            continue
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=idx).number_format = TEXT_FORMAT

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = _NO_WRAP

    for col_cells in ws.columns:
        letter = col_cells[0].column_letter
        max_width = max((_display_width(c.value) for c in col_cells if c.value is not None), default=0)
        ws.column_dimensions[letter].width = max_width + 2
