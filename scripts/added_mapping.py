"""서식이 제각각인 추가 신청 파일(또는 붙여넣은 표)을 원본 신청명단과 같은 칸(표준 양식)으로 바꾼다.

웹 화면의 "칸 맞추기"에서 쓴다. 사용자가 시트와 제목 줄(칸 이름이 있는 줄)을 고르고,
원본 신청명단의 칸마다 추가 신청 파일의 어느 칸을 쓸지 직접 정한다.
- 칸 이름이 같거나 같은 뜻 이름 목록(COLUMN_ALIASES)에 있거나, 서로 포함하면 자동으로 골라 둔다
  (띄어쓰기·대소문자 무시). 칸 이름이 있는 줄도 위쪽 줄 중 가장 많이 맞는 줄로 골라 둔다.
  자동으로 고른 것은 화면에서 드롭다운으로 언제든 바꿀 수 있다.
- 필수 칸(REQUIRED_FOR_ADDED)은 반드시 맞춰야 하고, 나머지 칸은 맞출 곳이 없으면 비워 둔다
  (값을 지어내서 채우지 않음).
- 값은 원본 그대로 옮기고 고치지 않는다(신청과정명 포함). 단, 교육시작일·교육종료일은 개강내역에서
  원본 신청명단(날짜 서식 셀)과 같은 날끼리 묶이도록 날짜로 읽어 맞춘다. 날짜로 읽지 못한 값은
  그대로 두고 확인필요 메모(pricing.INPUT_NOTE_COLUMN)를 남긴다. 연도가 없는 날짜는 추측하지 않는다.
  교육종료일은 필수가 아니라서 비어 있으면 그냥 비워 둔다.
- 생년월일에 주민등록번호가 들어오면 앞 6자리만 날짜(1994-03-19)로 바꾸고 뒷자리는 버린다. 세기는
  7번째 자리(성별 자리)로만 정하고, 그 자리가 없으면 추측하지 않고 앞 6자리 그대로 두고 확인필요로 남긴다.
칸 맞추기 결과는 바로 반영하지 않고 "확인 표"(build_review_rows/check_review_rows)로 보여 준 뒤, 사람이
고친 내용 그대로 표준 양식 파일(save_reviewed)을 만든다. 확인 표 검사는 기존 규칙(필수 칸, 날짜, 커리큘럼
과정명, 중복, 제외 규칙, 도쿄일렉트론코리아)을 미리 보여 주는 것뿐이고 새 판단을 하지 않는다. 확인 표에서
직접 고친 칸·직접 추가한 줄은 확인필요 메모로 남긴다.
만든 표준 양식 파일은 기존 add_applications.run()에 그대로 넘긴다.
"""
import csv
import io
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pricing import (  # noqa: E402
    INPUT_NOTE_COLUMN, exclusion_reason, is_tel_row, load_price_table, normalize,
)
from process_apply_list import REQUIRED_COLUMNS  # noqa: E402

# 이 칸들이 없으면 기업명·환급 판정·개강내역 묶음을 만들 수 없어서 꼭 맞춰야 한다.
REQUIRED_FOR_ADDED = ["회사명", "성함", "아이디", "신청과정명", "교육시작일"]

# 칸 이름 자동 선택에 쓰는 같은 뜻 이름(띄어쓰기·대소문자 무시). 원본 칸 이름 자체는 따로 넣지 않아도 된다.
COLUMN_ALIASES = {
    "회사명": ["회사", "기업명", "기업", "업체명", "소속", "고객사"],
    "성함": ["이름", "성명", "수강생", "수강자", "교육생"],
    "아이디": ["id", "계정", "회원id", "로그인id"],
    "신청과정명": ["과정", "과정명", "교육명", "교육과정", "강의명", "강좌명"],
    "교육시작일": ["시작일", "시작", "개강일", "교육시작", "수강시작일"],
    "교육종료일": ["종료일", "종료", "교육종료", "수강종료일"],
    "휴대전화": ["연락처", "전화번호", "휴대폰", "핸드폰", "전화"],
    "메일주소": ["이메일", "메일", "e-mail", "email"],
    "생년월일": ["생일", "주민등록번호", "주민번호"],
    "구분": ["신청구분", "과정유형"],
}

# 칸 이름이 있는 줄을 찾을 때 볼 위쪽 줄 수
HEADER_SEARCH_ROWS = 20

# 주민등록번호 앞 6자리 + 7번째 자리(성별·세기). 뒷자리가 *로 가려져 있어도 된다.
_RRN_PATTERN = re.compile(r"^(\d{2})(\d{2})(\d{2})\s*-?\s*(\d)?[\d*]*$")
_RRN_CENTURY = {"1": 1900, "2": 1900, "5": 1900, "6": 1900, "3": 2000, "4": 2000, "7": 2000, "8": 2000,
                "9": 1800, "0": 1800}

DATE_COLUMNS = ["교육시작일", "교육종료일"]

# 붙여넣은 표는 시트가 하나뿐이라 이 이름으로 부른다.
PASTE_SHEET = "붙여넣기"

# 2026-11-02 / 2026.11.02 / 2026. 11. 2. / 2026/11/02 (뒤에 00:00 또는 00:00:00이 붙어도 됨) / 20261102
_DATE_PATTERN = re.compile(r"^(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})\.?(?:\s+00:00(?::00)?)?$")
_COMPACT_DATE_PATTERN = re.compile(r"^(\d{4})(\d{2})(\d{2})$")


def read_sheets(path: str) -> dict:
    """시트 이름 -> 칸 이름 없이 읽은 표. CSV는 시트 하나("CSV")로 본다."""
    if path.lower().endswith(".csv"):
        try:
            return {"CSV": pd.read_csv(path, header=None, dtype=object, encoding="utf-8-sig")}
        except UnicodeDecodeError:
            # 한국어 윈도우 엑셀에서 저장한 CSV
            return {"CSV": pd.read_csv(path, header=None, dtype=object, encoding="cp949")}
    return pd.read_excel(path, sheet_name=None, header=None)


def read_pasted(text: str) -> dict:
    """엑셀·메일에서 복사해 붙여넣은 표(칸 사이가 탭)를 read_sheets()와 같은 모양으로 읽는다.
    탭이 하나도 없으면 표가 아니라고 보고 ValueError를 낸다."""
    if "\t" not in text:
        raise ValueError("붙여넣은 내용이 표 형식이 아닙니다(칸 구분이 없음).")
    rows = list(csv.reader(io.StringIO(text.strip("\r\n")), delimiter="\t"))
    width = max(len(r) for r in rows)
    rows = [[(cell.strip() or None) for cell in r] + [None] * (width - len(r)) for r in rows]
    return {PASTE_SHEET: pd.DataFrame(rows, dtype=object)}


def to_date(value):
    """교육시작일·교육종료일 값을 날짜로 읽는다. 반환값: (날짜 또는 원래 값, 읽었는지 여부).
    엑셀 날짜 서식 셀은 그대로 날짜로 쓰고, 글자는 연·월·일이 다 있는 모양만 읽는다."""
    if isinstance(value, datetime):
        return pd.Timestamp(value), True
    text = _cell_text(value)
    match = _DATE_PATTERN.match(text) or _COMPACT_DATE_PATTERN.match(text)
    if match:
        try:
            return pd.Timestamp(datetime(*(int(g) for g in match.groups()))), True
        except ValueError:  # 2026-13-40 같은 없는 날짜
            pass
    return value, False


def to_birth_date(value):
    """생년월일 칸 값이 주민등록번호면 앞 6자리를 날짜 글자(1994-03-19)로 바꾼다.
    반환값: (바꾼 값, 확인필요 메모 또는 "")."""
    match = _RRN_PATTERN.match(_cell_text(value))
    if not match:
        return value, ""
    yy, mm, dd, gender = match.groups()
    if gender is None:
        return f"{yy}{mm}{dd}", "생년월일(주민등록번호 앞 6자리)의 연도(1900/2000년대)를 알 수 없음 - 확인필요"
    try:
        return datetime(_RRN_CENTURY[gender] + int(yy), int(mm), int(dd)).strftime("%Y-%m-%d"), ""
    except ValueError:
        return f"{yy}{mm}{dd}", "생년월일(주민등록번호 앞 6자리)이 없는 날짜임 - 확인필요"


def convert_record(record: dict) -> list:
    """표준 양식 한 줄(dict)의 교육일·생년월일을 바꿔 넣고, 확인필요 메모 목록을 돌려준다."""
    notes = []
    for col in DATE_COLUMNS:
        value = record.get(col)
        if not _cell_text(value):
            record[col] = None
            if col in REQUIRED_FOR_ADDED:
                notes.append(f"{col} 비어 있음 - 확인필요")
            continue
        date, ok = to_date(value)
        record[col] = date
        if not ok:
            notes.append(f"{col}을 날짜로 읽지 못함(입력값: {_cell_text(value)}) - 날짜 확인필요")
    if _cell_text(record.get("생년월일")):
        record["생년월일"], note = to_birth_date(record["생년월일"])
        if note:
            notes.append(note)
    return notes


def _cell_text(value) -> str:
    return "" if pd.isna(value) else str(value).strip()


def _plain(text) -> str:
    return re.sub(r"\s+", "", str(text)).lower()


def _match_score(name: str, target: str) -> int:
    """칸 이름이 원본 칸과 얼마나 맞는지: 3 같은 이름, 2 같은 뜻 이름, 1 서로 포함, 0 안 맞음."""
    name = _plain(name)
    if not name:
        return 0
    if name == _plain(target):
        return 3
    aliases = [_plain(a) for a in COLUMN_ALIASES.get(target, [])]
    if name in aliases:
        return 2
    if any(word in name or name in word for word in [_plain(target)] + aliases if len(word) >= 2 and len(name) >= 2):
        return 1
    return 0


def suggest_mapping(header_names: list) -> dict:
    """{원본 칸 이름: 칸 index}. 잘 맞는 짝부터 하나씩 정하고, 칸 하나는 원본 칸 하나에만 쓴다."""
    pairs = sorted(((_match_score(name, target), -REQUIRED_COLUMNS.index(target), -idx, target, idx)
                    for idx, name in enumerate(header_names) for target in REQUIRED_COLUMNS), reverse=True)
    suggested, used = {}, set()
    for score, _, _, target, idx in pairs:
        if score and target not in suggested and idx not in used:
            suggested[target] = idx
            used.add(idx)
    return suggested


def guess_header_row(sheets: dict, sheet: str) -> int:
    """위쪽 줄 중 원본 칸 이름과 가장 많이 맞는 줄(1부터 셈). 하나도 안 맞으면 1."""
    best_row, best_count = 1, 0
    for i, row in enumerate(sheets[sheet].head(HEADER_SEARCH_ROWS).itertuples(index=False)):
        count = len(suggest_mapping([_cell_text(v) for v in row]))
        if count > best_count:
            best_row, best_count = i + 1, count
    return best_row


def sheet_info(sheets: dict, sheet: str, header_row: int) -> dict:
    """화면에 보여줄 정보. header_row는 1부터 센다.

    반환값: {"sheets": 시트 이름 목록, "row_count": 그 시트 줄 수,
             "columns": [{"index", "label"}] (제목 줄의 칸들),
             "suggested": {원본 칸 이름: 추천 칸 index},
             "first_rows": 시트 맨 위 몇 줄(제목 줄 고를 때 참고용)}"""
    df = sheets[sheet]
    columns = []
    suggested = {}
    if 1 <= header_row <= len(df):
        names = [_cell_text(v) for v in df.iloc[header_row - 1]]
        columns = [{"index": idx, "label": f"{_column_letter(idx)}열: {name or '(이름 없음)'}"}
                   for idx, name in enumerate(names)]
        suggested = suggest_mapping(names)
    first_rows = [[_cell_text(v) for v in row] for row in df.head(8).itertuples(index=False)]
    return {"sheets": list(sheets), "row_count": len(df), "columns": columns,
            "suggested": suggested, "first_rows": first_rows}


def _column_letter(idx: int) -> str:
    letters = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def build_standard(sheets: dict, sheet: str, header_row: int, mapping: dict) -> pd.DataFrame:
    """mapping: {원본 칸 이름: 추가 신청 파일의 칸 index 또는 직접 입력한 값(글자)}. 없는 칸은 비워 둔다.
    직접 입력한 값(예: 파일에 교육시작일 칸이 없어 사용자가 정한 날짜)은 모든 줄에 똑같이 넣는다.
    제목 줄 아래에서 파일에서 맞춘 칸이 전부 비어 있는 줄(빈 줄·합계 아래 여백 등)은 뺀다."""
    missing = [c for c in REQUIRED_FOR_ADDED if mapping.get(c) is None]
    if missing:
        raise ValueError(f"추가 신청 파일에서 다음 칸을 맞춰야 합니다: {missing}")
    data = sheets[sheet].iloc[header_row:].reset_index(drop=True)

    def column(c):
        source = mapping.get(c)
        if source is None:
            return pd.Series([None] * len(data), dtype=object)
        if isinstance(source, str):
            return pd.Series([source] * len(data), dtype=object)
        return data[source]

    out = pd.DataFrame({c: column(c) for c in REQUIRED_COLUMNS})
    mapped_cols = [c for c in REQUIRED_COLUMNS if mapping.get(c) is not None and not isinstance(mapping[c], str)]
    not_empty = out[mapped_cols].apply(lambda r: any(_cell_text(v) for v in r), axis=1)
    out = out[not_empty].reset_index(drop=True)

    records = out.to_dict("records")
    notes = [" / ".join(convert_record(record)) for record in records]
    out = pd.DataFrame(records, columns=REQUIRED_COLUMNS, dtype=object)
    out[INPUT_NOTE_COLUMN] = notes
    return out


def save_standard(sheets: dict, sheet: str, header_row: int, mapping: dict, output_path: str, log=print) -> str:
    """칸 맞추기 화면에서 읽어 둔 표(read_sheets()/read_pasted() 결과)를 표준 양식으로 바꿔
    output_path에 저장하고 그 경로를 돌려준다."""
    df = build_standard(sheets, sheet, header_row, mapping)
    df.to_excel(output_path, index=False)
    blank = [c for c in REQUIRED_COLUMNS if mapping.get(c) is None]
    source = "붙여넣기" if sheet == PASTE_SHEET else f"시트 '{sheet}'"
    log(f"칸 맞추기: 추가 신청 {len(df)}건 ({source}, 제목 줄 {header_row}번째 줄)")
    if blank:
        log(f"맞출 칸이 없어 비워 둔 칸: {', '.join(blank)}")
    bad_dates = int((df[INPUT_NOTE_COLUMN] != "").sum())
    if bad_dates:
        log(f"교육시작일·교육종료일을 날짜로 읽지 못한 추가 신청 {bad_dates}건 - 확인필요로 표시")
    return output_path


# ----- 확인 표 -----

def _review_text(value) -> str:
    """확인 표에 보여 줄 글자. 자정인 날짜는 날짜만 보여 준다."""
    if isinstance(value, datetime):
        stamp = pd.Timestamp(value)
        return stamp.strftime("%Y-%m-%d") if stamp == stamp.normalize() else str(stamp)
    return _cell_text(value)


def review_columns(mapping: dict) -> list:
    """확인 표에 보여 줄 칸: 필수 칸 + 칸 맞추기에서 고른 칸(원본 신청명단 칸 순서)."""
    return [c for c in REQUIRED_COLUMNS if c in REQUIRED_FOR_ADDED or mapping.get(c) is not None]


def build_review_rows(sheets: dict, sheet: str, header_row: int, mapping: dict) -> list:
    """칸 맞추기 결과를 확인 표 줄 목록으로 만든다: [{칸 이름: 글자}]. 날짜는 B단계 규칙으로 맞춘 값이다."""
    df = build_standard(sheets, sheet, header_row, mapping)
    columns = review_columns(mapping)
    return [{c: _review_text(row[c]) for c in columns} for _, row in df.iterrows()]


def load_check_context(apply_path: str, refund_path: str, general_path: str) -> dict:
    """확인 표 검사에 쓰는 원본 신청명단의 (아이디, 과정명)과 커리큘럼 과정명 목록."""
    apply_df = pd.read_excel(apply_path)
    refund_table = load_price_table(refund_path)
    general_table = load_price_table(general_path)
    return {
        "base_keys": {(normalize(i), normalize(c)) for i, c in zip(apply_df["아이디"], apply_df["신청과정명"])},
        "course_names": set(refund_table) | set(general_table),
    }


def course_names(context: dict) -> list:
    return sorted(context["course_names"])


def check_review_rows(rows: list, context: dict) -> list:
    """줄마다 {"cells": 노란색으로 표시할 칸 이름들, "messages": 확인 내용}을 돌려준다."""
    keys = [(normalize(r.get("아이디", "")), normalize(r.get("신청과정명", ""))) for r in rows]
    results = []
    for row, key in zip(rows, keys):
        cells, messages = [], []
        empty = [c for c in REQUIRED_FOR_ADDED if not _cell_text(row.get(c))]
        if empty:
            cells += empty
            messages.append(f"필수 칸이 비어 있음: {', '.join(empty)}")
        for col in DATE_COLUMNS:
            if _cell_text(row.get(col)) and not to_date(row.get(col))[1]:
                cells.append(col)
                messages.append(f"{col}을 날짜로 읽지 못함")
        if _cell_text(row.get("생년월일")):
            birth_note = to_birth_date(row.get("생년월일"))[1]
            if birth_note:
                cells.append("생년월일")
                messages.append(birth_note.replace(" - 확인필요", ""))
        course = _cell_text(row.get("신청과정명"))
        if course and normalize(course) not in context["course_names"]:
            cells.append("신청과정명")
            messages.append("과정명이 환급·비환급 커리큘럼 어디에도 없음")
        if all(key) and key in context["base_keys"]:
            messages.append("원본 신청명단에 같은 아이디·과정명 신청이 이미 있음")
        if all(key) and keys.count(key) > 1:
            messages.append("확인 표 안에 같은 아이디·과정명 줄이 또 있음")
        reason = exclusion_reason(row)
        if reason:
            messages.append(f"제외 규칙({reason})에 해당 - 이 줄은 결과에서 빠집니다")
        if is_tel_row(row):
            messages.append("도쿄일렉트론코리아 신청 - 1~3단계 결과에서 빠집니다")
        results.append({"cells": cells, "messages": messages})
    return results


def save_reviewed(rows: list, output_path: str, deleted: int = 0, log=print) -> str:
    """확인 표 내용을 표준 양식으로 저장한다. rows: [{"values": {칸: 글자}, "original": {칸: 글자} 또는 None}]
    (original이 None이면 확인 표에서 직접 추가한 줄). 직접 고친 칸·추가한 줄은 확인필요 메모로 남긴다."""
    records, notes = [], []
    edited = added = 0
    for r in rows:
        values, original = r["values"], r.get("original")
        record = {c: (_cell_text(values.get(c)) or None) for c in REQUIRED_COLUMNS}
        note = convert_record(record)
        if original is None:
            added += 1
            note.append("확인 표에서 직접 추가한 줄")
        else:
            changes = [f"{c}(원래 값: {_cell_text(original.get(c)) or '빈칸'})" for c in values
                       if _cell_text(values.get(c)) != _cell_text(original.get(c))]
            if changes:
                edited += 1
                note.append("확인 표에서 직접 수정: " + ", ".join(changes))
        records.append(record)
        notes.append(" / ".join(note))
    df = pd.DataFrame(records, columns=REQUIRED_COLUMNS, dtype=object)
    df[INPUT_NOTE_COLUMN] = notes
    df.to_excel(output_path, index=False)
    log(f"확인 표: 추가 신청 {len(df)}건 반영 (직접 수정 {edited}건, 직접 추가 {added}건, 삭제 {deleted}건)")
    return output_path
