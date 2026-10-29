"""서식이 제각각인 추가 신청 파일(또는 붙여넣은 표)을 원본 신청명단과 같은 칸(표준 양식)으로 바꾼다.

웹 화면의 "칸 맞추기"에서 쓴다. 사용자가 시트와 제목 줄(칸 이름이 있는 줄)을 고르고,
원본 신청명단의 칸마다 추가 신청 파일의 어느 칸을 쓸지 직접 정한다.
- 자동 추천은 칸 이름이 (띄어쓰기만 빼고) 똑같을 때만 한다. 비슷한 이름을 추측해서 맞추지 않는다.
- 필수 칸(REQUIRED_FOR_ADDED)은 반드시 맞춰야 하고, 나머지 칸은 맞출 곳이 없으면 비워 둔다
  (값을 지어내서 채우지 않음).
- 값은 원본 그대로 옮기고 고치지 않는다(신청과정명 포함). 단, 교육시작일·교육종료일은 개강내역에서
  원본 신청명단(날짜 서식 셀)과 같은 날끼리 묶이도록 날짜로 읽어 맞춘다. 날짜로 읽지 못한 값은
  그대로 두고 확인필요 메모(pricing.INPUT_NOTE_COLUMN)를 남긴다. 연도가 없는 날짜는 추측하지 않는다.
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
REQUIRED_FOR_ADDED = ["회사명", "성함", "아이디", "신청과정명", "교육시작일", "교육종료일"]

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


def _cell_text(value) -> str:
    return "" if pd.isna(value) else str(value).strip()


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
        header = df.iloc[header_row - 1]
        for idx, value in enumerate(header):
            name = _cell_text(value)
            columns.append({"index": idx, "label": f"{_column_letter(idx)}열: {name or '(이름 없음)'}"})
            for target in REQUIRED_COLUMNS:
                if name and re.sub(r"\s+", "", name) == target and target not in suggested:
                    suggested[target] = idx
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
    """mapping: {원본 칸 이름: 추가 신청 파일의 칸 index}. 없는 칸은 비워 둔다.
    제목 줄 아래에서 맞춘 칸이 전부 비어 있는 줄(빈 줄·합계 아래 여백 등)은 뺀다."""
    missing = [c for c in REQUIRED_FOR_ADDED if mapping.get(c) is None]
    if missing:
        raise ValueError(f"추가 신청 파일에서 다음 칸을 맞춰야 합니다: {missing}")
    data = sheets[sheet].iloc[header_row:].reset_index(drop=True)
    out = pd.DataFrame({c: (data[mapping[c]] if mapping.get(c) is not None else pd.Series([None] * len(data),
                                                                                          dtype=object))
                        for c in REQUIRED_COLUMNS})
    mapped_cols = [c for c in REQUIRED_COLUMNS if mapping.get(c) is not None]
    not_empty = out[mapped_cols].apply(lambda r: any(_cell_text(v) for v in r), axis=1)
    out = out[not_empty].reset_index(drop=True)

    notes = [[] for _ in range(len(out))]
    for col in DATE_COLUMNS:
        converted = []
        for i, value in enumerate(out[col]):
            date, ok = to_date(value)
            converted.append(date)
            if not ok:
                notes[i].append(f"{col}을 날짜로 읽지 못함(입력값: {_cell_text(value) or '빈칸'}) - 날짜 확인필요")
        out[col] = pd.Series(converted, dtype=object)
    out[INPUT_NOTE_COLUMN] = [" / ".join(n) for n in notes]
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
        note = []
        for col in DATE_COLUMNS:
            if record[col] is not None:
                date, ok = to_date(record[col])
                record[col] = date
                if not ok:
                    note.append(f"{col}을 날짜로 읽지 못함(입력값: {values.get(col)}) - 날짜 확인필요")
            else:
                note.append(f"{col}을 날짜로 읽지 못함(입력값: 빈칸) - 날짜 확인필요")
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
