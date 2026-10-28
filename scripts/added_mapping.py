"""서식이 제각각인 추가 신청 파일(또는 붙여넣은 표)을 원본 신청명단과 같은 칸(표준 양식)으로 바꾼다.

웹 화면의 "칸 맞추기"에서 쓴다. 사용자가 시트와 제목 줄(칸 이름이 있는 줄)을 고르고,
원본 신청명단의 칸마다 추가 신청 파일의 어느 칸을 쓸지 직접 정한다.
- 자동 추천은 칸 이름이 (띄어쓰기만 빼고) 똑같을 때만 한다. 비슷한 이름을 추측해서 맞추지 않는다.
- 필수 칸(REQUIRED_FOR_ADDED)은 반드시 맞춰야 하고, 나머지 칸은 맞출 곳이 없으면 비워 둔다
  (값을 지어내서 채우지 않음).
- 값은 원본 그대로 옮기고 고치지 않는다(신청과정명 포함). 단, 교육시작일·교육종료일은 개강내역에서
  원본 신청명단(날짜 서식 셀)과 같은 날끼리 묶이도록 날짜로 읽어 맞춘다. 날짜로 읽지 못한 값은
  그대로 두고 확인필요 메모(pricing.INPUT_NOTE_COLUMN)를 남긴다. 연도가 없는 날짜는 추측하지 않는다.
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
from pricing import INPUT_NOTE_COLUMN  # noqa: E402
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
