"""서식이 제각각인 추가 신청 파일을 원본 신청명단과 같은 칸(표준 양식)으로 바꾼다.

웹 화면의 "칸 맞추기"에서 쓴다. 사용자가 시트와 제목 줄(칸 이름이 있는 줄)을 고르고,
원본 신청명단의 칸마다 추가 신청 파일의 어느 칸을 쓸지 직접 정한다.
- 자동 추천은 칸 이름이 (띄어쓰기만 빼고) 똑같을 때만 한다. 비슷한 이름을 추측해서 맞추지 않는다.
- 필수 칸(REQUIRED_FOR_ADDED)은 반드시 맞춰야 하고, 나머지 칸은 맞출 곳이 없으면 비워 둔다
  (값을 지어내서 채우지 않음).
- 값은 원본 그대로 옮기고 고치지 않는다(신청과정명 포함).
만든 표준 양식 파일은 기존 add_applications.run()에 그대로 넘긴다.
"""
import re
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from process_apply_list import REQUIRED_COLUMNS  # noqa: E402

# 이 칸들이 없으면 기업명·환급 판정·개강내역 묶음을 만들 수 없어서 꼭 맞춰야 한다.
REQUIRED_FOR_ADDED = ["회사명", "성함", "아이디", "신청과정명", "교육시작일", "교육종료일"]


def read_sheets(path: str) -> dict:
    """시트 이름 -> 칸 이름 없이 읽은 표. CSV는 시트 하나("CSV")로 본다."""
    if path.lower().endswith(".csv"):
        try:
            return {"CSV": pd.read_csv(path, header=None, dtype=object, encoding="utf-8-sig")}
        except UnicodeDecodeError:
            # 한국어 윈도우 엑셀에서 저장한 CSV
            return {"CSV": pd.read_csv(path, header=None, dtype=object, encoding="cp949")}
    return pd.read_excel(path, sheet_name=None, header=None)


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
    return out[not_empty].reset_index(drop=True)


def save_standard(path: str, sheet: str, header_row: int, mapping: dict, output_path: str, log=print) -> str:
    """추가 신청 파일을 표준 양식으로 바꿔 output_path에 저장하고 그 경로를 돌려준다."""
    sheets = read_sheets(path)
    df = build_standard(sheets, sheet, header_row, mapping)
    df.to_excel(output_path, index=False)
    blank = [c for c in REQUIRED_COLUMNS if mapping.get(c) is None]
    log(f"칸 맞추기: 추가 신청 {len(df)}건 (시트 '{sheet}', 제목 줄 {header_row}번째 줄)")
    if blank:
        log(f"맞출 칸이 없어 비워 둔 칸: {', '.join(blank)}")
    return output_path
