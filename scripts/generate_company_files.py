#!/usr/bin/env python3
"""매달 받는 전체 신청명단을 기업별 발송용 파일로 나눠 생성한다.

사용법:
    python3 scripts/generate_company_files.py <신청명단.xlsx> <환급과정_커리큘럼.xlsx> <비환급과정_커리큘럼.xlsx> <출력폴더>

규칙:
- 과정명이 환급과정 커리큘럼에 있으면 "환급"(과정유형: 사업주), 없고
  비환급과정 커리큘럼에만 있으면 "비환급"(과정유형: 일반)으로 표기한다.
- 양쪽 리스트에 모두 있는 과정은 "환급"으로 표기하고, 금액도 환급과정 쪽
  금액을 쓴다. 이 경우 "확인필요" 칸에 전환 사실을 남긴다.
- 비환급 과정은 정부지원금이 없으므로 기업부담금 = 수강료 전액으로 계산한다.
- 어느 리스트에서도 과정명을 찾지 못하면 금액을 비워두고 "확인필요"로 표시한다
  (추측해서 채우지 않는다).
- 성함/아이디 등 개인정보는 마스킹하지 않는다. 과정명은 원본 그대로 둔다.

파일 모양은 지금 기업에 보내는 서식(templates/company_file_template.xlsx, 받은 서식에서 수강생 줄만 지운 것)을
그대로 쓴다: 1~16행(제목·총 수강료/총 정부지원금·안내문·표 머리글)은 서식 그대로, 17행부터 수강생 한 명씩
17행의 줄 서식을 복사해 채운다. 제목(B2)은 "2026년 11월 2일 (월) 개강 신청 내역", 파일 이름은
"(엘캠퍼스) 2026년 11월 02일_개강명단_기업명.xlsx"로 개강일을 넣는다(한 기업에 개강일이 여러 개면 가장 빠른
날짜를 쓰고 로그에 확인필요로 알린다). K6·L6은 =SUM 수식. 비환급(일반)의 정부지원금은 "-"로 쓴다.
확인필요 메모는 발송 파일에 넣지 않는다(1번 정리된 신청명단·처리 로그에서 확인).
"""
import re
import sys
from copy import copy
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pricing import (  # noqa: E402
    added_note, classify_course, clean_company_name, filter_by_date_range, filter_excluded_rows, load_price_table,
    read_apply_list, split_tel_rows,
)

REQUIRED_APPLY_COLUMNS = [
    "순번", "신청ID", "구분", "출처", "회사명", "지사구분", "성함", "아이디",
    "생년월일", "휴대전화", "메일주소", "산인공코드", "신청과정명",
    "교육시작일", "교육종료일", "신청일시", "결제수단", "결제일", "결제금액", "결제상태", "신청상태",
]


# 서식 파일 위치. exe(pyinstaller)로 묶었을 때는 --add-data "templates;templates"로 함께 묶는다.
TEMPLATE_PATH = (Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
                 / "templates" / "company_file_template.xlsx")
FIRST_DATA_ROW = 17
# 서식 파일 표의 칸(B~M) 순서
SHEET_COLUMNS = ["성함", "아이디", "생년월일", "휴대전화", "과정명", "개강일", "종료일", "과정유형",
                 "총수강료", "정부지원금", "기업부담금", "승인여부"]
FIRST_COLUMN = 2  # B열
WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]


def _number(value):
    """금액: 비어 있으면 None, 소수점 없는 값은 정수로."""
    if value is None or pd.isna(value):
        return None
    return int(value) if float(value).is_integer() else value


def _cell_value(value):
    return None if value is None or (not isinstance(value, str) and pd.isna(value)) else value


def _start_date(group: pd.DataFrame):
    """제목·파일 이름에 쓸 개강일(가장 빠른 날짜)과, 개강일이 여러 개인지/읽지 못했는지 알리는 메모."""
    dates = pd.to_datetime(group["개강일"], errors="coerce").dropna().dt.normalize().unique()
    if len(dates) == 0:
        return None, "개강일을 날짜로 읽지 못해 제목·파일 이름에 날짜를 넣지 못함 - 확인필요"
    first = min(dates)
    note = ""
    if len(dates) > 1:
        shown = ", ".join(pd.Timestamp(d).strftime("%Y-%m-%d") for d in sorted(dates))
        note = f"개강일이 여러 개({shown}) - 제목·파일 이름은 가장 빠른 날짜로 넣음, 확인필요"
    return pd.Timestamp(first), note


def write_company_file(group: pd.DataFrame, file_path: Path, start) -> None:
    """서식 파일에 한 기업의 수강생들을 채워 file_path로 저장한다."""
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    if start is not None:
        ws["B2"] = f"{start.year}년 {start.month}월 {start.day}일 ({WEEKDAYS[start.weekday()]}) 개강 신청 내역"
    else:
        ws["B2"] = "개강 신청 내역"

    style_row = [ws.cell(FIRST_DATA_ROW, FIRST_COLUMN + i) for i in range(len(SHEET_COLUMNS))]
    styles = [(copy(c.font), copy(c.border), copy(c.fill), c.number_format, copy(c.alignment), copy(c.protection))
              for c in style_row]
    height = ws.row_dimensions[FIRST_DATA_ROW].height
    last_row = FIRST_DATA_ROW + len(group) - 1
    for row_idx, (_, r) in enumerate(group.iterrows(), start=FIRST_DATA_ROW):
        ws.row_dimensions[row_idx].height = height
        for i, col in enumerate(SHEET_COLUMNS):
            cell = ws.cell(row_idx, FIRST_COLUMN + i)
            cell.font, cell.border, cell.fill, cell.number_format, cell.alignment, cell.protection = (
                copy(styles[i][0]), copy(styles[i][1]), copy(styles[i][2]), styles[i][3], copy(styles[i][4]),
                copy(styles[i][5]))
            cell.value = _cell_value(r[col])

    ws["K6"] = f"=SUM(J{FIRST_DATA_ROW}:J{last_row})"
    ws["L6"] = f"=SUM(K{FIRST_DATA_ROW}:K{last_row})"
    ws.auto_filter.ref = f"B16:M{last_row}"
    # 과정유형 칸 조건부 서식(일반/사업주/담당자개별신청 색)을 채운 줄까지 넓힌다.
    old_rules = ws.conditional_formatting
    ws.conditional_formatting = ConditionalFormattingList()
    for rng in old_rules:
        for rule in rng.rules:
            ws.conditional_formatting.add(f"I{FIRST_DATA_ROW}:I{last_row}", rule)
    wb.save(file_path)


def run(apply_path: str, refund_curriculum_path: str, general_price_path: str, out_dir: str,
        added_path: str = None, start_date: str = None, end_date: str = None) -> dict:
    """기업별 발송 파일들을 out_dir에 생성하고 요약 결과를 돌려준다.
    added_path(추가 신청 파일)를 주면 원본 + 추가 신청을 합친 뒤, 추가 신청이 있는
    기업의 파일만 만든다(그 기업의 기존 + 추가 신청이 모두 들어감).
    start_date/end_date를 주면 교육시작일이 그 기간 안에 있는 신청 건만 처리한다
    (둘 다 비어 있으면 전체 기간을 처리).
    입력 파일에 필요한 칸이 없으면 ValueError를 발생시킨다(추측해서 넘어가지 않음)."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    df = read_apply_list(apply_path, REQUIRED_APPLY_COLUMNS, added_path)

    df, date_filtered_count = filter_by_date_range(df, start_date, end_date)
    df, _, excluded_count = filter_excluded_rows(df)
    df, tel_df = split_tel_rows(df)

    refund_table = load_price_table(refund_curriculum_path)
    general_table = load_price_table(general_price_path)

    rows = []
    for _, r in df.iterrows():
        course = r["신청과정명"]
        result = classify_course(course, refund_table, general_table, r["구분"])

        rows.append({
            "기업명": clean_company_name(r["회사명"]),
            "성함": r["성함"],
            "아이디": r["아이디"],
            "생년월일": r["생년월일"],
            "휴대전화": r["휴대전화"],
            "과정명": course,
            "개강일": r["교육시작일"],
            "종료일": r["교육종료일"],
            "과정유형": result["과정유형"],
            "총수강료": _number(result["수강료"]),
            # 비환급(일반)은 정부지원금이 없어서 서식대로 "-"로 쓴다.
            "정부지원금": "-" if result["환급구분"] == "비환급" else _number(result["정부지원금"]),
            "기업부담금": _number(result["기업부담금"]),
            "승인여부": "승인대기",
            "확인필요": " / ".join(n for n in (result["확인필요"], r["_중복메모"], r["_입력메모"],
                                                 added_note(r["_추가"])) if n),
            "_추가": r["_추가"],
        })

    out_df = pd.DataFrame(rows)
    if added_path:
        added_companies = set(out_df.loc[out_df["_추가"], "기업명"])
        out_df = out_df[out_df["기업명"].isin(added_companies)]
    out_df = out_df.drop(columns=["_추가"])
    total_flagged = int((out_df["확인필요"] != "").sum())

    created_files = []
    for company, group in out_df.groupby("기업명", sort=False):
        safe_name = re.sub(r'[\\/:*?"<>|]', "_", str(company)) or "기업명없음"
        start, note = _start_date(group)
        date_part = f"{start.year}년 {start.month:02d}월 {start.day:02d}일" if start is not None else "개강일 확인필요"
        file_path = out_dir / f"(엘캠퍼스) {date_part}_개강명단_{safe_name}.xlsx"
        write_company_file(group, file_path, start)
        created_files.append({"path": file_path, "count": len(group), "note": note,
                              "flagged": int((group["확인필요"] != "").sum())})

    return {
        "total": len(out_df),
        "company_count": len(created_files),
        "flagged": total_flagged,
        "excluded": excluded_count,
        "tel_excluded": len(tel_df),
        "date_filtered": date_filtered_count,
        "files": created_files,
    }


def main():
    if len(sys.argv) != 5:
        print("사용법: python3 scripts/generate_company_files.py "
              "<신청명단.xlsx> <환급과정_커리큘럼.xlsx> <비환급과정_커리큘럼.xlsx> <출력폴더>")
        sys.exit(1)

    apply_path, refund_curriculum_path, general_price_path, out_dir = sys.argv[1:5]
    try:
        result = run(apply_path, refund_curriculum_path, general_price_path, out_dir)
    except ValueError as e:
        print(str(e))
        sys.exit(1)

    for f in result["files"]:
        print(f"생성: {f['path']} ({f['count']}명)" + (f" - {f['note']}" if f["note"] else ""))

    date_note = f", 선택한 기간 밖 {result['date_filtered']}건 제외" if result["date_filtered"] else ""
    print(f"\n완료: 총 {result['total']}건 (프리패스/테스트/카드결제 {result['excluded']}건, "
          f"도쿄일렉트론코리아 {result['tel_excluded']}건 제외{date_note}) / "
          f"기업 {result['company_count']}곳 / 확인필요 {result['flagged']}건")
    if result["flagged"]:
        print("확인필요 건은 발송 파일에는 표시되지 않습니다. 1번 정리된 신청명단에서 확인해주세요.")


if __name__ == "__main__":
    main()
