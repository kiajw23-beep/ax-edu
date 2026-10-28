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
"""
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pricing import (  # noqa: E402
    added_note, classify_course, clean_company_name, filter_by_date_range, filter_excluded_rows, load_price_table,
    read_apply_list, split_tel_rows, style_worksheet,
)

REQUIRED_APPLY_COLUMNS = [
    "순번", "신청ID", "구분", "출처", "회사명", "지사구분", "성함", "아이디",
    "생년월일", "휴대전화", "메일주소", "산인공코드", "신청과정명",
    "교육시작일", "교육종료일", "신청일시", "결제수단", "결제일", "결제금액", "결제상태", "신청상태",
]


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
            "총수강료": result["수강료"],
            "정부지원금": result["정부지원금"],
            "기업부담금": result["기업부담금"],
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

    fill = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")
    created_files = []
    for company, group in out_df.groupby("기업명", sort=False):
        safe_name = re.sub(r'[\\/:*?"<>|]', "_", str(company)) or "기업명없음"
        file_path = out_dir / f"{safe_name}_신청명단.xlsx"
        group = group.drop(columns=["기업명"])
        group.to_excel(file_path, index=False, sheet_name="신청명단")

        wb = load_workbook(file_path)
        ws = wb["신청명단"]
        check_col_idx = list(group.columns).index("확인필요") + 1
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=check_col_idx)
            if cell.value:
                cell.fill = fill
        style_worksheet(ws, currency_columns=["총수강료", "정부지원금", "기업부담금"], text_columns=["아이디"])
        wb.save(file_path)

        created_files.append({"path": file_path, "count": len(group)})

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
        print(f"생성: {f['path']} ({f['count']}명)")

    date_note = f", 선택한 기간 밖 {result['date_filtered']}건 제외" if result["date_filtered"] else ""
    print(f"\n완료: 총 {result['total']}건 (프리패스/테스트/카드결제 {result['excluded']}건, "
          f"도쿄일렉트론코리아 {result['tel_excluded']}건 제외{date_note}) / "
          f"기업 {result['company_count']}곳 / 확인필요 {result['flagged']}건")
    if result["flagged"]:
        print("확인필요 표시가 있는 줄은 금액을 자동으로 채우지 않았으니 직접 확인해주세요.")


if __name__ == "__main__":
    main()
