#!/usr/bin/env python3
"""매달 받는 전체 신청명단을 기업별 발송용 파일로 나눠 생성한다.

사용법:
    python3 scripts/generate_company_files.py <신청명단.xlsx> <환급과정_커리큘럼.xlsx> <비환급_가격표.xlsx> <출력폴더>

규칙:
- 과정명이 환급과정 커리큘럼에 있으면 "환급"(과정유형: 사업주), 없고 비환급
  가격표에만 있으면 "비환급"(과정유형: 일반)으로 표기한다.
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

REQUIRED_APPLY_COLUMNS = [
    "순번", "신청ID", "구분", "출처", "회사명", "지사구분", "성함", "아이디",
    "생년월일", "휴대전화", "메일주소", "산인공코드", "신청과정명",
    "교육시작일", "교육종료일", "신청일시", "결제수단", "결제일", "결제금액", "결제상태",
]

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
        print(f"가격표 파일({path})에 다음 칸을 찾을 수 없습니다: {missing}")
        sys.exit(1)

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


def main():
    if len(sys.argv) != 5:
        print("사용법: python3 scripts/generate_company_files.py "
              "<신청명단.xlsx> <환급과정_커리큘럼.xlsx> <비환급_가격표.xlsx> <출력폴더>")
        sys.exit(1)

    apply_path, refund_curriculum_path, general_price_path, out_dir = sys.argv[1:5]
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(apply_path)
    missing = [c for c in REQUIRED_APPLY_COLUMNS if c not in df.columns]
    if missing:
        print(f"신청명단 파일에 다음 칸이 없습니다: {missing}")
        sys.exit(1)

    refund_table = load_price_table(refund_curriculum_path)
    general_table = load_price_table(general_price_path)

    rows = []
    for _, r in df.iterrows():
        course = r["신청과정명"]
        course_norm = normalize(course)
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

        rows.append({
            "기업명": clean_company_name(r["회사명"]),
            "성함": r["성함"],
            "아이디": r["아이디"],
            "생년월일": r["생년월일"],
            "휴대전화": r["휴대전화"],
            "과정명": course,
            "개강일": r["교육시작일"],
            "종료일": r["교육종료일"],
            "과정유형": course_type,
            "총수강료": tuition,
            "정부지원금": gov,
            "기업부담금": company_burden,
            "승인여부": "승인대기",
            "확인필요": note,
        })

    out_df = pd.DataFrame(rows)
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
        wb.save(file_path)

        created_files.append(file_path)
        print(f"생성: {file_path} ({len(group)}명)")

    print(f"\n완료: 총 {len(out_df)}건 / 기업 {len(created_files)}곳 / 확인필요 {total_flagged}건")
    if total_flagged:
        print("확인필요 표시가 있는 줄은 금액을 자동으로 채우지 않았으니 직접 확인해주세요.")


if __name__ == "__main__":
    main()
