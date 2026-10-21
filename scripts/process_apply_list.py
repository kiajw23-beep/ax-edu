#!/usr/bin/env python3
"""매달 받는 전체 신청명단 raw data를 정리된 양식으로 변환한다.

사용법:
    python3 scripts/process_apply_list.py <신청명단.xlsx> <환급과정_커리큘럼.xlsx> <출력파일.xlsx>
"""
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pricing import clean_company_name, filter_excluded_rows  # noqa: E402

REQUIRED_COLUMNS = [
    "순번", "신청ID", "구분", "출처", "회사명", "지사구분", "성함", "아이디",
    "생년월일", "휴대전화", "메일주소", "산인공코드", "신청과정명",
    "교육시작일", "교육종료일", "신청일시", "결제수단", "결제일", "결제금액", "결제상태",
]

# 정리된 신청명단에는 필요 없어서 결과 파일에서 빼는 칸
DROP_FROM_OUTPUT = ["신청ID", "결제수단", "결제일", "결제금액", "결제상태"]

WHITESPACE_CHECK_COLUMNS = ["회사명", "성함", "아이디", "휴대전화", "메일주소"]


def has_whitespace_issue(value) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    text = str(value)
    if text != text.strip():
        return True
    if "  " in text:
        return True
    return False


def load_refund_course_names(path: str) -> set[str]:
    is_csv = path.lower().endswith(".csv")

    # 1) 헤더에 "과정명"이 들어간 칸이 있으면 그 칸만 사용한다.
    #    (과정 상세정보가 함께 들어있는 표라도 과정명만 정확히 뽑아내기 위함)
    df_with_header = pd.read_csv(path) if is_csv else pd.read_excel(path)
    course_column = next(
        (col for col in df_with_header.columns if isinstance(col, str) and "과정명" in col),
        None,
    )
    if course_column is not None:
        values = df_with_header[course_column]
    else:
        # 2) 헤더가 없는 단순 과정명 목록으로 보고, 모든 칸을 과정명 후보로 취급한다.
        df_flat = pd.read_csv(path, header=None) if is_csv else pd.read_excel(path, header=None)
        values = df_flat.values.flatten()

    names = set()
    for value in values:
        if pd.isna(value):
            continue
        text = re.sub(r"\s+", " ", str(value)).strip()
        if text:
            names.add(text)
    return names


def run(apply_list_path: str, curriculum_path: str, output_path: str) -> dict:
    """신청명단을 정리해 output_path에 저장하고 요약 결과를 돌려준다.
    입력 파일에 필요한 칸이 없으면 ValueError를 발생시킨다(추측해서 넘어가지 않음)."""
    df = pd.read_excel(apply_list_path)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"신청명단 파일에 다음 칸이 없습니다. 원본 파일 형식을 확인해주세요: {missing}")

    df, excluded_count = filter_excluded_rows(df)

    refund_courses = load_refund_course_names(curriculum_path)

    df["기업명"] = df["회사명"].apply(clean_company_name)

    def classify(course_name) -> str:
        normalized = re.sub(r"\s+", " ", str(course_name)).strip()
        return "환급" if normalized in refund_courses else "비환급"

    df["환급구분"] = df["신청과정명"].apply(classify)

    def whitespace_note(row) -> str:
        issues = [col for col in WHITESPACE_CHECK_COLUMNS if has_whitespace_issue(row.get(col))]
        return (", ".join(issues) + " 공백 확인필요") if issues else ""

    df["확인필요"] = df.apply(whitespace_note, axis=1)

    ordered_columns = [c for c in REQUIRED_COLUMNS if c not in DROP_FROM_OUTPUT] + ["기업명", "환급구분", "확인필요"]
    df = df[ordered_columns]

    df.to_excel(output_path, index=False, sheet_name="정리된_신청명단")

    wb = load_workbook(output_path)
    ws = wb["정리된_신청명단"]
    fill = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")
    check_col_idx = ordered_columns.index("확인필요") + 1
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=check_col_idx)
        if cell.value:
            cell.fill = fill
    wb.save(output_path)

    total = len(df)
    refund_count = int((df["환급구분"] == "환급").sum())
    nonrefund_count = total - refund_count
    flagged = int((df["확인필요"] != "").sum())

    return {
        "total": total,
        "refund": refund_count,
        "nonrefund": nonrefund_count,
        "flagged": flagged,
        "excluded": excluded_count,
        "output_path": output_path,
    }


def main():
    if len(sys.argv) != 4:
        print("사용법: python3 scripts/process_apply_list.py <신청명단.xlsx> <환급과정_커리큘럼.xlsx> <출력파일.xlsx>")
        sys.exit(1)

    apply_list_path, curriculum_path, output_path = sys.argv[1:4]
    try:
        result = run(apply_list_path, curriculum_path, output_path)
    except ValueError as e:
        print(str(e))
        sys.exit(1)

    print(f"완료: 총 {result['total']}건 (프리패스/테스트 {result['excluded']}건 제외) / "
          f"환급 {result['refund']}건 / 비환급 {result['nonrefund']}건 / 확인필요 {result['flagged']}건")
    print(f"결과 파일: {result['output_path']}")


if __name__ == "__main__":
    main()
