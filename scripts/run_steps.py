"""'실행' 버튼이 하는 1~3단계(+도쿄일렉트론코리아) 처리를 차례로 돌리고 진행 상황을 log로 알린다.

exe 화면(webview_app.py)과 웹 화면(web/app.js, 브라우저 안의 Pyodide)이 같이 쓴다.
입력 검증·저장 폴더 선택은 각 화면이 맡고, 여기서는 이미 정해진 경로로 처리만 한다.
"""
import sys
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_gaegang_sheet  # noqa: E402
import generate_company_files  # noqa: E402
import process_apply_list  # noqa: E402
import tel_list  # noqa: E402


def run_all(apply_path, refund_path, general_path, credit_path, out_dir, save_tel,
            start_date=None, end_date=None, log=print) -> None:
    out_dir = Path(out_dir)

    if start_date or end_date:
        log(f"기간 필터: 교육시작일 {start_date or '처음'} ~ {end_date or '끝'} 만 처리합니다.")
    log("=== 1단계: 신청명단 정리 ===")
    try:
        r1 = process_apply_list.run(apply_path, refund_path, general_path, out_dir / "1_정리된_신청명단.xlsx",
                                    start_date=start_date, end_date=end_date)
        date_note = f", 선택한 기간 밖 {r1['date_filtered']}건 제외" if r1["date_filtered"] else ""
        log(f"완료: 총 {r1['total']}건 (프리패스/테스트/카드결제 {r1['excluded']}건, "
            f"도쿄일렉트론코리아 {r1['tel_excluded']}건 제외{date_note}) / "
            f"환급 {r1['refund']}건 / 비환급 {r1['nonrefund']}건 / 확인필요 {r1['flagged']}건")
        log(f"결과 파일: {r1['output_path']}")
        log(f"제외된 신청명단 파일: {r1['excluded_output_path']}")
    except Exception:
        log("[오류] 1단계 실패:\n" + traceback.format_exc())

    log("\n=== 2단계: 기업별 발송 파일 생성 ===")
    try:
        r2 = generate_company_files.run(apply_path, refund_path, general_path, out_dir / "2_기업별발송",
                                         start_date=start_date, end_date=end_date)
        for f in r2["files"]:
            log(f"생성: {f['path']} ({f['count']}명)")
        date_note = f", 선택한 기간 밖 {r2['date_filtered']}건 제외" if r2["date_filtered"] else ""
        log(f"완료: 총 {r2['total']}건 (프리패스/테스트/카드결제 {r2['excluded']}건, "
            f"도쿄일렉트론코리아 {r2['tel_excluded']}건 제외{date_note}) / "
            f"기업 {r2['company_count']}곳 / 확인필요 {r2['flagged']}건")
    except Exception:
        log("[오류] 2단계 실패:\n" + traceback.format_exc())

    log("\n=== 3단계: 개강내역시트 작성 내용 생성 ===")
    try:
        r3 = build_gaegang_sheet.run(apply_path, refund_path, general_path, out_dir / "3_개강내역.xlsx",
                                     start_date=start_date, end_date=end_date)
        date_note = f", 선택한 기간 밖 {r3['date_filtered']}건 제외" if r3["date_filtered"] else ""
        log(f"완료: 개강내역 {r3['rows']}줄 (신청 {r3['students']}건을 묶음, "
            f"프리패스/테스트/카드결제 {r3['excluded']}건, "
            f"도쿄일렉트론코리아 {r3['tel_excluded']}건 제외{date_note}) / 확인필요 {r3['flagged']}줄")
        log(f"결과 파일: {r3['output_path']}")
    except Exception:
        log("[오류] 3단계 실패:\n" + traceback.format_exc())

    if save_tel:
        log("\n=== 도쿄일렉트론코리아 별도 파일 ===")
        try:
            r4 = tel_list.run(apply_path, general_path, credit_path, out_dir / "4_도쿄일렉트론코리아.xlsx",
                              start_date=start_date, end_date=end_date)
            log(f"완료: {r4['total']}건 / 확인필요 {r4['flagged']}건")
            log(f"결과 파일: {r4['output_path']}")
        except Exception:
            log("[오류] 도쿄일렉트론코리아 파일 생성 실패:\n" + traceback.format_exc())
