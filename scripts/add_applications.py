#!/usr/bin/env python3
"""개별로 더 들어온 신청(추가 신청 파일)을 기존 결과에 반영한다.

사용법:
    python3 scripts/add_applications.py <원본신청명단.xlsx> <추가신청.xlsx> <환급과정_커리큘럼.xlsx> <비환급_가격표.xlsx> <기존_개강내역.xlsx> <출력폴더> [학점파일.xlsx]

- 추가 신청 파일은 원본 신청명단과 같은 칸이어야 한다.
- 1단계 신청명단: 원본 + 추가를 합쳐 새 이름(1_정리된_신청명단_추가_월일_시분.xlsx)으로 저장. 기존 파일은 그대로.
- 2단계 기업별 발송: 추가 신청이 있는 기업만, 그 기업의 원본 + 추가 신청을 담아
  새 폴더(2_기업별발송_추가_월일_시분)에 저장. 기존 파일은 그대로.
- 3단계 개강내역: 고른 기존 개강내역 파일을 갱신(덮어씀). 교육회차는 그대로 옮기고,
  인원이 바뀐 줄·새로 생긴 줄은 확인필요(노란색)에 표시.
- 학점 파일을 주면 도쿄일렉트론코리아 파일도 원본 + 추가로 새 이름으로 만든다.
- 추가 신청 중 원본에 같은 아이디 + 과정명이 이미 있으면 확인필요로 표시만 한다.
"""
import sys
import traceback
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_gaegang_sheet  # noqa: E402
import generate_company_files  # noqa: E402
import process_apply_list  # noqa: E402
import tel_list  # noqa: E402


def run(apply_path, added_path, refund_path, general_path, existing_gaegang_path, out_dir,
        credit_path=None, log=print) -> None:
    out_dir = Path(out_dir)
    stamp = datetime.now().strftime("%m%d_%H%M")

    log("=== 추가 신청 반영: 신청명단 ===")
    try:
        r1 = process_apply_list.run(apply_path, refund_path, general_path,
                                    out_dir / f"1_정리된_신청명단_추가_{stamp}.xlsx", added_path)
        log(f"완료: 추가 신청 {r1['added']}건 반영 (기존 명단과 중복 의심 {r1['duplicates']}건) / "
            f"전체 {r1['total']}건 / 확인필요 {r1['flagged']}건")
        log(f"결과 파일: {r1['output_path']}")
        log(f"제외된 신청명단 파일: {r1['excluded_output_path']}")
    except Exception:
        log("[오류] 신청명단 실패:\n" + traceback.format_exc())

    log("\n=== 추가 신청 반영: 기업별 발송 (추가된 기업만) ===")
    try:
        r2 = generate_company_files.run(apply_path, refund_path, general_path,
                                        out_dir / f"2_기업별발송_추가_{stamp}", added_path)
        for f in r2["files"]:
            log(f"생성: {f['path']} ({f['count']}명)")
        log(f"완료: 기업 {r2['company_count']}곳 / 확인필요 {r2['flagged']}건")
    except Exception:
        log("[오류] 기업별 발송 실패:\n" + traceback.format_exc())

    log("\n=== 추가 신청 반영: 개강내역 갱신 ===")
    try:
        r3 = build_gaegang_sheet.run(apply_path, refund_path, general_path, existing_gaegang_path,
                                     added_path, existing_gaegang_path)
        log(f"완료: 개강내역 {r3['rows']}줄 / 확인필요 {r3['flagged']}줄")
        log(f"갱신한 파일: {r3['output_path']}")
    except PermissionError:
        log("[오류] 개강내역 파일을 저장할 수 없습니다. 엑셀에서 열려 있다면 닫고 다시 실행해주세요.")
    except Exception:
        log("[오류] 개강내역 갱신 실패:\n" + traceback.format_exc())

    if credit_path:
        log("\n=== 추가 신청 반영: 도쿄일렉트론코리아 ===")
        try:
            r4 = tel_list.run(apply_path, general_path, credit_path,
                              out_dir / f"4_도쿄일렉트론코리아_추가_{stamp}.xlsx", added_path)
            log(f"완료: {r4['total']}건 / 확인필요 {r4['flagged']}건")
            log(f"결과 파일: {r4['output_path']}")
        except Exception:
            log("[오류] 도쿄일렉트론코리아 파일 생성 실패:\n" + traceback.format_exc())


def main():
    if len(sys.argv) not in (7, 8):
        print("사용법: python3 scripts/add_applications.py <원본신청명단.xlsx> <추가신청.xlsx> "
              "<환급과정_커리큘럼.xlsx> <비환급_가격표.xlsx> <기존_개강내역.xlsx> <출력폴더> [학점파일.xlsx]")
        sys.exit(1)
    run(*sys.argv[1:7], credit_path=sys.argv[7] if len(sys.argv) == 8 else None)


if __name__ == "__main__":
    main()
