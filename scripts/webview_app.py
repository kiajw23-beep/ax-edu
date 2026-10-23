#!/usr/bin/env python3
"""월별 개강 업무 도구 - 화면(GUI) 버전, pywebview.

로직(파일 읽기/쓰기, 환급 판정, 집계 등)은 전부 process_apply_list.py 등
기존 스크립트 그대로 쓴다. 이 파일은 화면(로컬 HTML/CSS/JS, webview_ui/)과
파일 대화상자·실행 흐름만 담당한다.

신청명단 / 환급과정 커리큘럼 / 비환급 가격표, 3개 파일을 고르고 "실행"을
누르면 저장 폴더를 고른 뒤 아래 결과를 한 번에 만든다.
    1_정리된_신청명단.xlsx
    2_기업별발송/ (기업별 파일들)
    3_개강내역.xlsx
    4_도쿄일렉트론코리아.xlsx (체크박스 선택 + 학점 파일을 고른 경우에만)
    5_제외된_신청명단.xlsx (프리패스/테스트/카드결제로 결과에서 빠진 신청 건만 모은 파일, 항상 생성)
"추가 신청 반영" 버튼은 ⑤ 추가 신청 파일을 골라 개별로 더 들어온 신청을 반영한다
(자세한 규칙은 add_applications.py 참고). ⑥ 기존 개강내역 파일은 선택사항이며,
고르면 그 파일을 갱신(교육회차 유지 + 인원 변경 확인필요)하고, 고르지 않으면
3_개강내역_추가_월일_시분.xlsx로 새로 만든다(교육회차는 비워둠).
도쿄일렉트론코리아 신청 건은 체크 여부와 관계없이 1~3단계 결과에서 항상 빠진다.

각 파일 선택 줄의 "최근" 버튼을 누르면 그 칸에서 최근에 고른 파일(최대 3개,
칸마다 따로 기억)이 목록으로 나오고, 클릭하면 그 경로가 그대로 입력된다.
이 이력은 %APPDATA%\\월별개강업무도구\\최근파일이력.json 에 저장되어 다시 열어도
유지된다(파일 경로만 저장, 파일 내용은 저장하지 않음).

개인정보가 들어있는 파일을 다루므로 이 프로그램은 인터넷에 아무것도
전송하지 않는다(전부 이 컴퓨터 안에서만 처리).

exe로 만드는 법(윈도우 PC에서):
    pip install -r requirements.txt pyinstaller
    pyinstaller --onefile --windowed --name 월별개강업무도구 ^
        --add-data "scripts/webview_ui;webview_ui" scripts/webview_app.py
결과물은 dist/월별개강업무도구.exe 에 생성된다.
"""
import json
import os
import sys
import traceback
from pathlib import Path

import webview

sys.path.insert(0, str(Path(__file__).resolve().parent))
import process_apply_list  # noqa: E402
import generate_company_files  # noqa: E402
import add_applications  # noqa: E402
import build_gaegang_sheet  # noqa: E402
import tel_list  # noqa: E402

# 각 파일 선택 줄에서 최근 고른 파일 경로를 기억해두는 곳(파일칸 종류별로 따로 저장).
RECENT_FILES_PATH = Path(os.environ.get("APPDATA", str(Path.home()))) / "월별개강업무도구" / "최근파일이력.json"
MAX_RECENT_FILES = 3

FILE_TYPES = ("엑셀 CSV 파일 (*.xlsx;*.xls;*.csv)", "모든 파일 (*.*)")


def _resource_path(*parts) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base.joinpath(*parts)


class Api:
    def __init__(self):
        self.window = None
        self.recent_files = self._load_recent_files()

    # ----- 최근 파일 이력 -----
    def _load_recent_files(self) -> dict:
        try:
            with open(RECENT_FILES_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _save_recent_files(self):
        try:
            RECENT_FILES_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(RECENT_FILES_PATH, "w", encoding="utf-8") as f:
                json.dump(self.recent_files, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _remember_recent_file(self, slot_key: str, path: str):
        paths = [p for p in self.recent_files.get(slot_key, []) if p != path]
        paths.insert(0, path)
        self.recent_files[slot_key] = paths[:MAX_RECENT_FILES]
        self._save_recent_files()

    def get_recent_files(self, slot_key: str) -> list:
        return self.recent_files.get(slot_key, [])

    def use_recent_file(self, slot_key: str, path: str) -> str:
        self._remember_recent_file(slot_key, path)
        return path

    # ----- 파일/폴더 대화상자 -----
    def pick_file(self, slot_key: str):
        result = self.window.create_file_dialog(webview.OPEN_DIALOG, file_types=FILE_TYPES)
        if not result:
            return None
        path = result[0]
        self._remember_recent_file(slot_key, path)
        return path

    def _pick_directory(self):
        result = self.window.create_file_dialog(webview.FOLDER_DIALOG)
        return result[0] if result else None

    # ----- 화면에 로그 찍기 -----
    def _log(self, text: str):
        self.window.evaluate_js(f"window.appendLog({json.dumps(text)})")

    def _clear_log(self):
        self.window.evaluate_js("window.clearLog()")

    def _clear_fields(self):
        self.window.evaluate_js("window.clearFileFields()")

    # ----- 실행 -----
    def run_all(self, paths: dict, save_tel: bool) -> dict:
        apply_path = paths.get("apply", "")
        refund_path = paths.get("refund", "")
        general_path = paths.get("general", "")
        credit_path = paths.get("credit", "")

        if not (apply_path and refund_path and general_path):
            return {"ok": False, "message": "신청명단 / 환급과정 커리큘럼 / 비환급 가격표, 3개 파일을 모두 골라야 합니다."}
        if save_tel and not credit_path:
            return {"ok": False, "message": "도쿄일렉트론코리아 별도 파일을 만들려면 ④ 학점 파일을 골라야 합니다."}

        out_dir_str = self._pick_directory()
        if not out_dir_str:
            return {"ok": False, "message": None}
        out_dir = Path(out_dir_str)

        self._clear_log()
        self._log("=== 1단계: 신청명단 정리 ===")
        try:
            r1 = process_apply_list.run(apply_path, refund_path, general_path, out_dir / "1_정리된_신청명단.xlsx")
            self._log(f"완료: 총 {r1['total']}건 (프리패스/테스트/카드결제 {r1['excluded']}건, "
                      f"도쿄일렉트론코리아 {r1['tel_excluded']}건 제외) / "
                      f"환급 {r1['refund']}건 / 비환급 {r1['nonrefund']}건 / 확인필요 {r1['flagged']}건")
            self._log(f"결과 파일: {r1['output_path']}")
            self._log(f"제외된 신청명단 파일: {r1['excluded_output_path']}")
        except Exception:
            self._log("[오류] 1단계 실패:\n" + traceback.format_exc())

        self._log("\n=== 2단계: 기업별 발송 파일 생성 ===")
        try:
            r2 = generate_company_files.run(apply_path, refund_path, general_path, out_dir / "2_기업별발송")
            for f in r2["files"]:
                self._log(f"생성: {f['path']} ({f['count']}명)")
            self._log(f"완료: 총 {r2['total']}건 (프리패스/테스트/카드결제 {r2['excluded']}건, "
                      f"도쿄일렉트론코리아 {r2['tel_excluded']}건 제외) / "
                      f"기업 {r2['company_count']}곳 / 확인필요 {r2['flagged']}건")
        except Exception:
            self._log("[오류] 2단계 실패:\n" + traceback.format_exc())

        self._log("\n=== 3단계: 개강내역시트 작성 내용 생성 ===")
        try:
            r3 = build_gaegang_sheet.run(apply_path, refund_path, general_path, out_dir / "3_개강내역.xlsx")
            self._log(f"완료: 개강내역 {r3['rows']}줄 (신청 {r3['students']}건을 묶음, "
                      f"프리패스/테스트/카드결제 {r3['excluded']}건, "
                      f"도쿄일렉트론코리아 {r3['tel_excluded']}건 제외) / 확인필요 {r3['flagged']}줄")
            self._log(f"결과 파일: {r3['output_path']}")
        except Exception:
            self._log("[오류] 3단계 실패:\n" + traceback.format_exc())

        if save_tel:
            self._log("\n=== 도쿄일렉트론코리아 별도 파일 ===")
            try:
                r4 = tel_list.run(apply_path, general_path, credit_path, out_dir / "4_도쿄일렉트론코리아.xlsx")
                self._log(f"완료: {r4['total']}건 / 확인필요 {r4['flagged']}건")
                self._log(f"결과 파일: {r4['output_path']}")
            except Exception:
                self._log("[오류] 도쿄일렉트론코리아 파일 생성 실패:\n" + traceback.format_exc())

        self._log(f"\n모든 작업이 끝났습니다. 저장 위치: {out_dir}")
        self._clear_fields()

        if sys.platform == "win32":
            os.startfile(out_dir)  # noqa: S606

        return {"ok": True, "message": f"처리가 끝났습니다.\n저장 위치: {out_dir}\n\n"
                                        "노란색으로 표시된 '확인필요' 줄은 직접 확인해주세요."}

    def run_added(self, paths: dict, save_tel: bool) -> dict:
        apply_path = paths.get("apply", "")
        refund_path = paths.get("refund", "")
        general_path = paths.get("general", "")
        credit_path = paths.get("credit", "")
        added_path = paths.get("added", "")
        existing_gaegang_path = paths.get("existing_gaegang", "")

        if not (apply_path and refund_path and general_path and added_path):
            return {"ok": False, "message": "① 신청명단(원본) / ② 환급과정 커리큘럼 / ③ 비환급 가격표 / "
                                             "⑤ 추가 신청 파일을 모두 골라야 합니다. "
                                             "⑥ 기존 개강내역 파일은 선택사항입니다."}
        if save_tel and not credit_path:
            return {"ok": False, "message": "도쿄일렉트론코리아 별도 파일을 만들려면 ④ 학점 파일을 골라야 합니다."}

        out_dir_str = self._pick_directory()
        if not out_dir_str:
            return {"ok": False, "message": None}
        out_dir = Path(out_dir_str)

        self._clear_log()
        add_applications.run(apply_path, added_path, refund_path, general_path, existing_gaegang_path, out_dir,
                             credit_path=credit_path if save_tel else None, log=self._log)

        gaegang_msg = ("개강내역은 고른 파일을 갱신했습니다." if existing_gaegang_path
                       else "⑥을 고르지 않아 개강내역은 새 파일로 만들었고, 교육회차는 비어 있습니다.")
        self._log(f"\n추가 신청 반영이 끝났습니다. 저장 위치: {out_dir}")
        self._clear_fields()

        if sys.platform == "win32":
            os.startfile(out_dir)  # noqa: S606

        return {"ok": True, "message": f"추가 신청 반영이 끝났습니다.\n저장 위치: {out_dir}\n\n"
                                        f"{gaegang_msg} 노란색 '확인필요' 줄은 직접 확인해주세요."}


def main():
    api = Api()
    window = webview.create_window(
        "월별 개강 업무 도구",
        url=str(_resource_path("webview_ui", "index.html")),
        js_api=api,
        width=880,
        height=780,
        min_size=(760, 600),
    )
    api.window = window
    webview.start()


if __name__ == "__main__":
    main()
