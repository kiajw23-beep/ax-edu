#!/usr/bin/env python3
"""월별 개강 업무 도구 - 화면(GUI) 버전.

신청명단 / 환급과정 커리큘럼 / 비환급 가격표, 3개 파일을 고르고 "실행"을
누르면 저장 폴더를 고른 뒤 아래 3개 결과를 한 번에 만든다.
    1_정리된_신청명단.xlsx
    2_기업별발송/ (기업별 파일들)
    3_개강내역.xlsx

개인정보가 들어있는 파일을 다루므로 이 프로그램은 인터넷에 아무것도
전송하지 않는다(전부 이 컴퓨터 안에서만 처리).

exe로 만드는 법(윈도우 PC에서):
    pip install -r requirements.txt pyinstaller
    pyinstaller --onefile --windowed --name 월별개강업무도구 scripts/gui_app.py
결과물은 dist/월별개강업무도구.exe 에 생성된다.
"""
import sys
import traceback
from pathlib import Path
from tkinter import Tk, StringVar, filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

sys.path.insert(0, str(Path(__file__).resolve().parent))
import process_apply_list  # noqa: E402
import generate_company_files  # noqa: E402
import build_gaegang_sheet  # noqa: E402


class App:
    def __init__(self, root: Tk):
        self.root = root
        root.title("월별 개강 업무 도구")
        root.geometry("640x480")

        self.apply_path = StringVar()
        self.refund_path = StringVar()
        self.general_path = StringVar()

        frame = ttk.Frame(root, padding=12)
        frame.pack(fill="both", expand=True)

        self._add_file_row(frame, 0, "① 신청명단", self.apply_path)
        self._add_file_row(frame, 1, "② 환급과정 커리큘럼", self.refund_path)
        self._add_file_row(frame, 2, "③ 비환급 가격표", self.general_path)

        run_btn = ttk.Button(frame, text="실행 (저장 폴더 선택 후 1~3단계 전부 처리)", command=self.run_all)
        run_btn.grid(row=3, column=0, columnspan=3, pady=12, sticky="ew")

        ttk.Label(frame, text="처리 결과").grid(row=4, column=0, sticky="w")
        self.log = ScrolledText(frame, height=18, state="disabled")
        self.log.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(4, 0))

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(5, weight=1)

    def _add_file_row(self, parent, row, label, var: StringVar):
        ttk.Label(parent, text=label, width=20).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var, state="readonly").grid(row=row, column=1, sticky="ew", padx=6)
        ttk.Button(parent, text="파일 선택", command=lambda: self._pick_file(var)).grid(row=row, column=2)

    def _pick_file(self, var: StringVar):
        path = filedialog.askopenfilename(
            title="파일 선택", filetypes=[("엑셀/CSV 파일", "*.xlsx *.xls *.csv"), ("모든 파일", "*.*")]
        )
        if path:
            var.set(path)

    def _log(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.configure(state="disabled")
        self.log.see("end")
        self.root.update_idletasks()

    def run_all(self):
        if not (self.apply_path.get() and self.refund_path.get() and self.general_path.get()):
            messagebox.showwarning("파일을 모두 선택해주세요", "신청명단 / 환급과정 커리큘럼 / 비환급 가격표, 3개 파일을 모두 골라야 합니다.")
            return

        out_dir_str = filedialog.askdirectory(title="결과를 저장할 폴더를 선택하세요")
        if not out_dir_str:
            return
        out_dir = Path(out_dir_str)

        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        apply_path = self.apply_path.get()
        refund_path = self.refund_path.get()
        general_path = self.general_path.get()

        self._log("=== 1단계: 신청명단 정리 ===")
        try:
            r1 = process_apply_list.run(apply_path, refund_path, general_path, out_dir / "1_정리된_신청명단.xlsx")
            self._log(f"완료: 총 {r1['total']}건 (프리패스/테스트/카드결제 {r1['excluded']}건 제외) / "
                      f"환급 {r1['refund']}건 / 비환급 {r1['nonrefund']}건 / 확인필요 {r1['flagged']}건")
            self._log(f"결과 파일: {r1['output_path']}")
        except Exception:
            self._log("[오류] 1단계 실패:\n" + traceback.format_exc())

        self._log("\n=== 2단계: 기업별 발송 파일 생성 ===")
        try:
            r2 = generate_company_files.run(apply_path, refund_path, general_path, out_dir / "2_기업별발송")
            for f in r2["files"]:
                self._log(f"생성: {f['path']} ({f['count']}명)")
            self._log(f"완료: 총 {r2['total']}건 (프리패스/테스트/카드결제 {r2['excluded']}건 제외) / "
                      f"기업 {r2['company_count']}곳 / 확인필요 {r2['flagged']}건")
        except Exception:
            self._log("[오류] 2단계 실패:\n" + traceback.format_exc())

        self._log("\n=== 3단계: 개강내역시트 작성 내용 생성 ===")
        try:
            r3 = build_gaegang_sheet.run(apply_path, refund_path, general_path, out_dir / "3_개강내역.xlsx")
            self._log(f"완료: 개강내역 {r3['rows']}줄 (신청 {r3['students']}건을 묶음, "
                      f"프리패스/테스트/카드결제 {r3['excluded']}건 제외) / 확인필요 {r3['flagged']}줄")
            self._log(f"결과 파일: {r3['output_path']}")
        except Exception:
            self._log("[오류] 3단계 실패:\n" + traceback.format_exc())

        self._log(f"\n모든 작업이 끝났습니다. 저장 위치: {out_dir}")
        messagebox.showinfo("완료", f"처리가 끝났습니다.\n저장 위치: {out_dir}\n\n노란색으로 표시된 '확인필요' 줄은 직접 확인해주세요.")

        if sys.platform == "win32":
            import os
            os.startfile(out_dir)  # noqa: S606


def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
