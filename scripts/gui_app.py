#!/usr/bin/env python3
"""월별 개강 업무 도구 - 화면(GUI) 버전.

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
from tkinter import BooleanVar, Tk, StringVar, filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

sys.path.insert(0, str(Path(__file__).resolve().parent))
import process_apply_list  # noqa: E402
import generate_company_files  # noqa: E402
import add_applications  # noqa: E402
import build_gaegang_sheet  # noqa: E402
import tel_list  # noqa: E402


class App:
    def __init__(self, root: Tk):
        self.root = root
        root.title("월별 개강 업무 도구")
        root.geometry("640x620")

        self.apply_path = StringVar()
        self.refund_path = StringVar()
        self.general_path = StringVar()
        self.credit_path = StringVar()
        self.save_tel = BooleanVar(value=False)
        self.added_path = StringVar()
        self.existing_gaegang_path = StringVar()

        frame = ttk.Frame(root, padding=12)
        frame.pack(fill="both", expand=True)

        self._add_file_row(frame, 0, "① 신청명단", self.apply_path)
        self._add_file_row(frame, 1, "② 환급과정 커리큘럼", self.refund_path)
        self._add_file_row(frame, 2, "③ 비환급 가격표", self.general_path)
        ttk.Checkbutton(frame, text="도쿄일렉트론코리아 별도 파일 저장",
                        variable=self.save_tel).grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self._add_file_row(frame, 4, "④ 학점 파일", self.credit_path)

        run_btn = ttk.Button(frame, text="실행 (저장 폴더 선택 후 전부 처리)", command=self.run_all)
        run_btn.grid(row=5, column=0, columnspan=3, pady=12, sticky="ew")

        ttk.Separator(frame).grid(row=6, column=0, columnspan=3, sticky="ew", pady=4)
        ttk.Label(frame, text="개별 추가 신청 (①②③ 파일도 함께 사용)").grid(row=7, column=0, columnspan=3, sticky="w")
        self._add_file_row(frame, 8, "⑤ 추가 신청 파일", self.added_path)
        self._add_file_row(frame, 9, "⑥ 기존 개강내역 파일 (선택)", self.existing_gaegang_path)
        add_btn = ttk.Button(frame, text="추가 신청 반영 (저장 폴더 선택)", command=self.run_added)
        add_btn.grid(row=10, column=0, columnspan=3, pady=12, sticky="ew")

        ttk.Label(frame, text="처리 결과").grid(row=11, column=0, sticky="w")
        self.log = ScrolledText(frame, height=14, state="disabled")
        self.log.grid(row=12, column=0, columnspan=3, sticky="nsew", pady=(4, 0))

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(12, weight=1)

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

    def _clear_file_selections(self):
        """작업이 끝나면 골랐던 파일들을 비워서 다음 작업 때 다시 고르게 한다."""
        for var in (self.apply_path, self.refund_path, self.general_path, self.credit_path,
                    self.added_path, self.existing_gaegang_path):
            var.set("")

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
        if self.save_tel.get() and not self.credit_path.get():
            messagebox.showwarning("학점 파일을 선택해주세요", "도쿄일렉트론코리아 별도 파일을 만들려면 ④ 학점 파일을 골라야 합니다.")
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

        if self.save_tel.get():
            self._log("\n=== 도쿄일렉트론코리아 별도 파일 ===")
            try:
                r4 = tel_list.run(apply_path, general_path, self.credit_path.get(),
                                  out_dir / "4_도쿄일렉트론코리아.xlsx")
                self._log(f"완료: {r4['total']}건 / 확인필요 {r4['flagged']}건")
                self._log(f"결과 파일: {r4['output_path']}")
            except Exception:
                self._log("[오류] 도쿄일렉트론코리아 파일 생성 실패:\n" + traceback.format_exc())

        self._log(f"\n모든 작업이 끝났습니다. 저장 위치: {out_dir}")
        messagebox.showinfo("완료", f"처리가 끝났습니다.\n저장 위치: {out_dir}\n\n노란색으로 표시된 '확인필요' 줄은 직접 확인해주세요.")
        self._clear_file_selections()

        if sys.platform == "win32":
            import os
            os.startfile(out_dir)  # noqa: S606


    def run_added(self):
        if not (self.apply_path.get() and self.refund_path.get() and self.general_path.get()
                and self.added_path.get()):
            messagebox.showwarning("파일을 모두 선택해주세요",
                                   "① 신청명단(원본) / ② 환급과정 커리큘럼 / ③ 비환급 가격표 / "
                                   "⑤ 추가 신청 파일을 모두 골라야 합니다. "
                                   "⑥ 기존 개강내역 파일은 선택사항입니다.")
            return
        if self.save_tel.get() and not self.credit_path.get():
            messagebox.showwarning("학점 파일을 선택해주세요", "도쿄일렉트론코리아 별도 파일을 만들려면 ④ 학점 파일을 골라야 합니다.")
            return

        out_dir_str = filedialog.askdirectory(title="추가 결과를 저장할 폴더를 선택하세요")
        if not out_dir_str:
            return
        out_dir = Path(out_dir_str)

        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        add_applications.run(self.apply_path.get(), self.added_path.get(), self.refund_path.get(),
                             self.general_path.get(), self.existing_gaegang_path.get(), out_dir,
                             credit_path=self.credit_path.get() if self.save_tel.get() else None,
                             log=self._log)

        gaegang_msg = ("개강내역은 고른 파일을 갱신했습니다." if self.existing_gaegang_path.get()
                       else "⑥을 고르지 않아 개강내역은 새 파일로 만들었고, 교육회차는 비어 있습니다.")
        self._log(f"\n추가 신청 반영이 끝났습니다. 저장 위치: {out_dir}")
        messagebox.showinfo("완료", f"추가 신청 반영이 끝났습니다.\n저장 위치: {out_dir}\n\n"
                                    f"{gaegang_msg} 노란색 '확인필요' 줄은 직접 확인해주세요.")
        self._clear_file_selections()

        if sys.platform == "win32":
            import os
            os.startfile(out_dir)  # noqa: S606


def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
