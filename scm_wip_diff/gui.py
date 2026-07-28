"""tkinter GUI wiring parser -> comparator -> report together."""
import tkinter as tk
from tkinter import filedialog, messagebox

from scm_wip_diff.comparator import check_lot_overlap, compare_lots, compare_stage_summary
from scm_wip_diff.parser import ReportFormatError, parse_wip_sheet
from scm_wip_diff.report import (
    build_highlighted_today_file,
    build_variance_report,
    derive_output_paths,
)


class App:
    def __init__(self, root):
        self.root = root
        self.yesterday_path = tk.StringVar()
        self.today_path = tk.StringVar()

        tk.Label(root, text="어제 파일:").grid(row=0, column=0, sticky="w")
        tk.Entry(root, textvariable=self.yesterday_path, width=60).grid(row=0, column=1)
        tk.Button(root, text="찾아보기", command=self.choose_yesterday).grid(row=0, column=2)

        tk.Label(root, text="오늘 파일:").grid(row=1, column=0, sticky="w")
        tk.Entry(root, textvariable=self.today_path, width=60).grid(row=1, column=1)
        tk.Button(root, text="찾아보기", command=self.choose_today).grid(row=1, column=2)

        tk.Button(root, text="비교 실행", command=self.run_compare).grid(row=2, column=1)

        self.result_text = tk.Text(root, width=100, height=30)
        self.result_text.grid(row=3, column=0, columnspan=3)

    def choose_yesterday(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.yesterday_path.set(path)

    def choose_today(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.today_path.set(path)

    def run_compare(self):
        y_path = self.yesterday_path.get()
        t_path = self.today_path.get()
        if not y_path or not t_path:
            messagebox.showerror("입력 필요", "어제 파일과 오늘 파일을 모두 선택하세요")
            return

        try:
            try:
                yesterday = parse_wip_sheet(y_path)
            except ReportFormatError as e:
                messagebox.showerror("파일 형식 오류", f"어제 파일 오류 ({y_path}):\n{e}")
                return

            try:
                today = parse_wip_sheet(t_path)
            except ReportFormatError as e:
                messagebox.showerror("파일 형식 오류", f"오늘 파일 오류 ({t_path}):\n{e}")
                return

            overlap_warning = check_lot_overlap(yesterday, today)
            if overlap_warning:
                messagebox.showwarning("확인 필요", overlap_warning)

            stage_summary = compare_stage_summary(yesterday, today)
            lot_diff = compare_lots(yesterday, today)

            report_path, highlighted_path = derive_output_paths(t_path)

            try:
                build_variance_report(stage_summary, lot_diff, report_path, today.column_labels)
            except PermissionError:
                messagebox.showerror(
                    "저장 실패",
                    f"변동 리포트 파일이 열려있어 저장할 수 없습니다:\n{report_path}",
                )
                return

            try:
                build_highlighted_today_file(t_path, lot_diff, highlighted_path)
            except PermissionError:
                messagebox.showerror(
                    "저장 실패",
                    "변동 리포트는 저장되었지만, 하이라이트 파일이 열려있어 저장할 수 없습니다.\n\n"
                    f"저장됨: {report_path}\n실패: {highlighted_path}",
                )
                return

            self._show_preview(stage_summary, lot_diff)
            messagebox.showinfo("완료", f"저장 완료:\n{report_path}\n{highlighted_path}")
        except Exception as e:
            messagebox.showerror("예상치 못한 오류", str(e))

    def _show_preview(self, stage_summary, lot_diff):
        self.result_text.delete("1.0", tk.END)
        lines = ["[단계별 합계]"]
        for stage in ("전공정", "후공정", "완료"):
            s = stage_summary[stage]
            lines.append(f"{stage}: {s['yesterday']:,} -> {s['today']:,} ({s['delta']:+,})")
        lines.append("")
        lines.append(f"변경된 랏 수: {len(lot_diff['changed_lots'])}")
        lines.append(f"신규 랏 수: {len(lot_diff['new_lots'])}")
        lines.append(f"삭제된 랏 수: {len(lot_diff['removed_lots'])}")
        self.result_text.insert("1.0", "\n".join(lines))
