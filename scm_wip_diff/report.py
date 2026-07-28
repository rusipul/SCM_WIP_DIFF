"""Generate the variance report workbook and the highlighted today-file copy."""
import os
import re
import shutil

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from scm_wip_diff.parser import PROCESS_COL_START, PROCESS_COL_END

RED_FILL = PatternFill(fill_type="solid", fgColor="FFFF0000")
NEW_LOT_FILL = PatternFill(fill_type="solid", fgColor="FFADD8E6")

STAGE_ORDER = ["전공정", "후공정", "완료"]


def derive_output_paths(today_path):
    folder = os.path.dirname(today_path)
    basename = os.path.basename(today_path)
    name, ext = os.path.splitext(basename)
    match = re.match(r"^(\d{6})", basename)
    date_prefix = match.group(1) if match else name
    report_path = os.path.join(folder, f"{date_prefix}_GTK_WIP_변동리포트.xlsx")
    highlighted_path = os.path.join(folder, f"{name}_변동표시{ext}")
    return report_path, highlighted_path


PROCESS_COLS = range(PROCESS_COL_START, PROCESS_COL_END + 1)


def _process_column_headers(column_labels):
    return [f"{column_labels.get(col, '')}({get_column_letter(col)})" for col in PROCESS_COLS]


def build_variance_report(stage_summary, lot_diff, output_path, column_labels):
    wb = openpyxl.Workbook()

    summary_ws = wb.active
    summary_ws.title = "요약"
    summary_ws.append(["단계", "어제", "오늘", "증감"])
    for stage in STAGE_ORDER:
        if stage in stage_summary:
            s = stage_summary[stage]
            summary_ws.append([stage, s["yesterday"], s["today"], s["delta"]])
    summary_ws.append([])
    summary_ws.append(["변경된 랏 수", len(lot_diff["changed_lots"])])
    summary_ws.append(["신규 랏 수", len(lot_diff["new_lots"])])
    summary_ws.append(["삭제된 랏 수", len(lot_diff["removed_lots"])])

    changed_ws = wb.create_sheet("변동랏")
    changed_ws.append(["MO", "랏번호", "디바이스", "변경컬럼", "어제값", "오늘값"])
    for lot in lot_diff["changed_lots"]:
        mo, lot_no, device, _ = lot["key"]
        for change in lot["changes"]:
            changed_ws.append([mo, lot_no, device, change["label"], change["before"], change["after"]])

    process_headers = _process_column_headers(column_labels)

    new_ws = wb.create_sheet("신규랏")
    new_ws.append(["MO", "랏번호", "디바이스"] + process_headers)
    for lot in lot_diff["new_lots"]:
        mo, lot_no, device, _ = lot["key"]
        values = lot["values"]
        new_ws.append([mo, lot_no, device] + [values.get(col, 0) for col in PROCESS_COLS])

    removed_ws = wb.create_sheet("삭제랏")
    removed_ws.append(["MO", "랏번호", "디바이스"] + process_headers)
    for lot in lot_diff["removed_lots"]:
        mo, lot_no, device, _ = lot["key"]
        values = lot["values"]
        removed_ws.append([mo, lot_no, device] + [values.get(col, 0) for col in PROCESS_COLS])

    wb.save(output_path)


def build_highlighted_today_file(today_path, lot_diff, output_path):
    shutil.copyfile(today_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    ws = wb[wb.sheetnames[0]]

    for lot in lot_diff["changed_lots"]:
        row = lot["row_in_today"]
        for change in lot["changes"]:
            ws.cell(row=row, column=change["col"]).fill = RED_FILL

    for lot in lot_diff["new_lots"]:
        row = lot["row_in_today"]
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = NEW_LOT_FILL

    wb.save(output_path)
