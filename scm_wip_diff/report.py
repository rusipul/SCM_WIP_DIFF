"""Generate the variance report workbook and the highlighted today-file copy."""
import os
import re
import shutil

import openpyxl
from openpyxl.styles import PatternFill

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


def build_variance_report(stage_summary, lot_diff, output_path):
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

    new_ws = wb.create_sheet("신규랏")
    new_ws.append(["MO", "랏번호", "디바이스"])
    for lot in lot_diff["new_lots"]:
        mo, lot_no, device, _ = lot["key"]
        new_ws.append([mo, lot_no, device])

    removed_ws = wb.create_sheet("삭제랏")
    removed_ws.append(["MO", "랏번호", "디바이스"])
    for lot in lot_diff["removed_lots"]:
        mo, lot_no, device, _ = lot["key"]
        removed_ws.append([mo, lot_no, device])

    wb.save(output_path)
