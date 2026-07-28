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
