"""Detect whether a WIP Excel file is GTK or ATX format and parse it accordingly."""
import openpyxl

from scm_wip_diff.atx_parser import find_atx_sheet, parse_atx_wip_sheet
from scm_wip_diff.parser import ReportFormatError, find_report_anchor, parse_wip_sheet

GTK = "GTK"
ATX = "ATX"


def detect_format(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    try:
        find_atx_sheet(wb)
        return ATX
    except ReportFormatError:
        pass

    ws = wb[wb.sheetnames[0]]
    try:
        find_report_anchor(ws)
        return GTK
    except ReportFormatError:
        pass

    raise ReportFormatError("GTK 또는 ATX 형식으로 인식할 수 없는 파일입니다")


def parse_wip_file(path):
    fmt = detect_format(path)
    if fmt == ATX:
        return fmt, parse_atx_wip_sheet(path)
    return fmt, parse_wip_sheet(path)
