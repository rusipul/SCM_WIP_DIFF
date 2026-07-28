"""Parse ATX WIP Excel reports (KSWIPAY sheet) into structured lot-level data."""

from scm_wip_diff.parser import ParsedWip, ReportFormatError

ATX_SHEET_PREFIX = "KSWIPAY"
HEADER_ROW = 2
DATA_START_ROW = 3
VALUE_NUMBER_FORMAT = "#,##0.00"


def find_atx_sheet(wb):
    for name in wb.sheetnames:
        if name.startswith(ATX_SHEET_PREFIX):
            return wb[name]
    raise ReportFormatError(f"'{ATX_SHEET_PREFIX}'로 시작하는 시트를 찾을 수 없습니다")
