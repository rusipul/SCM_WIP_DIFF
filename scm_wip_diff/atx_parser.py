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


REQUIRED_HEADER_LABELS = ["UNISSUE", "FE Total", "Ftape", "BE Total"]


def _find_label_column(ws, label):
    for col in range(1, ws.max_column + 1):
        if str(ws.cell(row=HEADER_ROW, column=col).value or "").strip() == label:
            return col
    return None


def get_atx_stage_groups(ws):
    positions = {label: _find_label_column(ws, label) for label in REQUIRED_HEADER_LABELS}
    missing = [label for label, col in positions.items() if col is None]
    if missing:
        raise ReportFormatError(f"필수 컬럼 헤더를 찾을 수 없습니다: {missing}")

    unissue, fe_total, ftape, be_total = (positions[label] for label in REQUIRED_HEADER_LABELS)
    return {
        "전공정": list(range(unissue, fe_total)),
        "후공정": list(range(ftape, be_total)),
    }
