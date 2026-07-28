"""Parse GTK WIP Excel reports into structured lot-level data."""

REPORT_TITLE = "Report for Assy or Turn-key WIP"
PROCESS_COL_START = 9   # I
PROCESS_COL_END = 30    # AD


class ReportFormatError(Exception):
    """Raised when the expected WIP report table cannot be located."""


def find_report_anchor(ws, title=REPORT_TITLE):
    for row in ws.iter_rows(min_row=1, max_row=10, max_col=4):
        for cell in row:
            if cell.value and title in str(cell.value):
                return cell.row
    raise ReportFormatError(f"'{title}' 표를 찾을 수 없습니다")
