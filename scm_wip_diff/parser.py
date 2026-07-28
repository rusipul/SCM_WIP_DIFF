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


def get_stage_groups(ws, anchor_row):
    group_row = anchor_row + 1
    groups = {}
    for merged in ws.merged_cells.ranges:
        if merged.min_row == group_row:
            label = str(ws.cell(row=group_row, column=merged.min_col).value).strip()
            groups[label] = list(range(merged.min_col, merged.max_col + 1))
    return groups


def build_column_labels(ws, anchor_row):
    row1 = anchor_row + 2
    row2 = anchor_row + 3
    labels = {}
    for col in range(1, ws.max_column + 1):
        v1 = ws.cell(row=row1, column=col).value
        v2 = ws.cell(row=row2, column=col).value
        v1 = str(v1).strip() if v1 else ""
        v2 = str(v2).strip() if v2 else ""
        labels[col] = f"{v1} {v2}".strip() if v1 else v2
    return labels
