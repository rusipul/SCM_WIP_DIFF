import openpyxl
import pytest

from scm_wip_diff.parser import find_report_anchor, get_stage_groups, ReportFormatError

FIXTURE_260721 = "tests/fixtures/260721 GTK WIP.xlsx"
FIXTURE_260722 = "tests/fixtures/260722 GTK WIP.xlsx"


def test_find_report_anchor_locates_title_row():
    wb = openpyxl.load_workbook(FIXTURE_260721, data_only=True)
    ws = wb[wb.sheetnames[0]]

    anchor = find_report_anchor(ws)

    assert anchor == 2


def test_find_report_anchor_raises_when_title_missing():
    wb = openpyxl.Workbook()
    ws = wb.active

    with pytest.raises(ReportFormatError):
        find_report_anchor(ws)


def test_get_stage_groups_matches_merged_header_cells():
    wb = openpyxl.load_workbook(FIXTURE_260721, data_only=True)
    ws = wb[wb.sheetnames[0]]
    anchor = find_report_anchor(ws)

    groups = get_stage_groups(ws, anchor)

    assert groups["전공정"] == [9, 10, 11]
    assert groups["후공정"] == list(range(12, 29))
    assert groups["완료"] == [29, 30]
