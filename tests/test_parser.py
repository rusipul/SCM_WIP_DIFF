import openpyxl
import pytest

from scm_wip_diff.parser import find_report_anchor, ReportFormatError

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
