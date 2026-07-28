import openpyxl
import pytest

from scm_wip_diff.atx_parser import find_atx_sheet
from scm_wip_diff.parser import ReportFormatError

FIXTURE_260723 = "tests/fixtures/260723 ATX WIP.xlsx"
FIXTURE_260724 = "tests/fixtures/260724 ATX WIP.xlsx"


def test_find_atx_sheet_matches_by_prefix_regardless_of_suffix():
    wb_723 = openpyxl.load_workbook(FIXTURE_260723, data_only=True)
    wb_724 = openpyxl.load_workbook(FIXTURE_260724, data_only=True)

    assert find_atx_sheet(wb_723).title == "KSWIPAY (PKG)"
    assert find_atx_sheet(wb_724).title == "KSWIPAY"


def test_find_atx_sheet_raises_when_no_matching_sheet():
    wb = openpyxl.Workbook()

    with pytest.raises(ReportFormatError):
        find_atx_sheet(wb)
