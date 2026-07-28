import openpyxl
import pytest

from scm_wip_diff.atx_parser import find_atx_sheet, get_atx_stage_groups
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


def test_get_atx_stage_groups_uses_header_text_not_merged_cells():
    wb = openpyxl.load_workbook(FIXTURE_260723, data_only=True)
    ws = find_atx_sheet(wb)

    groups = get_atx_stage_groups(ws)

    assert groups["전공정"] == list(range(13, 23))
    assert groups["후공정"] == list(range(24, 47))


def test_get_atx_stage_groups_matches_across_both_fixtures_despite_merge_drift():
    wb_723 = openpyxl.load_workbook(FIXTURE_260723, data_only=True)
    wb_724 = openpyxl.load_workbook(FIXTURE_260724, data_only=True)

    groups_723 = get_atx_stage_groups(find_atx_sheet(wb_723))
    groups_724 = get_atx_stage_groups(find_atx_sheet(wb_724))

    assert groups_723 == groups_724


def test_get_atx_stage_groups_raises_when_required_label_missing():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=2, column=13, value="UNISSUE")
    ws.cell(row=2, column=23, value="FE Total")
    ws.cell(row=2, column=24, value="Ftape")
    # "BE Total" 헤더를 의도적으로 생략해 포맷 변경을 시뮬레이션

    with pytest.raises(ReportFormatError):
        get_atx_stage_groups(ws)
