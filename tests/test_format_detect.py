import openpyxl
import pytest

from scm_wip_diff.format_detect import ATX, GTK, detect_format, parse_wip_file
from scm_wip_diff.parser import ReportFormatError

FIXTURE_GTK = "tests/fixtures/260721 GTK WIP.xlsx"
FIXTURE_ATX = "tests/fixtures/260723 ATX WIP.xlsx"


def test_detect_format_identifies_gtk_file():
    assert detect_format(FIXTURE_GTK) == GTK


def test_detect_format_identifies_atx_file():
    assert detect_format(FIXTURE_ATX) == ATX


def test_detect_format_raises_for_unrecognized_file(tmp_path):
    wb = openpyxl.Workbook()
    path = tmp_path / "blank.xlsx"
    wb.save(path)

    with pytest.raises(ReportFormatError):
        detect_format(str(path))


def test_parse_wip_file_dispatches_to_correct_parser():
    fmt, parsed = parse_wip_file(FIXTURE_GTK)
    assert fmt == GTK
    assert len(parsed.lots) == 174

    fmt, parsed = parse_wip_file(FIXTURE_ATX)
    assert fmt == ATX
    assert len(parsed.lots) == 125
