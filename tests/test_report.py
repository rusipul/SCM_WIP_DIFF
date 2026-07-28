import os

import openpyxl

from scm_wip_diff.report import build_variance_report, derive_output_paths


def test_derive_output_paths_uses_date_prefix_and_same_folder():
    today_path = os.path.join("C:", os.sep, "data", "260722 GTK WIP.xlsx")

    report_path, highlighted_path = derive_output_paths(today_path)

    assert report_path == os.path.join("C:", os.sep, "data", "260722_GTK_WIP_변동리포트.xlsx")
    assert highlighted_path == os.path.join("C:", os.sep, "data", "260722 GTK WIP_변동표시.xlsx")


STAGE_SUMMARY = {
    "전공정": {"yesterday": 100, "today": 50, "delta": -50},
    "후공정": {"yesterday": 10, "today": 110, "delta": 100},
    "완료": {"yesterday": 0, "today": 0, "delta": 0},
}

COLUMN_LABELS = {9: "Saw", 10: "Die Mount", 12: "Mold"}

LOT_DIFF = {
    "changed_lots": [
        {
            "key": ("m1", "l1", "d1", 0),
            "row_in_today": 7,
            "changes": [{"col": 9, "label": "Saw(I)", "before": 347638, "after": 0}],
        },
        {
            "key": ("m4", "l4", "d4", 0),
            "row_in_today": 10,
            "changes": [
                {"col": 9, "label": "Saw(I)", "before": 100, "after": 0},
                {"col": 10, "label": "Die Mount(J)", "before": 0, "after": 100},
            ],
        },
    ],
    "new_lots": [{"key": ("m2", "l2", "d2", 0), "row_in_today": 8, "values": {9: 123, 12: 456}}],
    "removed_lots": [{"key": ("m3", "l3", "d3", 0), "row_in_yesterday": 9, "values": {9: 789, 12: 321}}],
}


def test_build_variance_report_writes_expected_sheets(tmp_path):
    output_path = tmp_path / "report.xlsx"

    build_variance_report(STAGE_SUMMARY, LOT_DIFF, str(output_path), COLUMN_LABELS)

    wb = openpyxl.load_workbook(str(output_path))
    assert wb.sheetnames == ["요약", "변동랏", "신규랏", "삭제랏"]

    summary_ws = wb["요약"]
    assert summary_ws["A2"].value == "전공정"
    assert summary_ws["B2"].value == 100
    assert summary_ws["C2"].value == 50
    assert summary_ws["D2"].value == -50

    changed_ws = wb["변동랏"]
    assert changed_ws["B2"].value == "l1"
    assert changed_ws["D2"].value == "Saw(I)"
    assert changed_ws["E2"].value == 347638
    assert changed_ws["F2"].value == 0

    # second changed lot has 2 changes -> flattens into rows 3 and 4
    assert changed_ws["B3"].value == "l4"
    assert changed_ws["D3"].value == "Saw(I)"
    assert changed_ws["E3"].value == 100
    assert changed_ws["F3"].value == 0
    assert changed_ws["B4"].value == "l4"
    assert changed_ws["D4"].value == "Die Mount(J)"
    assert changed_ws["E4"].value == 0
    assert changed_ws["F4"].value == 100

    new_ws = wb["신규랏"]
    assert new_ws["D1"].value == "Saw(I)"
    assert new_ws["E1"].value == "Die Mount(J)"
    assert new_ws["B2"].value == "l2"
    assert new_ws["D2"].value == 123
    assert new_ws["E2"].value == 0

    removed_ws = wb["삭제랏"]
    assert removed_ws["D1"].value == "Saw(I)"
    assert removed_ws["B2"].value == "l3"
    assert removed_ws["D2"].value == 789
    assert removed_ws["E2"].value == 0
