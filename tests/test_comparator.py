from scm_wip_diff.parser import ParsedWip, parse_wip_sheet
from scm_wip_diff.comparator import compare_stage_summary

FIXTURE_260721 = "tests/fixtures/260721 GTK WIP.xlsx"
FIXTURE_260722 = "tests/fixtures/260722 GTK WIP.xlsx"


def make_parsed(lots):
    stage_groups = {"전공정": [9], "후공정": [12], "완료": [29]}
    column_labels = {9: "Saw", 12: "Mold", 29: "TR Stock"}
    rows = {key: i + 1 for i, key in enumerate(lots)}
    return ParsedWip(column_labels=column_labels, stage_groups=stage_groups, lots=lots, rows=rows)


def test_compare_stage_summary_sums_each_stage_group():
    yesterday = make_parsed({
        ("m1", "l1", "d1", 0): {9: 100, 12: 0, 29: 0},
        ("m2", "l2", "d2", 0): {9: 50, 12: 10, 29: 0},
    })
    today = make_parsed({
        ("m1", "l1", "d1", 0): {9: 0, 12: 100, 29: 0},
        ("m2", "l2", "d2", 0): {9: 50, 12: 10, 29: 0},
    })

    summary = compare_stage_summary(yesterday, today)

    assert summary["전공정"] == {"yesterday": 150, "today": 50, "delta": -100}
    assert summary["후공정"] == {"yesterday": 10, "today": 110, "delta": 100}
    assert summary["완료"] == {"yesterday": 0, "today": 0, "delta": 0}


def test_compare_stage_summary_matches_real_fixture_totals():
    yesterday = parse_wip_sheet(FIXTURE_260721)
    today = parse_wip_sheet(FIXTURE_260722)

    summary = compare_stage_summary(yesterday, today)

    assert summary["전공정"] == {"yesterday": 2960364, "today": 3438725, "delta": 478361}
    assert summary["후공정"] == {"yesterday": 2010299, "today": 1782025, "delta": -228274}
    assert summary["완료"] == {"yesterday": 2948085, "today": 3130907, "delta": 182822}
