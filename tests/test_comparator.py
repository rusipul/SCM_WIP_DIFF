from scm_wip_diff.parser import ParsedWip, parse_wip_sheet
from scm_wip_diff.comparator import compare_stage_summary
from scm_wip_diff.comparator import compare_lots
from scm_wip_diff.comparator import check_lot_overlap

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


def test_compare_lots_detects_changed_columns_in_column_order():
    yesterday = make_parsed({
        ("m1", "l1", "d1", 0): {9: 347638, 12: 0, 29: 0},
    })
    today = make_parsed({
        ("m1", "l1", "d1", 0): {9: 0, 12: 316800, 29: 0},
    })

    diff = compare_lots(yesterday, today)

    assert len(diff["changed_lots"]) == 1
    lot = diff["changed_lots"][0]
    assert lot["key"] == ("m1", "l1", "d1", 0)
    assert lot["changes"] == [
        {"col": 9, "label": "Saw(I)", "before": 347638, "after": 0},
        {"col": 12, "label": "Mold(L)", "before": 0, "after": 316800},
    ]


def test_compare_lots_detects_new_and_removed_keys():
    yesterday = make_parsed({
        ("m1", "l1", "d1", 0): {9: 0, 12: 0, 29: 0},
    })
    today = make_parsed({
        ("m2", "l2", "d2", 0): {9: 0, 12: 0, 29: 0},
    })

    diff = compare_lots(yesterday, today)

    assert diff["changed_lots"] == []
    assert [lot["key"] for lot in diff["new_lots"]] == [("m2", "l2", "d2", 0)]
    assert [lot["key"] for lot in diff["removed_lots"]] == [("m1", "l1", "d1", 0)]


def test_compare_lots_matches_real_fixture_counts():
    yesterday = parse_wip_sheet(FIXTURE_260721)
    today = parse_wip_sheet(FIXTURE_260722)

    diff = compare_lots(yesterday, today)

    assert len(diff["changed_lots"]) == 19
    assert len(diff["new_lots"]) == 1
    assert len(diff["removed_lots"]) == 2
    assert diff["new_lots"][0]["key"] == ("MNS08M6622326", "1B3C76", "TMP1230", 0)

    saw_to_die_mount = next(
        lot for lot in diff["changed_lots"]
        if lot["key"] == ("TNT0896622324", "1B3C75", "TMP1202D", 0)
    )
    assert saw_to_die_mount["row_in_today"] == 21
    assert saw_to_die_mount["changes"] == [
        {"col": 9, "label": "Saw(I)", "before": 347638, "after": 0},
        {"col": 10, "label": "Die Mount(J)", "before": 316800, "after": 664438},
    ]


def test_check_lot_overlap_returns_none_when_files_mostly_match():
    yesterday = make_parsed({(f"m{i}", f"l{i}", "d", 0): {9: 0, 12: 0, 29: 0} for i in range(10)})
    today = make_parsed({(f"m{i}", f"l{i}", "d", 0): {9: 0, 12: 0, 29: 0} for i in range(9)})

    assert check_lot_overlap(yesterday, today) is None


def test_check_lot_overlap_warns_when_files_mostly_differ():
    yesterday = make_parsed({(f"m{i}", f"l{i}", "d", 0): {9: 0, 12: 0, 29: 0} for i in range(10)})
    today = make_parsed({(f"n{i}", f"k{i}", "d", 0): {9: 0, 12: 0, 29: 0} for i in range(10)})

    warning = check_lot_overlap(yesterday, today)

    assert warning is not None
    assert "겹치는 랏이 적습니다" in warning


def test_check_lot_overlap_matches_real_fixtures_without_warning():
    yesterday = parse_wip_sheet(FIXTURE_260721)
    today = parse_wip_sheet(FIXTURE_260722)

    assert check_lot_overlap(yesterday, today) is None
