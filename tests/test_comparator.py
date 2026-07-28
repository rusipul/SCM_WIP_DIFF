import pytest

from scm_wip_diff.parser import ParsedWip, parse_wip_sheet
from scm_wip_diff.atx_parser import parse_atx_wip_sheet
from scm_wip_diff.comparator import compare_stage_summary
from scm_wip_diff.comparator import compare_stage_summary_by_device
from scm_wip_diff.comparator import compare_lots
from scm_wip_diff.comparator import check_lot_overlap

FIXTURE_260721 = "tests/fixtures/260721 GTK WIP.xlsx"
FIXTURE_260722 = "tests/fixtures/260722 GTK WIP.xlsx"
FIXTURE_260723_ATX = "tests/fixtures/260723 ATX WIP.xlsx"
FIXTURE_260724_ATX = "tests/fixtures/260724 ATX WIP.xlsx"


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


def test_compare_stage_summary_omits_stages_not_present_in_source_data():
    stage_groups = {"전공정": [9], "후공정": [12]}  # "완료" 없음
    column_labels = {9: "Saw", 12: "Mold"}
    yesterday = ParsedWip(
        column_labels=column_labels,
        stage_groups=stage_groups,
        lots={("m1", "l1", "d1", 0): {9: 10, 12: 5}},
        rows={("m1", "l1", "d1", 0): 1},
    )
    today = ParsedWip(
        column_labels=column_labels,
        stage_groups=stage_groups,
        lots={("m1", "l1", "d1", 0): {9: 5, 12: 10}},
        rows={("m1", "l1", "d1", 0): 1},
    )

    summary = compare_stage_summary(yesterday, today)

    assert set(summary.keys()) == {"전공정", "후공정"}


def test_compare_stage_summary_omits_stage_present_only_in_today():
    # yesterday has no "완료" group at all; today does, with real nonzero
    # data. The yesterday-only presence check must still omit "완료" from
    # the summary rather than pick it up from today's side.
    yesterday = ParsedWip(
        column_labels={9: "Saw"},
        stage_groups={"전공정": [9]},
        lots={("m1", "l1", "d1", 0): {9: 10}},
        rows={("m1", "l1", "d1", 0): 1},
    )
    today = ParsedWip(
        column_labels={9: "Saw", 29: "TR Stock"},
        stage_groups={"전공정": [9], "완료": [29]},
        lots={("m1", "l1", "d1", 0): {9: 5, 29: 123}},
        rows={("m1", "l1", "d1", 0): 1},
    )

    summary = compare_stage_summary(yesterday, today)

    assert "완료" not in summary


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
    assert diff["new_lots"][0]["values"] == today.lots[("m2", "l2", "d2", 0)]
    assert diff["removed_lots"][0]["values"] == yesterday.lots[("m1", "l1", "d1", 0)]


def test_compare_lots_returns_process_columns_from_stage_group_union():
    yesterday = make_parsed({("m1", "l1", "d1", 0): {9: 0, 12: 0, 29: 0}})
    today = make_parsed({("m1", "l1", "d1", 0): {9: 0, 12: 0, 29: 0}})

    diff = compare_lots(yesterday, today)

    assert diff["process_columns"] == [9, 12, 29]


def test_compare_lots_handles_noncontiguous_process_columns():
    stage_groups = {"전공정": [13, 14], "후공정": [20, 21]}  # 15~19는 그룹에 없는 소계 컬럼 흉내
    column_labels = {13: "A", 14: "B", 20: "C", 21: "D"}
    yesterday = ParsedWip(
        column_labels=column_labels,
        stage_groups=stage_groups,
        lots={("m1", "l1", "d1", 0): {13: 1.0, 14: 2.0, 20: 3.0, 21: 4.0}},
        rows={("m1", "l1", "d1", 0): 1},
    )
    today = ParsedWip(
        column_labels=column_labels,
        stage_groups=stage_groups,
        lots={("m1", "l1", "d1", 0): {13: 0.0, 14: 2.0, 20: 3.0, 21: 5.0}},
        rows={("m1", "l1", "d1", 0): 1},
    )

    diff = compare_lots(yesterday, today)

    assert diff["process_columns"] == [13, 14, 20, 21]
    assert len(diff["changed_lots"]) == 1
    assert diff["changed_lots"][0]["changes"] == [
        {"col": 13, "label": "A(M)", "before": 1.0, "after": 0.0},
        {"col": 21, "label": "D(U)", "before": 4.0, "after": 5.0},
    ]


def test_compare_lots_tolerates_column_missing_from_todays_labels():
    # Simulates ATX-style drift: yesterday's stage_groups includes a column
    # that today's stage_groups (and therefore today's column_labels) no
    # longer has. The union in process_columns still includes it, so
    # compare_lots must not raise KeyError when looking up the label.
    yesterday = ParsedWip(
        column_labels={13: "A", 14: "B", 15: "E", 20: "C"},
        stage_groups={"전공정": [13, 14, 15], "후공정": [20]},
        lots={("m1", "l1", "d1", 0): {13: 1.0, 14: 2.0, 15: 5.0, 20: 3.0}},
        rows={("m1", "l1", "d1", 0): 1},
    )
    today = ParsedWip(
        column_labels={13: "A", 14: "B", 20: "C"},
        stage_groups={"전공정": [13, 14], "후공정": [20]},
        lots={("m1", "l1", "d1", 0): {13: 1.0, 14: 2.0, 20: 3.0}},
        rows={("m1", "l1", "d1", 0): 1},
    )

    diff = compare_lots(yesterday, today)

    assert diff["process_columns"] == [13, 14, 15, 20]
    assert len(diff["changed_lots"]) == 1
    assert diff["changed_lots"][0]["changes"] == [
        {"col": 15, "label": "(O)", "before": 5.0, "after": 0},
    ]


def test_compare_lots_matches_real_fixture_counts():
    yesterday = parse_wip_sheet(FIXTURE_260721)
    today = parse_wip_sheet(FIXTURE_260722)

    diff = compare_lots(yesterday, today)

    assert diff["process_columns"] == list(range(9, 31))
    assert len(diff["changed_lots"]) == 19
    assert len(diff["new_lots"]) == 1
    assert len(diff["removed_lots"]) == 2
    assert diff["new_lots"][0]["key"] == ("MNS08M6622326", "1B3C76", "TMP1230", 0)
    new_lot_key = diff["new_lots"][0]["key"]
    assert diff["new_lots"][0]["values"] == today.lots[new_lot_key]
    assert diff["new_lots"][0]["values"][9] == 665285

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


def test_compare_stage_summary_by_device_aggregates_and_sorts_by_device_name():
    yesterday = make_parsed({
        ("m1", "l1", "deviceB", 0): {9: 100, 12: 0, 29: 0},
        ("m2", "l2", "deviceA", 0): {9: 50, 12: 10, 29: 0},
        ("m3", "l3", "deviceA", 0): {9: 20, 12: 0, 29: 0},
    })
    today = make_parsed({
        ("m1", "l1", "deviceB", 0): {9: 0, 12: 100, 29: 0},
        ("m2", "l2", "deviceA", 0): {9: 50, 12: 10, 29: 0},
        ("m4", "l4", "deviceC", 0): {9: 0, 12: 0, 29: 5},
    })

    summary = compare_stage_summary_by_device(yesterday, today)

    assert list(summary.keys()) == ["deviceA", "deviceB", "deviceC"]
    assert summary["deviceA"]["전공정"] == {"yesterday": 70, "today": 50, "delta": -20}
    assert summary["deviceB"]["전공정"] == {"yesterday": 100, "today": 0, "delta": -100}
    assert summary["deviceB"]["후공정"] == {"yesterday": 0, "today": 100, "delta": 100}
    # deviceC only exists today (removed lot m3/deviceA is gone from today,
    # new lot m4/deviceC only in today) -> deviceC yesterday side is all 0
    assert summary["deviceC"]["완료"] == {"yesterday": 0, "today": 5, "delta": 5}
    assert summary["deviceC"]["전공정"] == {"yesterday": 0, "today": 0, "delta": 0}


def test_compare_stage_summary_by_device_omits_stages_not_present():
    stage_groups = {"전공정": [9], "후공정": [12]}  # "완료" 없음
    column_labels = {9: "Saw", 12: "Mold"}
    yesterday = ParsedWip(
        column_labels=column_labels,
        stage_groups=stage_groups,
        lots={("m1", "l1", "d1", 0): {9: 10, 12: 5}},
        rows={("m1", "l1", "d1", 0): 1},
        device_key_index=2,
    )
    today = ParsedWip(
        column_labels=column_labels,
        stage_groups=stage_groups,
        lots={("m1", "l1", "d1", 0): {9: 5, 12: 10}},
        rows={("m1", "l1", "d1", 0): 1},
        device_key_index=2,
    )

    summary = compare_stage_summary_by_device(yesterday, today)

    assert set(summary["d1"].keys()) == {"전공정", "후공정"}


def test_compare_stage_summary_by_device_matches_real_gtk_fixture():
    yesterday = parse_wip_sheet(FIXTURE_260721)
    today = parse_wip_sheet(FIXTURE_260722)

    summary = compare_stage_summary_by_device(yesterday, today)

    assert len(summary) == 34
    assert list(summary.keys())[0] == "TMP1200D"
    assert summary["TMP1200D"]["전공정"] == {"yesterday": 666004, "today": 666004, "delta": 0}
    assert summary["TMP1200D"]["후공정"] == {"yesterday": 339484, "today": 107484, "delta": -232000}
    assert summary["TMP1200D"]["완료"] == {"yesterday": 556627, "today": 788627, "delta": 232000}


def test_compare_stage_summary_by_device_matches_real_atx_fixture():
    yesterday = parse_atx_wip_sheet(FIXTURE_260723_ATX)
    today = parse_atx_wip_sheet(FIXTURE_260724_ATX)

    summary = compare_stage_summary_by_device(yesterday, today)

    assert len(summary) == 11
    assert list(summary.keys())[0] == "GTMP122007"
    assert summary["GTMP122007"]["전공정"]["yesterday"] == pytest.approx(1119.27, abs=0.01)
    assert summary["GTMP122007"]["전공정"]["today"] == pytest.approx(1020.23, abs=0.01)
    assert summary["GTMP122007"]["전공정"]["delta"] == pytest.approx(-99.04, abs=0.01)
    assert summary["GTMP122007"]["후공정"] == {
        "yesterday": pytest.approx(516.65, abs=0.01),
        "today": pytest.approx(615.28, abs=0.01),
        "delta": pytest.approx(98.63, abs=0.01),
    }
    assert "완료" not in summary["GTMP122007"]
