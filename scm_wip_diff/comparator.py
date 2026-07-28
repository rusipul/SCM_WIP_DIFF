"""Compare two ParsedWip snapshots and produce a diff."""

from openpyxl.utils import get_column_letter

STAGE_ORDER = ["전공정", "후공정", "완료"]


def compare_stage_summary(yesterday, today):
    summary = {}
    for stage in STAGE_ORDER:
        y_cols = yesterday.stage_groups.get(stage, [])
        t_cols = today.stage_groups.get(stage, [])
        y_total = sum(sum(v.get(c, 0) for c in y_cols) for v in yesterday.lots.values())
        t_total = sum(sum(v.get(c, 0) for c in t_cols) for v in today.lots.values())
        summary[stage] = {
            "yesterday": y_total,
            "today": t_total,
            "delta": t_total - y_total,
        }
    return summary


def _process_columns(yesterday, today):
    y_cols = {c for cols in yesterday.stage_groups.values() for c in cols}
    t_cols = {c for cols in today.stage_groups.values() for c in cols}
    return sorted(y_cols | t_cols)


def compare_lots(yesterday, today):
    y_keys = set(yesterday.lots.keys())
    t_keys = set(today.lots.keys())
    common_keys = y_keys & t_keys
    process_cols = _process_columns(yesterday, today)

    changed_lots = []
    for key in sorted(common_keys, key=lambda k: today.rows[k]):
        y_vals = yesterday.lots[key]
        t_vals = today.lots[key]
        changes = []
        for col in process_cols:
            before = y_vals.get(col, 0)
            after = t_vals.get(col, 0)
            if before != after:
                label = f"{today.column_labels[col]}({get_column_letter(col)})"
                changes.append({"col": col, "label": label, "before": before, "after": after})
        if changes:
            changed_lots.append({
                "key": key,
                "row_in_today": today.rows[key],
                "changes": changes,
            })

    new_lots = [
        {"key": key, "row_in_today": today.rows[key], "values": today.lots[key]}
        for key in sorted(t_keys - y_keys, key=lambda k: today.rows[k])
    ]
    removed_lots = [
        {"key": key, "row_in_yesterday": yesterday.rows[key], "values": yesterday.lots[key]}
        for key in sorted(y_keys - t_keys, key=lambda k: yesterday.rows[k])
    ]

    return {
        "changed_lots": changed_lots,
        "new_lots": new_lots,
        "removed_lots": removed_lots,
        "process_columns": process_cols,
    }


OVERLAP_WARNING_THRESHOLD = 0.5


def check_lot_overlap(yesterday, today):
    y_keys = set(yesterday.lots.keys())
    t_keys = set(today.lots.keys())
    if not y_keys or not t_keys:
        return None

    overlap = len(y_keys & t_keys)
    ratio = overlap / min(len(y_keys), len(t_keys))
    if ratio < OVERLAP_WARNING_THRESHOLD:
        return (
            f"어제/오늘 파일 간 겹치는 랏이 적습니다 ({ratio:.0%}). "
            "잘못된 파일을 선택하지 않았는지 확인하세요."
        )
    return None
