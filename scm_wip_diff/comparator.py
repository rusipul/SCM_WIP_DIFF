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


def compare_lots(yesterday, today):
    y_keys = set(yesterday.lots.keys())
    t_keys = set(today.lots.keys())
    common_keys = y_keys & t_keys

    changed_lots = []
    for key in sorted(common_keys, key=lambda k: today.rows[k]):
        y_vals = yesterday.lots[key]
        t_vals = today.lots[key]
        changes = []
        for col in range(9, 31):
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
        {"key": key, "row_in_today": today.rows[key]}
        for key in sorted(t_keys - y_keys, key=lambda k: today.rows[k])
    ]
    removed_lots = [
        {"key": key, "row_in_yesterday": yesterday.rows[key]}
        for key in sorted(y_keys - t_keys, key=lambda k: yesterday.rows[k])
    ]

    return {
        "changed_lots": changed_lots,
        "new_lots": new_lots,
        "removed_lots": removed_lots,
    }
