"""Compare two ParsedWip snapshots and produce a diff."""

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
