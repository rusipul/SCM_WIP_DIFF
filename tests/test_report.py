import os

from scm_wip_diff.report import derive_output_paths


def test_derive_output_paths_uses_date_prefix_and_same_folder():
    today_path = os.path.join("C:", os.sep, "data", "260722 GTK WIP.xlsx")

    report_path, highlighted_path = derive_output_paths(today_path)

    assert report_path == os.path.join("C:", os.sep, "data", "260722_GTK_WIP_변동리포트.xlsx")
    assert highlighted_path == os.path.join("C:", os.sep, "data", "260722 GTK WIP_변동표시.xlsx")
