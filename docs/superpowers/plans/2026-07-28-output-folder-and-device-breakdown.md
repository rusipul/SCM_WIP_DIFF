# 저장 폴더 선택 + 디바이스별 단계 수량 표시 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 사용자가 변동리포트/하이라이트 파일의 저장 폴더를 GUI에서 직접 지정할 수 있게 하고, GUI 미리보기와 변동리포트의 "요약" 시트에 전공정/후공정/완료 단계별 수량을 디바이스별로도 보여준다.

**Architecture:** `derive_output_paths`가 폴더를 인자로 받도록 바꾸고, `comparator.py`에 디바이스별 집계 함수를 추가하며, `build_variance_report`가 그 결과를 받아 "요약" 시트 아래쪽에 추가 섹션으로 렌더링한다. `ParsedWip`에 `device_key_index`(GTK=2, ATX=1) 필드를 추가해 랏 키 튜플에서 디바이스명을 포맷에 안전하게 뽑아낸다.

**Tech Stack:** Python 3.14, openpyxl, tkinter, pytest

**설계서:** `docs/superpowers/specs/2026-07-28-output-folder-and-device-breakdown-design.md`

---

## 파일 구조

```
scm_wip_diff/
├── parser.py          # 수정: ParsedWip에 device_key_index 필드 추가 (GTK=2)
├── atx_parser.py       # 수정: device_key_index=1 채우기
├── comparator.py       # 수정: compare_stage_summary_by_device 추가
├── report.py            # 수정: derive_output_paths에 output_folder 인자, build_variance_report에 device_summary 인자
└── gui.py               # 수정: 저장 폴더 선택 UI, 디바이스별 미리보기
tests/
├── test_parser.py        # 수정 (device_key_index 테스트)
├── test_atx_parser.py     # 수정 (device_key_index 테스트)
├── test_comparator.py     # 수정 (compare_stage_summary_by_device 테스트)
└── test_report.py          # 수정 (output_folder, device_summary 테스트)
```

## 핵심 데이터 구조 변경

```python
# parser.ParsedWip (필드 1개 추가)
@dataclass
class ParsedWip:
    column_labels: dict
    stage_groups: dict
    lots: dict
    rows: dict
    sheet_name: str = ""
    value_number_format: str = "#,##0"
    key_labels: tuple = ("MO", "랏번호", "디바이스")
    device_key_index: int = 2

# comparator.compare_stage_summary_by_device(yesterday, today) 반환값
{
  "TMP1200D": {
    "전공정": {"yesterday": 666004, "today": 666004, "delta": 0},
    "후공정": {"yesterday": 339484, "today": 107484, "delta": -232000},
    "완료": {"yesterday": 556627, "today": 788627, "delta": 232000},
  },
  ... (디바이스명 가나다순 정렬)
}

# report.derive_output_paths(today_path, output_folder) 반환값: 변경 없음 (report_path, highlighted_path 튜플), 폴더 부분만 output_folder 사용
```

---

### Task 1: parser.py / atx_parser.py — `device_key_index` 필드 추가

**Files:**
- Modify: `scm_wip_diff/parser.py`
- Modify: `scm_wip_diff/atx_parser.py`
- Test: `tests/test_parser.py`
- Test: `tests/test_atx_parser.py`

**배경:** GTK 랏 키는 `(MO, 랏번호, 디바이스, 순번)`으로 디바이스가 인덱스 2, ATX 랏 키는 `(웨이퍼랏, 디바이스, 컨트롤랏, 순번)`으로 디바이스가 인덱스 1이다. 디바이스별로 집계하려면 이 인덱스를 포맷별로 알아야 한다 (이전에 `key_labels`로 헤더 문제를 해결한 것과 같은 종류의 포맷 차이).

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/test_parser.py`에 추가:
```python
def test_parse_wip_sheet_populates_device_key_index():
    parsed = parse_wip_sheet(FIXTURE_260721)

    assert parsed.device_key_index == 2
```

`tests/test_atx_parser.py`에 추가:
```python
def test_parse_atx_wip_sheet_populates_device_key_index():
    parsed = parse_atx_wip_sheet(FIXTURE_260723)

    assert parsed.device_key_index == 1
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_parser.py::test_parse_wip_sheet_populates_device_key_index tests/test_atx_parser.py::test_parse_atx_wip_sheet_populates_device_key_index -v`
Expected: FAIL with `AttributeError: 'ParsedWip' object has no attribute 'device_key_index'`

- [ ] **Step 3: ParsedWip에 필드 추가**

`scm_wip_diff/parser.py`의 `ParsedWip` 데이터클래스를 다음으로 교체:
```python
@dataclass
class ParsedWip:
    column_labels: dict
    stage_groups: dict
    lots: dict
    rows: dict
    sheet_name: str = ""
    value_number_format: str = "#,##0"
    key_labels: tuple = ("MO", "랏번호", "디바이스")
    device_key_index: int = 2
```

`parse_wip_sheet`의 `return ParsedWip(...)` 호출부를 다음으로 교체:
```python
    return ParsedWip(
        column_labels=labels,
        stage_groups=groups,
        lots=lots,
        rows=rows,
        sheet_name=ws.title,
        value_number_format="#,##0",
        key_labels=("MO", "랏번호", "디바이스"),
        device_key_index=2,
    )
```

`scm_wip_diff/atx_parser.py`의 `parse_atx_wip_sheet`의 `return ParsedWip(...)` 호출부를 다음으로 교체:
```python
    return ParsedWip(
        column_labels=labels,
        stage_groups=groups,
        lots=lots,
        rows=rows,
        sheet_name=ws.title,
        value_number_format=VALUE_NUMBER_FORMAT,
        key_labels=("웨이퍼랏", "디바이스", "컨트롤랏"),
        device_key_index=1,
    )
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_parser.py tests/test_atx_parser.py -v`
Expected: PASS (21 passed — 11 from test_parser.py + 10 from test_atx_parser.py)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/parser.py scm_wip_diff/atx_parser.py tests/test_parser.py tests/test_atx_parser.py
git commit -m "feat: add device_key_index to ParsedWip for format-safe device grouping"
```

---

### Task 2: comparator.py — 디바이스별 단계 수량 집계

**Files:**
- Modify: `scm_wip_diff/comparator.py`
- Test: `tests/test_comparator.py`

**배경:** 어제/오늘 각 랏의 키에서 `device_key_index`로 디바이스명을 뽑아, 존재하는 단계(전공정/후공정/완료 중 실제 있는 것)별로 수량을 합산한다. 어제 또는 오늘 파일에 존재하는 모든 디바이스를 포함하고, 디바이스명 가나다순으로 정렬한다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_comparator.py`에 추가:
```python
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
```

`tests/test_comparator.py`는 현재 맨 위 4줄이 다음과 같다:
```python
from scm_wip_diff.parser import ParsedWip, parse_wip_sheet
from scm_wip_diff.comparator import compare_stage_summary
from scm_wip_diff.comparator import compare_lots
from scm_wip_diff.comparator import check_lot_overlap
```
이 4줄을 다음으로 교체한다 (전부 새로 작성, 기존 4줄을 대체):
```python
import pytest

from scm_wip_diff.parser import ParsedWip, parse_wip_sheet
from scm_wip_diff.atx_parser import parse_atx_wip_sheet
from scm_wip_diff.comparator import compare_stage_summary
from scm_wip_diff.comparator import compare_stage_summary_by_device
from scm_wip_diff.comparator import compare_lots
from scm_wip_diff.comparator import check_lot_overlap
```
그 아래 `FIXTURE_260721`/`FIXTURE_260722` 상수 2줄은 그대로 두고, 바로 다음 줄에 ATX 픽스처 상수 2개를 추가한다:
```python
FIXTURE_260723_ATX = "tests/fixtures/260723 ATX WIP.xlsx"
FIXTURE_260724_ATX = "tests/fixtures/260724 ATX WIP.xlsx"
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_comparator.py -k by_device -v`
Expected: FAIL with `ImportError: cannot import name 'compare_stage_summary_by_device'`

- [ ] **Step 3: compare_stage_summary_by_device 구현**

`scm_wip_diff/comparator.py` 파일 끝에 추가:
```python
def compare_stage_summary_by_device(yesterday, today):
    device_idx = yesterday.device_key_index
    present_stages = [stage for stage in STAGE_ORDER if stage in yesterday.stage_groups]

    y_by_device = {}
    for key, values in yesterday.lots.items():
        device = key[device_idx]
        totals = y_by_device.setdefault(device, {})
        for stage in present_stages:
            cols = yesterday.stage_groups.get(stage, [])
            totals[stage] = totals.get(stage, 0) + sum(values.get(c, 0) for c in cols)

    t_by_device = {}
    for key, values in today.lots.items():
        device = key[device_idx]
        totals = t_by_device.setdefault(device, {})
        for stage in present_stages:
            cols = today.stage_groups.get(stage, [])
            totals[stage] = totals.get(stage, 0) + sum(values.get(c, 0) for c in cols)

    devices = sorted(set(y_by_device) | set(t_by_device))
    summary = {}
    for device in devices:
        summary[device] = {}
        for stage in present_stages:
            y_total = y_by_device.get(device, {}).get(stage, 0)
            t_total = t_by_device.get(device, {}).get(stage, 0)
            summary[device][stage] = {
                "yesterday": y_total,
                "today": t_total,
                "delta": t_total - y_total,
            }
    return summary
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_comparator.py -v`
Expected: PASS (17 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/comparator.py tests/test_comparator.py
git commit -m "feat: aggregate stage-level quantities per device"
```

---

### Task 3: report.py — `derive_output_paths`가 저장 폴더를 인자로 받도록 변경

**Files:**
- Modify: `scm_wip_diff/report.py`
- Test: `tests/test_report.py`

**배경:** 지금까지 변동리포트/하이라이트 파일은 항상 "오늘 파일"과 같은 폴더에 저장됐다. 사용자가 원하는 폴더에 저장할 수 있도록, 폴더 부분을 명시적으로 인자로 받는다. 파일명(날짜/회사 토큰 추출, 접미사)은 그대로 `today_path`의 파일명에서 뽑는다.

- [ ] **Step 1: 실패하는 테스트로 기존 테스트 갱신**

`tests/test_report.py`의 4개 `test_derive_output_paths_*` 함수를 다음으로 교체 (각각 `derive_output_paths` 호출에 `output_folder` 인자를 명시적으로 추가):
```python
def test_derive_output_paths_uses_date_prefix_and_given_folder():
    today_path = os.path.join("C:", os.sep, "data", "260722 GTK WIP.xlsx")
    output_folder = os.path.join("D:", os.sep, "reports")

    report_path, highlighted_path = derive_output_paths(today_path, output_folder)

    assert report_path == os.path.join("D:", os.sep, "reports", "260722_GTK_WIP_변동리포트.xlsx")
    assert highlighted_path == os.path.join("D:", os.sep, "reports", "260722 GTK WIP_변동표시.xlsx")


def test_derive_output_paths_works_for_atx_filename():
    today_path = os.path.join("C:", os.sep, "data", "260723 ATX WIP.xlsx")
    output_folder = os.path.join("C:", os.sep, "data")

    report_path, highlighted_path = derive_output_paths(today_path, output_folder)

    assert report_path == os.path.join("C:", os.sep, "data", "260723_ATX_WIP_변동리포트.xlsx")
    assert highlighted_path == os.path.join("C:", os.sep, "data", "260723 ATX WIP_변동표시.xlsx")


def test_derive_output_paths_falls_back_when_pattern_does_not_match():
    today_path = os.path.join("C:", os.sep, "data", "random_file.xlsx")
    output_folder = os.path.join("C:", os.sep, "data")

    report_path, highlighted_path = derive_output_paths(today_path, output_folder)

    assert report_path == os.path.join("C:", os.sep, "data", "random_file_WIP_변동리포트.xlsx")
    assert highlighted_path == os.path.join("C:", os.sep, "data", "random_file_변동표시.xlsx")


def test_derive_output_paths_falls_back_when_no_company_token_present():
    # Known accepted trade-off: a date+WIP filename with no company token in
    # between (no real GTK/ATX file looks like this) doesn't match the
    # date+company+WIP regex, so it falls through to the generic fallback,
    # producing a cosmetically odd but harmless "WIP_WIP" filename.
    today_path = os.path.join("C:", os.sep, "data", "260723 WIP.xlsx")
    output_folder = os.path.join("C:", os.sep, "data")

    report_path, highlighted_path = derive_output_paths(today_path, output_folder)

    assert report_path == os.path.join("C:", os.sep, "data", "260723 WIP_WIP_변동리포트.xlsx")
    assert highlighted_path == os.path.join("C:", os.sep, "data", "260723 WIP_변동표시.xlsx")
```

Add one more new test right after these, proving the output folder is independent from the input file's folder:
```python
def test_derive_output_paths_uses_output_folder_not_today_paths_folder():
    today_path = os.path.join("C:", os.sep, "somewhere", "260722 GTK WIP.xlsx")
    output_folder = os.path.join("E:", os.sep, "다른폴더")

    report_path, highlighted_path = derive_output_paths(today_path, output_folder)

    assert report_path == os.path.join("E:", os.sep, "다른폴더", "260722_GTK_WIP_변동리포트.xlsx")
    assert highlighted_path == os.path.join("E:", os.sep, "다른폴더", "260722 GTK WIP_변동표시.xlsx")
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_report.py -k derive_output_paths -v`
Expected: FAIL with `TypeError: derive_output_paths() missing 1 required positional argument: 'output_folder'`

- [ ] **Step 3: derive_output_paths 수정**

`scm_wip_diff/report.py`의 `derive_output_paths`를 다음으로 교체:
```python
def derive_output_paths(today_path, output_folder):
    basename = os.path.basename(today_path)
    name, ext = os.path.splitext(basename)
    match = re.match(r"^(\d{6})\s+(\S+)\s+WIP", basename)
    if match:
        date_prefix, company = match.group(1), match.group(2)
        report_name = f"{date_prefix}_{company}_WIP_변동리포트.xlsx"
    else:
        report_name = f"{name}_WIP_변동리포트.xlsx"
    report_path = os.path.join(output_folder, report_name)
    highlighted_path = os.path.join(output_folder, f"{name}_변동표시{ext}")
    return report_path, highlighted_path
```
(`folder = os.path.dirname(today_path)` 줄을 삭제하고, 이후 `os.path.join(folder, ...)`이었던 두 곳을 `os.path.join(output_folder, ...)`로 바꾼다.)

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_report.py -v`
Expected: PASS (12 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/report.py tests/test_report.py
git commit -m "feat: let derive_output_paths write to a caller-specified folder"
```

---

### Task 4: report.py — `build_variance_report`에 디바이스별 단계 수량 섹션 추가

**Files:**
- Modify: `scm_wip_diff/report.py`
- Test: `tests/test_report.py`

**배경:** "요약" 시트의 기존 내용(단계별 합계, 변경/신규/삭제 랏 수) 아래에 "[디바이스별 단계 수량]" 섹션을 추가한다. 반복 대상 단계 목록은 `device_summary`가 아니라 이미 계산되어 있는 `stage_summary`의 키(`STAGE_ORDER` 순서)를 사용해서, `device_summary`가 비어 있어도 헤더 행이 항상 올바르게 만들어지도록 한다.

- [ ] **Step 1: 실패하는 테스트로 기존 테스트 데이터 갱신 + 신규 테스트 작성**

`tests/test_report.py`의 `LOT_DIFF` 딕셔너리 아래에 다음 상수를 추가:
```python
DEVICE_SUMMARY = {
    "d1": {
        "전공정": {"yesterday": 100, "today": 0, "delta": -100},
        "후공정": {"yesterday": 0, "today": 100, "delta": 100},
        "완료": {"yesterday": 0, "today": 0, "delta": 0},
    },
    "d2": {
        "전공정": {"yesterday": 0, "today": 0, "delta": 0},
        "후공정": {"yesterday": 10, "today": 10, "delta": 0},
        "완료": {"yesterday": 0, "today": 0, "delta": 0},
    },
}
```

기존 `test_build_variance_report_writes_expected_sheets`, `test_build_variance_report_applies_readability_formatting`, `test_build_variance_report_uses_atx_key_labels_not_gtk_hardcoded_headers`, `test_build_variance_report_rejects_key_labels_of_wrong_length` 4개 함수의 `build_variance_report(...)` 호출부 전부에 7번째 인자로 `DEVICE_SUMMARY`를 추가한다. 예:
```python
    build_variance_report(
        STAGE_SUMMARY, LOT_DIFF, str(output_path), COLUMN_LABELS, "#,##0", ("MO", "랏번호", "디바이스"), DEVICE_SUMMARY
    )
```
(`test_build_variance_report_uses_atx_key_labels_not_gtk_hardcoded_headers`와 `test_build_variance_report_rejects_key_labels_of_wrong_length`도 동일하게 `DEVICE_SUMMARY`를 마지막 인자로 추가 — 후자는 어차피 `ValueError`가 먼저 발생하므로 값 자체는 테스트 결과에 영향 없음.)

`test_build_variance_report_writes_expected_sheets`에 다음 assertion을 추가 (함수 맨 끝에):
```python
    assert summary_ws["A10"].value == "[디바이스별 단계 수량]"
    assert summary_ws["A11"].value == "디바이스"
    assert summary_ws["B11"].value == "전공정 어제"
    assert summary_ws["C11"].value == "전공정 오늘"
    assert summary_ws["D11"].value == "전공정 증감"
    assert summary_ws["E11"].value == "후공정 어제"
    assert summary_ws["F11"].value == "후공정 오늘"
    assert summary_ws["G11"].value == "후공정 증감"
    assert summary_ws["H11"].value == "완료 어제"
    assert summary_ws["I11"].value == "완료 오늘"
    assert summary_ws["J11"].value == "완료 증감"
    assert summary_ws["A12"].value == "d1"
    assert summary_ws["B12"].value == 100
    assert summary_ws["C12"].value == 0
    assert summary_ws["D12"].value == -100
    assert summary_ws["E12"].value == 0
    assert summary_ws["F12"].value == 100
    assert summary_ws["G12"].value == 100
    assert summary_ws["A13"].value == "d2"
    assert summary_ws["F13"].value == 10
```

`test_build_variance_report_applies_readability_formatting`에 다음 assertion을 추가 (함수 맨 끝에):
```python
    assert summary_ws["B12"].number_format == "#,##0"
    assert summary_ws["G12"].number_format == "#,##0"
    assert summary_ws["I12"].number_format == "#,##0"
```

새 테스트를 추가해, 단계가 2개뿐인 경우(ATX처럼 "완료"가 없는 경우) 디바이스 섹션 열이 올바르게 줄어드는지 확인:
```python
def test_build_variance_report_device_section_omits_absent_stages(tmp_path):
    output_path = tmp_path / "report.xlsx"
    two_stage_summary = {
        "전공정": {"yesterday": 100, "today": 50, "delta": -50},
        "후공정": {"yesterday": 10, "today": 110, "delta": 100},
    }
    two_stage_device_summary = {
        "d1": {
            "전공정": {"yesterday": 100, "today": 50, "delta": -50},
            "후공정": {"yesterday": 10, "today": 110, "delta": 100},
        },
    }

    build_variance_report(
        two_stage_summary,
        LOT_DIFF,
        str(output_path),
        COLUMN_LABELS,
        "#,##0.00",
        ("웨이퍼랏", "디바이스", "컨트롤랏"),
        two_stage_device_summary,
    )

    wb = openpyxl.load_workbook(str(output_path))
    summary_ws = wb["요약"]
    assert summary_ws["A9"].value == "[디바이스별 단계 수량]"
    assert summary_ws["A10"].value == "디바이스"
    assert summary_ws["B10"].value == "전공정 어제"
    assert summary_ws["E10"].value == "후공정 어제"
    assert summary_ws["G10"].value == "후공정 증감"
    assert summary_ws["A11"].value == "d1"
    assert summary_ws["E11"].value == 10
    assert summary_ws["G11"].value == 100
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_report.py -k build_variance_report -v`
Expected: FAIL with `TypeError: build_variance_report() missing 1 required positional argument: 'device_summary'`

- [ ] **Step 3: build_variance_report 수정**

`scm_wip_diff/report.py`의 `build_variance_report`를 다음으로 교체:
```python
def build_variance_report(
    stage_summary, lot_diff, output_path, column_labels, value_number_format, key_labels, device_summary
):
    if len(key_labels) != 3:
        raise ValueError(f"key_labels must have exactly 3 elements, got {len(key_labels)}: {key_labels}")

    wb = openpyxl.Workbook()

    summary_ws = wb.active
    summary_ws.title = "요약"
    summary_ws.append(["단계", "어제", "오늘", "증감"])
    for stage in STAGE_ORDER:
        if stage in stage_summary:
            s = stage_summary[stage]
            summary_ws.append([stage, s["yesterday"], s["today"], s["delta"]])
    summary_ws.append([])
    summary_ws.append(["변경된 랏 수", len(lot_diff["changed_lots"])])
    summary_ws.append(["신규 랏 수", len(lot_diff["new_lots"])])
    summary_ws.append(["삭제된 랏 수", len(lot_diff["removed_lots"])])

    present_stages = [stage for stage in STAGE_ORDER if stage in stage_summary]
    device_number_format_columns = []
    if present_stages:
        summary_ws.append([])
        summary_ws.append(["[디바이스별 단계 수량]"])
        device_header = ["디바이스"]
        for stage in present_stages:
            device_header += [f"{stage} 어제", f"{stage} 오늘", f"{stage} 증감"]
        summary_ws.append(device_header)
        for device in device_summary:
            row = [device]
            for stage in present_stages:
                s = device_summary[device].get(stage, {"yesterday": 0, "today": 0, "delta": 0})
                row += [s["yesterday"], s["today"], s["delta"]]
            summary_ws.append(row)
        device_number_format_columns = list(range(5, 2 + len(present_stages) * 3))

    _apply_readability_formatting(
        summary_ws,
        number_format_columns=[2, 3, 4] + device_number_format_columns,
        number_format=value_number_format,
    )

    changed_ws = wb.create_sheet("변동랏")
    changed_ws.append(list(key_labels) + ["변경컬럼", "어제값", "오늘값"])
    for lot in lot_diff["changed_lots"]:
        key1, key2, key3, _ = lot["key"]
        for change in lot["changes"]:
            changed_ws.append([key1, key2, key3, change["label"], change["before"], change["after"]])
    _apply_readability_formatting(changed_ws, number_format_columns=[5, 6], number_format=value_number_format)

    process_columns = lot_diff["process_columns"]
    process_headers = _process_column_headers(column_labels, process_columns)
    process_number_format_columns = list(range(4, 4 + len(process_headers)))

    new_ws = wb.create_sheet("신규랏")
    new_ws.append(list(key_labels) + process_headers)
    for lot in lot_diff["new_lots"]:
        key1, key2, key3, _ = lot["key"]
        values = lot["values"]
        new_ws.append([key1, key2, key3] + [values.get(col, 0) for col in process_columns])
    _apply_readability_formatting(new_ws, number_format_columns=process_number_format_columns, number_format=value_number_format)

    removed_ws = wb.create_sheet("삭제랏")
    removed_ws.append(list(key_labels) + process_headers)
    for lot in lot_diff["removed_lots"]:
        key1, key2, key3, _ = lot["key"]
        values = lot["values"]
        removed_ws.append([key1, key2, key3] + [values.get(col, 0) for col in process_columns])
    _apply_readability_formatting(removed_ws, number_format_columns=process_number_format_columns, number_format=value_number_format)

    wb.save(output_path)
```

(이 태스크는 `changed_ws`/`new_ws`/`removed_ws` 관련 코드는 건드리지 않는다 — `summary_ws` 빌드 부분만 변경되고, 나머지는 그대로 복사되어 있는 것을 확인하는 차원에서 전체 함수를 다시 옮겨 적었다.)

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_report.py -v`
Expected: PASS (13 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/report.py tests/test_report.py
git commit -m "feat: add per-device stage quantity section to the variance report summary sheet"
```

---

### Task 5: report.py — ATX 통합 테스트에 output_folder/device_summary 반영

**Files:**
- Modify: `tests/test_report.py`

**배경:** `test_atx_end_to_end_report_and_highlight`는 실제 파이프라인을 그대로 재현하는 통합 테스트다. Task 3, 4에서 `derive_output_paths`와 `build_variance_report`의 시그니처가 바뀌었으므로, 이 테스트도 실제 흐름과 일치하도록 갱신한다.

- [ ] **Step 1: 실패하는 테스트로 기존 통합 테스트 갱신**

`tests/test_report.py`의 `test_atx_end_to_end_report_and_highlight` 함수를 다음으로 교체:
```python
def test_atx_end_to_end_report_and_highlight(tmp_path):
    yesterday = parse_atx_wip_sheet(FIXTURE_260723_ATX)
    today = parse_atx_wip_sheet(FIXTURE_260724_ATX)

    stage_summary = compare_stage_summary(yesterday, today)
    device_summary = compare_stage_summary_by_device(yesterday, today)
    lot_diff = compare_lots(yesterday, today)

    assert set(stage_summary.keys()) == {"전공정", "후공정"}
    assert len(lot_diff["changed_lots"]) == 81
    assert len(lot_diff["new_lots"]) == 25
    assert len(lot_diff["removed_lots"]) == 9
    assert len(device_summary) == 11
    assert list(device_summary.keys())[0] == "GTMP122007"

    report_path, highlighted_path = derive_output_paths(str(FIXTURE_260724_ATX), str(tmp_path))
    build_variance_report(
        stage_summary,
        lot_diff,
        report_path,
        today.column_labels,
        today.value_number_format,
        today.key_labels,
        device_summary,
    )
    report_wb = openpyxl.load_workbook(report_path)
    assert report_wb["요약"]["A2"].value == "전공정"
    assert report_wb["요약"]["A9"].value == "[디바이스별 단계 수량]"
    assert report_wb["요약"]["A10"].value == "디바이스"
    assert report_wb["요약"]["A11"].value == "GTMP122007"
    assert report_wb["변동랏"]["E2"].number_format == "#,##0.00"
    assert report_wb["변동랏"]["A1"].value == "웨이퍼랏"
    assert report_wb["변동랏"]["B1"].value == "디바이스"
    assert report_wb["변동랏"]["C1"].value == "컨트롤랏"
    assert report_wb["신규랏"]["A1"].value == "웨이퍼랏"
    assert report_wb["신규랏"]["B1"].value == "디바이스"
    assert report_wb["신규랏"]["C1"].value == "컨트롤랏"
    # Rows are written in lot_diff["new_lots"] order, so row 2 corresponds to
    # the first new lot: confirm WAFERLOT/DEVICE/CONTROLLOT values land under
    # their correctly-labeled columns (A=웨이퍼랏, B=디바이스, C=컨트롤랏).
    first_new_key = lot_diff["new_lots"][0]["key"]
    new_ws = report_wb["신규랏"]
    assert (new_ws["A2"].value, new_ws["B2"].value, new_ws["C2"].value) == first_new_key[:3]

    today_copy = tmp_path / "260724 ATX WIP.xlsx"
    shutil.copyfile(FIXTURE_260724_ATX, today_copy)
    build_highlighted_today_file(str(today_copy), lot_diff, highlighted_path, today.sheet_name)

    highlighted_wb = openpyxl.load_workbook(highlighted_path)
    ws = highlighted_wb[today.sheet_name]
    first_changed = next(
        lot for lot in lot_diff["changed_lots"]
        if lot["key"] == ("BQ8886001A", "GTMP122007", "TPSJ26N009", 0)
    )
    changed_col = first_changed["changes"][0]["col"]
    assert ws.cell(row=first_changed["row_in_today"], column=changed_col).fill.fgColor.rgb == "FFFF0000"
```

Note: `report_path`/`highlighted_path`는 이제 `derive_output_paths`가 반환한 그대로 사용한다 (이전에는 `tmp_path / "report.xlsx"`처럼 직접 만들었지만, 이번 태스크는 `derive_output_paths`가 실제로 `tmp_path`를 폴더로 써서 올바른 파일명을 만드는지도 함께 검증한다). `highlighted_path`는 `build_highlighted_today_file` 호출 시 그대로 출력 경로로 쓰면 되고, 이 함수는 `today_path`(원본 복사 대상)와 `output_path`(결과 저장 대상)를 별도 인자로 받으므로 `today_copy`(원본 복사본)와 `highlighted_path`(derive_output_paths가 계산한 최종 저장 경로)를 각각 맞는 자리에 넣는다.

`tests/test_report.py` 상단 import에 `compare_stage_summary_by_device`를 추가:
```python
from scm_wip_diff.comparator import compare_lots, compare_stage_summary, compare_stage_summary_by_device
```

- [ ] **Step 2: 테스트 실행**

Run: `python -m pytest tests/test_report.py::test_atx_end_to_end_report_and_highlight -v`
Expected: PASS (Task 1~4가 모두 올바르게 구현되었다면 첫 실행에 통과해야 함. 실패하면 어느 태스크가 잘못됐는지 조사한다 — 숫자를 임의로 맞추지 말 것)

- [ ] **Step 3: 전체 테스트 스위트 실행**

Run: `python -m pytest tests/ -v`
Expected: 모두 PASS

- [ ] **Step 4: 커밋**

```bash
git add tests/test_report.py
git commit -m "test: update ATX end-to-end test for new derive_output_paths/build_variance_report signatures"
```

---

### Task 6: gui.py — 저장 폴더 선택 UI + 디바이스별 미리보기 연동

**Files:**
- Modify: `scm_wip_diff/gui.py`

**배경:** GUI에 "저장 폴더" 선택란을 추가하고, "오늘 파일"을 고를 때마다 그 폴더로 자동 채운다. `run_compare`가 `compare_stage_summary_by_device`를 호출해 얻은 결과를 `build_variance_report`에 전달하고, `derive_output_paths`에도 저장 폴더를 넘긴다. `_show_preview`에 디바이스별 단계 수량 섹션을 추가한다.

이 태스크는 GUI라 자동 테스트 대신 수동/기능적 스모크 테스트로 검증한다 (기존 정책과 동일).

- [ ] **Step 1: gui.py 전체 교체**

`scm_wip_diff/gui.py`:
```python
"""tkinter GUI wiring format detection -> comparator -> report together."""
import os
import tkinter as tk
from tkinter import filedialog, messagebox

from scm_wip_diff.comparator import (
    check_lot_overlap,
    compare_lots,
    compare_stage_summary,
    compare_stage_summary_by_device,
)
from scm_wip_diff.format_detect import parse_wip_file
from scm_wip_diff.parser import ReportFormatError
from scm_wip_diff.report import (
    build_highlighted_today_file,
    build_variance_report,
    derive_output_paths,
)


class App:
    def __init__(self, root):
        self.root = root
        self.yesterday_path = tk.StringVar()
        self.today_path = tk.StringVar()
        self.output_folder = tk.StringVar()

        tk.Label(root, text="어제 파일:").grid(row=0, column=0, sticky="w")
        tk.Entry(root, textvariable=self.yesterday_path, width=60).grid(row=0, column=1)
        tk.Button(root, text="찾아보기", command=self.choose_yesterday).grid(row=0, column=2)

        tk.Label(root, text="오늘 파일:").grid(row=1, column=0, sticky="w")
        tk.Entry(root, textvariable=self.today_path, width=60).grid(row=1, column=1)
        tk.Button(root, text="찾아보기", command=self.choose_today).grid(row=1, column=2)

        tk.Label(root, text="저장 폴더:").grid(row=2, column=0, sticky="w")
        tk.Entry(root, textvariable=self.output_folder, width=60).grid(row=2, column=1)
        tk.Button(root, text="찾아보기", command=self.choose_output_folder).grid(row=2, column=2)

        tk.Button(root, text="비교 실행", command=self.run_compare).grid(row=3, column=1)

        self.result_text = tk.Text(root, width=100, height=30)
        self.result_text.grid(row=4, column=0, columnspan=3)

    def choose_yesterday(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.yesterday_path.set(path)

    def choose_today(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.today_path.set(path)
            self.output_folder.set(os.path.dirname(path))

    def choose_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder.set(folder)

    def run_compare(self):
        y_path = self.yesterday_path.get()
        t_path = self.today_path.get()
        if not y_path or not t_path:
            messagebox.showerror("입력 필요", "어제 파일과 오늘 파일을 모두 선택하세요")
            return

        output_folder = self.output_folder.get() or os.path.dirname(t_path)

        try:
            try:
                y_format, yesterday = parse_wip_file(y_path)
            except ReportFormatError as e:
                messagebox.showerror("파일 형식 오류", f"어제 파일 오류 ({y_path}):\n{e}")
                return

            try:
                t_format, today = parse_wip_file(t_path)
            except ReportFormatError as e:
                messagebox.showerror("파일 형식 오류", f"오늘 파일 오류 ({t_path}):\n{e}")
                return

            if y_format != t_format:
                messagebox.showerror(
                    "형식 불일치",
                    f"어제 파일은 {y_format} 형식, 오늘 파일은 {t_format} 형식입니다.\n"
                    "같은 회사 파일끼리 비교해주세요.",
                )
                return

            overlap_warning = check_lot_overlap(yesterday, today)
            if overlap_warning:
                messagebox.showwarning("확인 필요", overlap_warning)

            stage_summary = compare_stage_summary(yesterday, today)
            device_summary = compare_stage_summary_by_device(yesterday, today)
            lot_diff = compare_lots(yesterday, today)

            report_path, highlighted_path = derive_output_paths(t_path, output_folder)

            try:
                build_variance_report(
                    stage_summary,
                    lot_diff,
                    report_path,
                    today.column_labels,
                    today.value_number_format,
                    today.key_labels,
                    device_summary,
                )
            except PermissionError:
                messagebox.showerror(
                    "저장 실패",
                    f"변동 리포트 파일이 열려있어 저장할 수 없습니다:\n{report_path}",
                )
                return

            try:
                build_highlighted_today_file(t_path, lot_diff, highlighted_path, today.sheet_name)
            except PermissionError:
                messagebox.showerror(
                    "저장 실패",
                    "변동 리포트는 저장되었지만, 하이라이트 파일이 열려있어 저장할 수 없습니다.\n\n"
                    f"저장됨: {report_path}\n실패: {highlighted_path}",
                )
                return

            self._show_preview(stage_summary, device_summary, lot_diff)
            messagebox.showinfo("완료", f"저장 완료:\n{report_path}\n{highlighted_path}")
        except Exception as e:
            messagebox.showerror("예상치 못한 오류", str(e))

    def _show_preview(self, stage_summary, device_summary, lot_diff):
        self.result_text.delete("1.0", tk.END)
        lines = ["[단계별 합계]"]
        for stage in stage_summary:
            s = stage_summary[stage]
            lines.append(f"{stage}: {s['yesterday']:,} -> {s['today']:,} ({s['delta']:+,})")
        lines.append("")
        lines.append(f"변경된 랏 수: {len(lot_diff['changed_lots'])}")
        lines.append(f"신규 랏 수: {len(lot_diff['new_lots'])}")
        lines.append(f"삭제된 랏 수: {len(lot_diff['removed_lots'])}")
        lines.append("")
        lines.append("[디바이스별 단계 수량]")
        for device, stages in device_summary.items():
            parts = [
                f"{stage} {s['yesterday']:,}->{s['today']:,}({s['delta']:+,})"
                for stage, s in stages.items()
            ]
            lines.append(f"{device}: {', '.join(parts)}")
        self.result_text.insert("1.0", "\n".join(lines))
```

- [ ] **Step 2: 기능적 스모크 테스트 (GTK)**

```bash
python -c "
from scm_wip_diff.format_detect import parse_wip_file
from scm_wip_diff.comparator import compare_stage_summary, compare_stage_summary_by_device, compare_lots
from scm_wip_diff.report import build_variance_report, build_highlighted_today_file, derive_output_paths
import tempfile, os, shutil

y_fmt, yesterday = parse_wip_file('tests/fixtures/260721 GTK WIP.xlsx')
t_fmt, today = parse_wip_file('tests/fixtures/260722 GTK WIP.xlsx')

stage_summary = compare_stage_summary(yesterday, today)
device_summary = compare_stage_summary_by_device(yesterday, today)
lot_diff = compare_lots(yesterday, today)
assert len(device_summary) == 34
assert list(device_summary.keys())[0] == 'TMP1200D'

with tempfile.TemporaryDirectory() as d:
    report_path, highlighted_path = derive_output_paths('tests/fixtures/260722 GTK WIP.xlsx', d)
    build_variance_report(stage_summary, lot_diff, report_path, today.column_labels, today.value_number_format, today.key_labels, device_summary)
    today_copy = os.path.join(d, 'today.xlsx')
    shutil.copyfile('tests/fixtures/260722 GTK WIP.xlsx', today_copy)
    build_highlighted_today_file(today_copy, lot_diff, highlighted_path, today.sheet_name)
    assert os.path.exists(report_path)
    assert os.path.exists(highlighted_path)
    print('GTK end-to-end with device summary + custom output folder OK')
"
```
Expected: `GTK end-to-end with device summary + custom output folder OK` 출력, 예외 없음

- [ ] **Step 3: 기능적 스모크 테스트 (ATX)**

```bash
python -c "
from scm_wip_diff.format_detect import parse_wip_file
from scm_wip_diff.comparator import compare_stage_summary, compare_stage_summary_by_device, compare_lots
from scm_wip_diff.report import build_variance_report, build_highlighted_today_file, derive_output_paths
import tempfile, os, shutil

y_fmt, yesterday = parse_wip_file('tests/fixtures/260723 ATX WIP.xlsx')
t_fmt, today = parse_wip_file('tests/fixtures/260724 ATX WIP.xlsx')

stage_summary = compare_stage_summary(yesterday, today)
device_summary = compare_stage_summary_by_device(yesterday, today)
lot_diff = compare_lots(yesterday, today)
assert len(device_summary) == 11
assert '완료' not in device_summary['GTMP122007']

with tempfile.TemporaryDirectory() as d:
    report_path, highlighted_path = derive_output_paths('tests/fixtures/260724 ATX WIP.xlsx', d)
    build_variance_report(stage_summary, lot_diff, report_path, today.column_labels, today.value_number_format, today.key_labels, device_summary)
    today_copy = os.path.join(d, 'today.xlsx')
    shutil.copyfile('tests/fixtures/260724 ATX WIP.xlsx', today_copy)
    build_highlighted_today_file(today_copy, lot_diff, highlighted_path, today.sheet_name)
    assert os.path.exists(report_path)
    assert os.path.exists(highlighted_path)
    print('ATX end-to-end with device summary + custom output folder OK')
"
```
Expected: `ATX end-to-end with device summary + custom output folder OK` 출력, 예외 없음

- [ ] **Step 4: 실제 GUI 창 육안 확인 (가능한 환경이면)**

```bash
python -m scm_wip_diff.main
```
확인 사항:
1. "저장 폴더" 항목이 보이는지
2. "오늘 파일"을 고르면 "저장 폴더"가 자동으로 채워지는지
3. "저장 폴더"의 "찾아보기"로 다른 폴더를 고르면 그 값으로 바뀌는지, 이후 "오늘 파일"을 다시 고르기 전까지 유지되는지
4. "비교 실행" 후 미리보기 텍스트에 "[디바이스별 단계 수량]" 섹션이 보이고, 지정한 저장 폴더에 실제로 두 파일이 생성되는지

디스플레이가 없는 환경이면 이 단계는 생략하고 Step 2~3의 결과로 충분하다고 보고한다.

- [ ] **Step 5: 전체 테스트 스위트 재확인**

Run: `python -m pytest tests/ -v`
Expected: 모두 PASS (gui.py는 자동 테스트 대상 아님, 변경 없음)

- [ ] **Step 6: 커밋**

```bash
git add scm_wip_diff/gui.py
git commit -m "feat: add output folder picker and per-device stage breakdown to the GUI"
```

---

## 실행 시 참고사항

- 저장 폴더 선택값은 세션 내에서만 유지되며, 프로그램을 다시 시작하면 초기화된다 (설계서 "범위 밖" 참조).
- 디바이스별 단계 수량은 어제 또는 오늘 파일에 존재하는 모든 디바이스를 디바이스명 가나다순으로 보여준다 (변동 없는 디바이스도 포함).
- `column_labels`/`value_number_format`/`key_labels`/`device_summary`를 하나의 객체로 묶는 리팩터링은 이번 범위에 포함하지 않는다 (설계서 "범위 밖" 참조, 별도 기술 부채로 관리).
