# ATX WIP 지원 확장 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 기존 GTK WIP 일일 비교 도구가 ATX(다른 협력사) WIP 파일의 `KSWIPAY` 시트도 자동으로 인식해 동일한 방식(변동리포트 엑셀 + 하이라이트된 오늘 파일)으로 비교할 수 있도록 확장한다.

**Architecture:** `atx_parser.py`(ATX 전용 파서, 헤더 텍스트 기반 그룹 판별)와 `format_detect.py`(GTK/ATX 자동 판별 + 파서 dispatch)를 새로 추가한다. `comparator.py`/`report.py`는 그동안 GTK 전용 상수(`PROCESS_COL_START`/`PROCESS_COL_END`, `"#,##0"` 고정 포맷, `wb.sheetnames[0]` 가정)에 암묵적으로 의존하고 있었음이 재검토 과정에서 드러나, 이번에 `ParsedWip` 구조만으로 동작하도록 일반화한다.

**Tech Stack:** Python 3.14, openpyxl, tkinter, pytest

**설계서:** `docs/superpowers/specs/2026-07-28-atx-wip-support-design.md`

---

## 파일 구조

```
scm_wip_diff/
├── parser.py          # 수정: ParsedWip에 sheet_name/value_number_format 필드 추가
├── atx_parser.py       # 신규: ATX 전용 파서
├── format_detect.py    # 신규: GTK/ATX 자동 판별 + 파서 dispatch
├── comparator.py       # 수정: process_columns 동적 계산, 존재하는 단계만 요약
├── report.py            # 수정: GTK 전용 상수 제거, sheet_name/value_number_format 파라미터화, 파일명 회사 토큰 일반화
├── gui.py               # 수정: format_detect 연동, 형식 불일치 처리
└── main.py              # 변경 없음
tests/
├── fixtures/
│   ├── 260723 ATX WIP.xlsx  # 신규
│   └── 260724 ATX WIP.xlsx  # 신규
├── test_atx_parser.py        # 신규
├── test_format_detect.py     # 신규
├── test_parser.py             # 수정 (필드 추가 테스트)
├── test_comparator.py         # 수정 (process_columns, 단계 생략 테스트)
└── test_report.py              # 수정 (시그니처 변경 반영 + ATX 통합 테스트)
```

## 핵심 데이터 구조 변경

```python
# parser.ParsedWip (필드 2개 추가, 기본값 있어 기존 호출부는 그대로 동작)
@dataclass
class ParsedWip:
    column_labels: dict
    stage_groups: dict
    lots: dict
    rows: dict
    sheet_name: str = ""
    value_number_format: str = "#,##0"

# comparator.compare_lots 반환값에 키 1개 추가
{
  "changed_lots": [...],
  "new_lots": [...],
  "removed_lots": [...],
  "process_columns": [9, 10, ..., 30],  # 신규: stage_groups 합집합에서 동적 계산
}

# format_detect.parse_wip_file(path) 반환값
("GTK" | "ATX", ParsedWip)
```

---

### Task 1: ATX 테스트 픽스처 추가

**Files:**
- Create: `tests/fixtures/260723 ATX WIP.xlsx` (프로젝트 루트 파일 복사)
- Create: `tests/fixtures/260724 ATX WIP.xlsx` (프로젝트 루트 파일 복사)

- [ ] **Step 1: 픽스처 복사**

```bash
cp "260723 ATX WIP.xlsx" "tests/fixtures/260723 ATX WIP.xlsx"
cp "260724 ATX WIP.xlsx" "tests/fixtures/260724 ATX WIP.xlsx"
```

- [ ] **Step 2: 커밋**

```bash
git add tests/fixtures/"260723 ATX WIP.xlsx" tests/fixtures/"260724 ATX WIP.xlsx"
git commit -m "test: add ATX WIP fixture files"
```

---

### Task 2: parser.py - ParsedWip에 sheet_name/value_number_format 필드 추가

**Files:**
- Modify: `scm_wip_diff/parser.py`
- Test: `tests/test_parser.py`

**배경:** `build_highlighted_today_file`이 지금까지 `wb.sheetnames[0]`(첫 번째 시트)을 무조건 사용해왔는데, ATX 파일은 대상 시트 이름이 날짜마다 바뀐다(`KSWIPAY (PKG)` vs `KSWIPAY`). 어떤 시트에 하이라이트를 적용해야 하는지 `ParsedWip`이 직접 들고 있게 한다. 또한 GTK(정수)와 ATX(소수)의 표시 형식이 다르므로 `value_number_format`도 `ParsedWip`에 담는다.

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/test_parser.py`에 추가:
```python
def test_parse_wip_sheet_populates_sheet_name_and_number_format():
    parsed = parse_wip_sheet(FIXTURE_260721)

    assert parsed.sheet_name == "gtk3387"
    assert parsed.value_number_format == "#,##0"
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_parser.py::test_parse_wip_sheet_populates_sheet_name_and_number_format -v`
Expected: FAIL with `AttributeError: 'ParsedWip' object has no attribute 'sheet_name'`

- [ ] **Step 3: ParsedWip 필드 추가 및 parse_wip_sheet에서 채우기**

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
```

`parse_wip_sheet` 함수 맨 끝의 `return` 문을 다음으로 교체:
```python
    return ParsedWip(
        column_labels=labels,
        stage_groups=groups,
        lots=lots,
        rows=rows,
        sheet_name=ws.title,
        value_number_format="#,##0",
    )
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_parser.py -v`
Expected: PASS (9 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/parser.py tests/test_parser.py
git commit -m "feat: add sheet_name and value_number_format to ParsedWip"
```

---

### Task 3: atx_parser.py - KSWIPAY 시트 접두사 매칭

**Files:**
- Create: `scm_wip_diff/atx_parser.py`
- Test: `tests/test_atx_parser.py`

**배경:** 실제 파일을 확인해보니 대상 시트 이름이 `260723` 파일은 `KSWIPAY (PKG)`, `260724` 파일은 `KSWIPAY`로 날짜마다 다르다. 정확한 이름이 아니라 `KSWIPAY`로 **시작하는** 시트를 찾아야 한다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_atx_parser.py`:
```python
import openpyxl
import pytest

from scm_wip_diff.atx_parser import find_atx_sheet
from scm_wip_diff.parser import ReportFormatError

FIXTURE_260723 = "tests/fixtures/260723 ATX WIP.xlsx"
FIXTURE_260724 = "tests/fixtures/260724 ATX WIP.xlsx"


def test_find_atx_sheet_matches_by_prefix_regardless_of_suffix():
    wb_723 = openpyxl.load_workbook(FIXTURE_260723, data_only=True)
    wb_724 = openpyxl.load_workbook(FIXTURE_260724, data_only=True)

    assert find_atx_sheet(wb_723).title == "KSWIPAY (PKG)"
    assert find_atx_sheet(wb_724).title == "KSWIPAY"


def test_find_atx_sheet_raises_when_no_matching_sheet():
    wb = openpyxl.Workbook()

    with pytest.raises(ReportFormatError):
        find_atx_sheet(wb)
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_atx_parser.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'scm_wip_diff.atx_parser'`

- [ ] **Step 3: atx_parser.py에 find_atx_sheet 구현**

`scm_wip_diff/atx_parser.py`:
```python
"""Parse ATX WIP Excel reports (KSWIPAY sheet) into structured lot-level data."""

from scm_wip_diff.parser import ParsedWip, ReportFormatError

ATX_SHEET_PREFIX = "KSWIPAY"
HEADER_ROW = 2
DATA_START_ROW = 3
VALUE_NUMBER_FORMAT = "#,##0.00"


def find_atx_sheet(wb):
    for name in wb.sheetnames:
        if name.startswith(ATX_SHEET_PREFIX):
            return wb[name]
    raise ReportFormatError(f"'{ATX_SHEET_PREFIX}'로 시작하는 시트를 찾을 수 없습니다")
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_atx_parser.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/atx_parser.py tests/test_atx_parser.py
git commit -m "feat: locate ATX KSWIPAY sheet by name prefix"
```

---

### Task 4: atx_parser.py - 헤더 텍스트 기반 전공정/후공정 그룹 판별

**Files:**
- Modify: `scm_wip_diff/atx_parser.py`
- Test: `tests/test_atx_parser.py`

**배경:** GTK는 병합 셀로 그룹 경계를 판별했지만, ATX는 `260723`과 `260724` 파일 사이에 후공정 병합 범위 자체가 `X1:AT1` → `X1:AU1`로 흔들리는 것을 확인했다 (BE Total 소계 컬럼이 병합에 포함되기도, 안 되기도 함). 병합 셀 대신 2행의 헤더 텍스트로 판별한다: `UNISSUE` 컬럼부터 `FE Total` 컬럼 직전까지가 전공정, `Ftape` 컬럼부터 `BE Total` 컬럼 직전까지가 후공정.

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/test_atx_parser.py`에 추가:
```python
from scm_wip_diff.atx_parser import get_atx_stage_groups


def test_get_atx_stage_groups_uses_header_text_not_merged_cells():
    wb = openpyxl.load_workbook(FIXTURE_260723, data_only=True)
    ws = find_atx_sheet(wb)

    groups = get_atx_stage_groups(ws)

    assert groups["전공정"] == list(range(13, 23))
    assert groups["후공정"] == list(range(24, 47))


def test_get_atx_stage_groups_matches_across_both_fixtures_despite_merge_drift():
    wb_723 = openpyxl.load_workbook(FIXTURE_260723, data_only=True)
    wb_724 = openpyxl.load_workbook(FIXTURE_260724, data_only=True)

    groups_723 = get_atx_stage_groups(find_atx_sheet(wb_723))
    groups_724 = get_atx_stage_groups(find_atx_sheet(wb_724))

    assert groups_723 == groups_724


def test_get_atx_stage_groups_raises_when_required_label_missing():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.cell(row=2, column=13, value="UNISSUE")
    ws.cell(row=2, column=23, value="FE Total")
    ws.cell(row=2, column=24, value="Ftape")
    # "BE Total" 헤더를 의도적으로 생략해 포맷 변경을 시뮬레이션

    with pytest.raises(ReportFormatError):
        get_atx_stage_groups(ws)
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_atx_parser.py -k get_atx_stage_groups -v`
Expected: FAIL with `ImportError: cannot import name 'get_atx_stage_groups'`

- [ ] **Step 3: get_atx_stage_groups 구현**

`scm_wip_diff/atx_parser.py`에 추가:
```python
REQUIRED_HEADER_LABELS = ["UNISSUE", "FE Total", "Ftape", "BE Total"]


def _find_label_column(ws, label):
    for col in range(1, ws.max_column + 1):
        if str(ws.cell(row=HEADER_ROW, column=col).value or "").strip() == label:
            return col
    return None


def get_atx_stage_groups(ws):
    positions = {label: _find_label_column(ws, label) for label in REQUIRED_HEADER_LABELS}
    missing = [label for label, col in positions.items() if col is None]
    if missing:
        raise ReportFormatError(f"필수 컬럼 헤더를 찾을 수 없습니다: {missing}")

    unissue, fe_total, ftape, be_total = (positions[label] for label in REQUIRED_HEADER_LABELS)
    return {
        "전공정": list(range(unissue, fe_total)),
        "후공정": list(range(ftape, be_total)),
    }
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_atx_parser.py -v`
Expected: PASS (5 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/atx_parser.py tests/test_atx_parser.py
git commit -m "feat: derive ATX stage groups from header text instead of merged cells"
```

---

### Task 5: atx_parser.py - 전체 시트 파싱 (parse_atx_wip_sheet)

**Files:**
- Modify: `scm_wip_diff/atx_parser.py`
- Test: `tests/test_atx_parser.py`

**배경:** `find_atx_sheet` + `get_atx_stage_groups`를 조합해 3행부터 A열(RSOD)이 빌 때까지 데이터를 읽는다. 키는 WAFERLOT(E열)+DEVICE(C열)+CONTROLLOT(F열)이며, 완전히 동일한 키가 여러 행에 나타나는 경우(실측 확인됨: `BQ8873301A`/`GTMP17500D`/`TPNJ28N009`가 13~15행에 3번 등장, 수량만 다름)를 GTK와 동일하게 등장 순번으로 구분한다. 값은 정수가 아닌 소수이므로 `float`로 저장한다.

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/test_atx_parser.py`에 추가:
```python
from scm_wip_diff.atx_parser import parse_atx_wip_sheet


def test_parse_atx_wip_sheet_reads_all_lot_rows():
    parsed = parse_atx_wip_sheet(FIXTURE_260723)

    assert len(parsed.lots) == 125
    assert parsed.sheet_name == "KSWIPAY (PKG)"
    assert parsed.value_number_format == "#,##0.00"


def test_parse_atx_wip_sheet_extracts_process_values_for_first_data_row():
    parsed = parse_atx_wip_sheet(FIXTURE_260723)
    key = ("BQ8886001A", "GTMP122007", "TPSJ26N009", 0)

    assert parsed.rows[key] == 3
    assert parsed.lots[key][13] == 0.0     # UNISSUE
    assert parsed.lots[key][16] == 54.66   # Die_Bond


def test_parse_atx_wip_sheet_disambiguates_duplicate_keys_by_occurrence():
    parsed = parse_atx_wip_sheet(FIXTURE_260723)
    base = ("BQ8873301A", "GTMP17500D", "TPNJ28N009")

    assert parsed.rows[base + (0,)] == 13
    assert parsed.rows[base + (1,)] == 14
    assert parsed.rows[base + (2,)] == 15
    assert parsed.lots[base + (0,)][17] == 22.98   # Wire_Bond
    assert parsed.lots[base + (1,)][17] == 28.61
    assert parsed.lots[base + (2,)][17] == 34.48
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_atx_parser.py -k parse_atx_wip_sheet -v`
Expected: FAIL with `ImportError: cannot import name 'parse_atx_wip_sheet'`

- [ ] **Step 3: parse_atx_wip_sheet 구현**

`scm_wip_diff/atx_parser.py` 상단에 import 추가:
```python
import openpyxl
```

파일 끝에 추가:
```python
def parse_atx_wip_sheet(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = find_atx_sheet(wb)

    groups = get_atx_stage_groups(ws)
    process_cols = sorted({c for cols in groups.values() for c in cols})
    labels = {
        col: str(ws.cell(row=HEADER_ROW, column=col).value or "").strip()
        for col in process_cols
    }

    lots = {}
    rows = {}
    occurrence = {}
    r = DATA_START_ROW
    while ws.cell(row=r, column=1).value is not None:
        device = str(ws.cell(row=r, column=3).value or "").strip()
        wafer_lot = str(ws.cell(row=r, column=5).value or "").strip()
        control_lot = str(ws.cell(row=r, column=6).value or "").strip()
        base_key = (wafer_lot, device, control_lot)
        idx = occurrence.get(base_key, 0)
        occurrence[base_key] = idx + 1
        key = (wafer_lot, device, control_lot, idx)

        values = {}
        for col in process_cols:
            v = ws.cell(row=r, column=col).value
            values[col] = float(v) if isinstance(v, (int, float)) else 0.0

        lots[key] = values
        rows[key] = r
        r += 1

    return ParsedWip(
        column_labels=labels,
        stage_groups=groups,
        lots=lots,
        rows=rows,
        sheet_name=ws.title,
        value_number_format=VALUE_NUMBER_FORMAT,
    )
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_atx_parser.py -v`
Expected: PASS (8 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/atx_parser.py tests/test_atx_parser.py
git commit -m "feat: parse ATX KSWIPAY sheet into ParsedWip"
```

---

### Task 6: format_detect.py - GTK/ATX 자동 판별 및 파서 dispatch

**Files:**
- Create: `scm_wip_diff/format_detect.py`
- Test: `tests/test_format_detect.py`

**배경:** GUI에서 사용자가 고른 파일이 GTK인지 ATX인지 파일 내용(시트 구성)으로 자동 판별한다. 시트 이름이 `KSWIPAY`로 시작하면 ATX, 그렇지 않고 첫 번째 시트에서 GTK의 `find_report_anchor`가 성공하면 GTK로 판정한다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_format_detect.py`:
```python
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
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_format_detect.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'scm_wip_diff.format_detect'`

- [ ] **Step 3: format_detect.py 구현**

`scm_wip_diff/format_detect.py`:
```python
"""Detect whether a WIP Excel file is GTK or ATX format and parse it accordingly."""
import openpyxl

from scm_wip_diff.atx_parser import find_atx_sheet, parse_atx_wip_sheet
from scm_wip_diff.parser import ReportFormatError, find_report_anchor, parse_wip_sheet

GTK = "GTK"
ATX = "ATX"


def detect_format(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    try:
        find_atx_sheet(wb)
        return ATX
    except ReportFormatError:
        pass

    ws = wb[wb.sheetnames[0]]
    try:
        find_report_anchor(ws)
        return GTK
    except ReportFormatError:
        pass

    raise ReportFormatError("GTK 또는 ATX 형식으로 인식할 수 없는 파일입니다")


def parse_wip_file(path):
    fmt = detect_format(path)
    if fmt == ATX:
        return fmt, parse_atx_wip_sheet(path)
    return fmt, parse_wip_sheet(path)
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_format_detect.py -v`
Expected: PASS (4 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/format_detect.py tests/test_format_detect.py
git commit -m "feat: auto-detect GTK vs ATX format and dispatch to the right parser"
```

---

### Task 7: comparator.py - compare_lots의 비교 컬럼을 stage_groups에서 동적 계산

**Files:**
- Modify: `scm_wip_diff/comparator.py`
- Test: `tests/test_comparator.py`

**배경:** 코드를 다시 확인해보니 `compare_lots`가 GTK 전용 상수 `range(PROCESS_COL_START, PROCESS_COL_END + 1)`(9~30)을 하드코딩하고 있었다. ATX는 컬럼 범위가 다르고(13~22, 24~46) 중간에 소계 컬럼(23열)으로 끊겨 있어 이 상수를 그대로 쓸 수 없다. 비교 대상 컬럼을 `yesterday`/`today`의 `stage_groups`에 실제로 포함된 컬럼 인덱스의 합집합에서 동적으로 계산하도록 바꾸고, 이 목록을 반환값에 `process_columns`로 포함시켜 `report.py`가 재사용할 수 있게 한다.

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/test_comparator.py`에 추가:
```python
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
```

`test_compare_lots_matches_real_fixture_counts`에 다음 assertion을 추가:
```python
    assert diff["process_columns"] == list(range(9, 31))
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_comparator.py -k "process_columns or noncontiguous" -v`
Expected: FAIL with `KeyError: 'process_columns'`

- [ ] **Step 3: compare_lots 수정**

`scm_wip_diff/comparator.py` 상단의 import를 다음으로 교체 (parser 상수 import 제거):
```python
"""Compare two ParsedWip snapshots and produce a diff."""

from openpyxl.utils import get_column_letter

STAGE_ORDER = ["전공정", "후공정", "완료"]
```

`compare_lots` 함수를 다음으로 교체:
```python
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
```

`tests/test_comparator.py` 상단에 `ParsedWip` import가 이미 있는지 확인 (`from scm_wip_diff.parser import ParsedWip, parse_wip_sheet` — 이미 있음, 수정 불필요).

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_comparator.py -v`
Expected: PASS (10 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/comparator.py tests/test_comparator.py
git commit -m "feat: derive compare_lots process columns from stage_groups instead of a GTK-only constant"
```

---

### Task 8: comparator.py - compare_stage_summary가 존재하는 단계만 반환하도록 수정

**Files:**
- Modify: `scm_wip_diff/comparator.py`
- Test: `tests/test_comparator.py`

**배경:** `compare_stage_summary`가 `STAGE_ORDER`(전공정/후공정/완료)를 무조건 3개 다 순회해서, ATX처럼 "완료" 그룹이 아예 없는 데이터도 `{"yesterday": 0, "today": 0, "delta": 0}`으로 채워 넣는다. 이러면 GUI/리포트에 실제로 존재하지 않는 단계가 마치 변동 없는 단계인 것처럼 표시되어 오해를 준다. `yesterday.stage_groups`에 실제로 있는 단계만 결과에 포함하도록 고친다 (GTK는 3개 다 있으므로 동작 변화 없음).

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/test_comparator.py`에 추가:
```python
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
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_comparator.py::test_compare_stage_summary_omits_stages_not_present_in_source_data -v`
Expected: FAIL with `AssertionError` (summary에 "완료": {yesterday:0,today:0,delta:0}이 포함되어 `{"전공정","후공정","완료"} != {"전공정","후공정"}`)

- [ ] **Step 3: compare_stage_summary 수정**

`scm_wip_diff/comparator.py`의 `compare_stage_summary`를 다음으로 교체:
```python
def compare_stage_summary(yesterday, today):
    summary = {}
    for stage in STAGE_ORDER:
        if stage not in yesterday.stage_groups:
            continue
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
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_comparator.py -v`
Expected: PASS (11 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/comparator.py tests/test_comparator.py
git commit -m "feat: omit stages absent from source data in compare_stage_summary"
```

---

### Task 9: report.py - GTK 전용 상수 제거하고 process_columns/value_number_format 사용

**Files:**
- Modify: `scm_wip_diff/report.py`
- Test: `tests/test_report.py`

**배경:** `report.py`가 `scm_wip_diff.parser`의 `PROCESS_COL_START`/`PROCESS_COL_END`를 import해서 모듈 레벨 `PROCESS_COLS = range(9, 31)`을 만들고 있었다 (GTK 전용). 또한 숫자 포맷도 `"#,##0"`으로 고정되어 있어 ATX의 소수 값에는 맞지 않는다. `build_variance_report`가 `lot_diff["process_columns"]`(Task 7에서 추가됨)와 새 `value_number_format` 인자를 받아 사용하도록 바꾼다.

- [ ] **Step 1: 실패하는 테스트로 기존 테스트 데이터 갱신**

`tests/test_report.py`의 `LOT_DIFF` 딕셔너리에 `process_columns` 키를 추가:
```python
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
    "process_columns": [9, 10, 12],
}
```

`test_build_variance_report_writes_expected_sheets`와 `test_build_variance_report_applies_readability_formatting`의 `build_variance_report(...)` 호출부에 5번째 인자로 `"#,##0"`을 추가:
```python
    build_variance_report(STAGE_SUMMARY, LOT_DIFF, str(output_path), COLUMN_LABELS, "#,##0")
```
(두 테스트 함수 모두 동일하게 수정)

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_report.py -k build_variance_report -v`
Expected: FAIL with `TypeError: build_variance_report() takes 4 positional arguments but 5 were given`

- [ ] **Step 3: report.py 수정**

`scm_wip_diff/report.py` 상단의 import를 다음으로 교체 (parser 상수 import 제거):
```python
"""Generate the variance report workbook and the highlighted today-file copy."""
import os
import re
import shutil

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

RED_FILL = PatternFill(fill_type="solid", fgColor="FFFF0000")
NEW_LOT_FILL = PatternFill(fill_type="solid", fgColor="FFADD8E6")

STAGE_ORDER = ["전공정", "후공정", "완료"]
```

`PROCESS_COLS = range(PROCESS_COL_START, PROCESS_COL_END + 1)` 줄과 `NUMBER_FORMAT = "#,##0"` 줄을 삭제하고, 나머지 상수는 유지:
```python
HEADER_FONT = Font(bold=True)
MIN_COLUMN_WIDTH = 10
MAX_COLUMN_WIDTH = 40


def _process_column_headers(column_labels, process_columns):
    return [f"{column_labels.get(col, '')}({get_column_letter(col)})" for col in process_columns]


def _apply_readability_formatting(ws, number_format_columns, number_format):
    for cell in ws[1]:
        if cell.value is not None:
            cell.font = HEADER_FONT
    ws.freeze_panes = "A2"

    for col_cells in ws.columns:
        col_letter = col_cells[0].column_letter
        max_len = max((len(str(c.value)) for c in col_cells if c.value is not None), default=0)
        ws.column_dimensions[col_letter].width = min(
            max(max_len + 2, MIN_COLUMN_WIDTH), MAX_COLUMN_WIDTH
        )

    for col in number_format_columns:
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=col).number_format = number_format


def build_variance_report(stage_summary, lot_diff, output_path, column_labels, value_number_format):
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
    _apply_readability_formatting(summary_ws, number_format_columns=[2, 3, 4], number_format=value_number_format)

    changed_ws = wb.create_sheet("변동랏")
    changed_ws.append(["MO", "랏번호", "디바이스", "변경컬럼", "어제값", "오늘값"])
    for lot in lot_diff["changed_lots"]:
        mo, lot_no, device, _ = lot["key"]
        for change in lot["changes"]:
            changed_ws.append([mo, lot_no, device, change["label"], change["before"], change["after"]])
    _apply_readability_formatting(changed_ws, number_format_columns=[5, 6], number_format=value_number_format)

    process_columns = lot_diff["process_columns"]
    process_headers = _process_column_headers(column_labels, process_columns)
    process_number_format_columns = list(range(4, 4 + len(process_headers)))

    new_ws = wb.create_sheet("신규랏")
    new_ws.append(["MO", "랏번호", "디바이스"] + process_headers)
    for lot in lot_diff["new_lots"]:
        mo, lot_no, device, _ = lot["key"]
        values = lot["values"]
        new_ws.append([mo, lot_no, device] + [values.get(col, 0) for col in process_columns])
    _apply_readability_formatting(new_ws, number_format_columns=process_number_format_columns, number_format=value_number_format)

    removed_ws = wb.create_sheet("삭제랏")
    removed_ws.append(["MO", "랏번호", "디바이스"] + process_headers)
    for lot in lot_diff["removed_lots"]:
        mo, lot_no, device, _ = lot["key"]
        values = lot["values"]
        removed_ws.append([mo, lot_no, device] + [values.get(col, 0) for col in process_columns])
    _apply_readability_formatting(removed_ws, number_format_columns=process_number_format_columns, number_format=value_number_format)

    wb.save(output_path)
```

(이 파일의 `build_highlighted_today_file` 함수는 이번 태스크에서 건드리지 않는다 — Task 10에서 수정.)

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_report.py -v`
Expected: PASS (4 passed — 이 태스크는 `build_highlighted_today_file`을 건드리지 않으므로 그 테스트도 그대로 통과해야 함)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/report.py tests/test_report.py
git commit -m "feat: generalize build_variance_report to use dynamic process columns and configurable number format"
```

---

### Task 10: report.py - build_highlighted_today_file이 sheet_name을 명시적으로 받도록 변경

**Files:**
- Modify: `scm_wip_diff/report.py`
- Test: `tests/test_report.py`

**배경:** `build_highlighted_today_file`이 지금까지 `wb.sheetnames[0]`(첫 번째 시트)을 무조건 사용했다. ATX 파일은 대상 시트 이름이 날짜마다 바뀌므로(그리고 항상 첫 번째 시트라는 보장도 원칙적으로 없으므로), 어떤 시트에 하이라이트를 적용할지 명시적으로 인자를 받는다.

- [ ] **Step 1: 실패하는 테스트로 기존 테스트 갱신**

`tests/test_report.py`의 `test_build_highlighted_today_file_marks_cells_without_touching_original` 안의 호출부를 수정:
```python
    output_path = tmp_path / "260722 GTK WIP_변동표시.xlsx"
    build_highlighted_today_file(str(today_copy), lot_diff, str(output_path), "gtk3387")
```
(기존에는 4번째 인자 없이 3개 인자로 호출하던 것을 `"gtk3387"` — 실제 GTK 픽스처의 시트 이름 — 추가)

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_report.py::test_build_highlighted_today_file_marks_cells_without_touching_original -v`
Expected: FAIL with `TypeError: build_highlighted_today_file() takes 3 positional arguments but 4 were given`

- [ ] **Step 3: build_highlighted_today_file 수정**

`scm_wip_diff/report.py`의 `build_highlighted_today_file`을 다음으로 교체:
```python
def build_highlighted_today_file(today_path, lot_diff, output_path, sheet_name):
    shutil.copyfile(today_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    ws = wb[sheet_name]

    for lot in lot_diff["changed_lots"]:
        row = lot["row_in_today"]
        for change in lot["changes"]:
            ws.cell(row=row, column=change["col"]).fill = RED_FILL

    for lot in lot_diff["new_lots"]:
        row = lot["row_in_today"]
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = NEW_LOT_FILL

    wb.save(output_path)
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_report.py -v`
Expected: PASS (4 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/report.py tests/test_report.py
git commit -m "feat: require explicit sheet_name for highlighting instead of assuming the first sheet"
```

---

### Task 11: report.py - derive_output_paths의 회사명 하드코딩 제거

**Files:**
- Modify: `scm_wip_diff/report.py`
- Test: `tests/test_report.py`

**배경:** `derive_output_paths`가 리포트 파일명에 `"GTK"`를 하드코딩하고 있어서(`f"{date_prefix}_GTK_WIP_변동리포트.xlsx"`), ATX 파일에 적용하면 `260723_GTK_WIP_변동리포트.xlsx`처럼 잘못된 이름이 나온다. 파일명에서 날짜와 함께 회사 토큰도 추출하도록 일반화한다. 이 추출은 표시용 라벨일 뿐이며, GTK/ATX 판별 자체는 `format_detect.py`가 파일 내용으로 이미 수행한다.

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/test_report.py`에 추가:
```python
def test_derive_output_paths_works_for_atx_filename():
    today_path = os.path.join("C:", os.sep, "data", "260723 ATX WIP.xlsx")

    report_path, highlighted_path = derive_output_paths(today_path)

    assert report_path == os.path.join("C:", os.sep, "data", "260723_ATX_WIP_변동리포트.xlsx")
    assert highlighted_path == os.path.join("C:", os.sep, "data", "260723 ATX WIP_변동표시.xlsx")


def test_derive_output_paths_falls_back_when_pattern_does_not_match():
    today_path = os.path.join("C:", os.sep, "data", "random_file.xlsx")

    report_path, highlighted_path = derive_output_paths(today_path)

    assert report_path == os.path.join("C:", os.sep, "data", "random_file_WIP_변동리포트.xlsx")
    assert highlighted_path == os.path.join("C:", os.sep, "data", "random_file_변동표시.xlsx")
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_report.py -k derive_output_paths -v`
Expected: FAIL (`test_derive_output_paths_works_for_atx_filename`이 `260723_GTK_WIP_변동리포트.xlsx`를 반환해 assertion 실패)

- [ ] **Step 3: derive_output_paths 수정**

`scm_wip_diff/report.py`의 `derive_output_paths`를 다음으로 교체:
```python
def derive_output_paths(today_path):
    folder = os.path.dirname(today_path)
    basename = os.path.basename(today_path)
    name, ext = os.path.splitext(basename)
    match = re.match(r"^(\d{6})\s+(\S+)\s+WIP", basename)
    if match:
        date_prefix, company = match.group(1), match.group(2)
        report_name = f"{date_prefix}_{company}_WIP_변동리포트.xlsx"
    else:
        report_name = f"{name}_WIP_변동리포트.xlsx"
    report_path = os.path.join(folder, report_name)
    highlighted_path = os.path.join(folder, f"{name}_변동표시{ext}")
    return report_path, highlighted_path
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_report.py -v`
Expected: PASS (6 passed — 기존 `test_derive_output_paths_uses_date_prefix_and_same_folder`도 그대로 통과해야 함)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/report.py tests/test_report.py
git commit -m "feat: extract company token from filename instead of hardcoding GTK in output names"
```

---

### Task 12: report.py - ATX 실제 픽스처로 변동리포트/하이라이트 통합 테스트

**Files:**
- Modify: `tests/test_report.py`

**배경:** Task 7~11에서 각 함수를 개별적으로 일반화했으니, 실제 ATX 픽스처로 전체 파이프라인(파싱 → 비교 → 리포트 생성 → 하이라이트)이 맞물려 동작하는지 통합 테스트로 확인한다.

- [ ] **Step 1: 실패할 수 있는 통합 테스트 작성**

`tests/test_report.py`에 추가:
```python
from scm_wip_diff.atx_parser import parse_atx_wip_sheet
from scm_wip_diff.comparator import compare_lots, compare_stage_summary

FIXTURE_260723_ATX = "tests/fixtures/260723 ATX WIP.xlsx"
FIXTURE_260724_ATX = "tests/fixtures/260724 ATX WIP.xlsx"


def test_atx_end_to_end_report_and_highlight(tmp_path):
    yesterday = parse_atx_wip_sheet(FIXTURE_260723_ATX)
    today = parse_atx_wip_sheet(FIXTURE_260724_ATX)

    stage_summary = compare_stage_summary(yesterday, today)
    lot_diff = compare_lots(yesterday, today)

    assert set(stage_summary.keys()) == {"전공정", "후공정"}
    assert len(lot_diff["changed_lots"]) == 81
    assert len(lot_diff["new_lots"]) == 25
    assert len(lot_diff["removed_lots"]) == 9

    report_path = tmp_path / "report.xlsx"
    build_variance_report(
        stage_summary, lot_diff, str(report_path), today.column_labels, today.value_number_format
    )
    report_wb = openpyxl.load_workbook(str(report_path))
    assert report_wb["요약"]["A2"].value == "전공정"
    assert report_wb["변동랏"]["D2"].number_format == "#,##0.00"

    today_copy = tmp_path / "260724 ATX WIP.xlsx"
    shutil.copyfile(FIXTURE_260724_ATX, today_copy)
    highlighted_path = tmp_path / "260724 ATX WIP_변동표시.xlsx"
    build_highlighted_today_file(str(today_copy), lot_diff, str(highlighted_path), today.sheet_name)

    highlighted_wb = openpyxl.load_workbook(str(highlighted_path))
    ws = highlighted_wb[today.sheet_name]
    first_changed = next(
        lot for lot in lot_diff["changed_lots"]
        if lot["key"] == ("BQ8886001A", "GTMP122007", "TPSJ26N009", 0)
    )
    changed_col = first_changed["changes"][0]["col"]
    assert ws.cell(row=first_changed["row_in_today"], column=changed_col).fill.fgColor.rgb == "FFFF0000"
```

- [ ] **Step 2: 테스트 실행**

Run: `python -m pytest tests/test_report.py::test_atx_end_to_end_report_and_highlight -v`
Expected: PASS (Task 7~11이 모두 올바르게 구현되었다면 첫 실행에 통과해야 함. 실패하면 Task 7~11 중 어느 단계가 잘못됐는지 조사한다 — 숫자를 임의로 맞추지 말 것)

- [ ] **Step 3: 전체 테스트 스위트 실행**

Run: `python -m pytest tests/ -v`
Expected: 모두 PASS

- [ ] **Step 4: 커밋**

```bash
git add tests/test_report.py
git commit -m "test: add end-to-end ATX report and highlight integration test"
```

---

### Task 13: gui.py - format_detect 연동 및 형식 불일치 처리

**Files:**
- Modify: `scm_wip_diff/gui.py`

**배경:** GUI가 GTK 전용 `parse_wip_sheet`를 직접 호출하던 것을 `format_detect.parse_wip_file`을 통하도록 바꾼다. 어제/오늘 파일의 형식이 다르면 비교를 중단하고 안내한다. `build_variance_report`/`build_highlighted_today_file`의 새 시그니처(`value_number_format`, `sheet_name`)에 맞춰 호출부도 갱신한다. `_show_preview`는 `stage_summary`에 실제로 있는 단계만 순회하도록 고친다 (ATX는 "완료"가 없으므로).

이 태스크는 GUI라 자동 테스트 대신 수동/기능적 스모크 테스트로 검증한다 (기존 GTK 작업과 동일한 정책).

- [ ] **Step 1: gui.py 전체 교체**

`scm_wip_diff/gui.py`:
```python
"""tkinter GUI wiring format detection -> comparator -> report together."""
import tkinter as tk
from tkinter import filedialog, messagebox

from scm_wip_diff.comparator import check_lot_overlap, compare_lots, compare_stage_summary
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

        tk.Label(root, text="어제 파일:").grid(row=0, column=0, sticky="w")
        tk.Entry(root, textvariable=self.yesterday_path, width=60).grid(row=0, column=1)
        tk.Button(root, text="찾아보기", command=self.choose_yesterday).grid(row=0, column=2)

        tk.Label(root, text="오늘 파일:").grid(row=1, column=0, sticky="w")
        tk.Entry(root, textvariable=self.today_path, width=60).grid(row=1, column=1)
        tk.Button(root, text="찾아보기", command=self.choose_today).grid(row=1, column=2)

        tk.Button(root, text="비교 실행", command=self.run_compare).grid(row=2, column=1)

        self.result_text = tk.Text(root, width=100, height=30)
        self.result_text.grid(row=3, column=0, columnspan=3)

    def choose_yesterday(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.yesterday_path.set(path)

    def choose_today(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.today_path.set(path)

    def run_compare(self):
        y_path = self.yesterday_path.get()
        t_path = self.today_path.get()
        if not y_path or not t_path:
            messagebox.showerror("입력 필요", "어제 파일과 오늘 파일을 모두 선택하세요")
            return

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
            lot_diff = compare_lots(yesterday, today)

            report_path, highlighted_path = derive_output_paths(t_path)

            try:
                build_variance_report(
                    stage_summary, lot_diff, report_path, today.column_labels, today.value_number_format
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

            self._show_preview(stage_summary, lot_diff)
            messagebox.showinfo("완료", f"저장 완료:\n{report_path}\n{highlighted_path}")
        except Exception as e:
            messagebox.showerror("예상치 못한 오류", str(e))

    def _show_preview(self, stage_summary, lot_diff):
        self.result_text.delete("1.0", tk.END)
        lines = ["[단계별 합계]"]
        for stage in stage_summary:
            s = stage_summary[stage]
            lines.append(f"{stage}: {s['yesterday']:,} -> {s['today']:,} ({s['delta']:+,})")
        lines.append("")
        lines.append(f"변경된 랏 수: {len(lot_diff['changed_lots'])}")
        lines.append(f"신규 랏 수: {len(lot_diff['new_lots'])}")
        lines.append(f"삭제된 랏 수: {len(lot_diff['removed_lots'])}")
        self.result_text.insert("1.0", "\n".join(lines))
```

- [ ] **Step 2: 기능적 스모크 테스트 (GTK 회귀 확인)**

`run_compare`와 동일한 호출 순서를 real GTK 픽스처로 재현해 회귀가 없는지 확인한다 (디스플레이 없는 환경이면 tkinter 창 대신 아래처럼 로직만 직접 실행):
```bash
python -c "
from scm_wip_diff.format_detect import parse_wip_file
from scm_wip_diff.comparator import compare_stage_summary, compare_lots
from scm_wip_diff.report import build_variance_report, build_highlighted_today_file, derive_output_paths
import tempfile, os, shutil

y_fmt, yesterday = parse_wip_file('tests/fixtures/260721 GTK WIP.xlsx')
t_fmt, today = parse_wip_file('tests/fixtures/260722 GTK WIP.xlsx')
assert y_fmt == t_fmt == 'GTK'

stage_summary = compare_stage_summary(yesterday, today)
lot_diff = compare_lots(yesterday, today)
assert len(lot_diff['changed_lots']) == 19

with tempfile.TemporaryDirectory() as d:
    report_path = os.path.join(d, 'report.xlsx')
    build_variance_report(stage_summary, lot_diff, report_path, today.column_labels, today.value_number_format)
    today_copy = os.path.join(d, 'today.xlsx')
    shutil.copyfile('tests/fixtures/260722 GTK WIP.xlsx', today_copy)
    highlighted_path = os.path.join(d, 'highlighted.xlsx')
    build_highlighted_today_file(today_copy, lot_diff, highlighted_path, today.sheet_name)
    print('GTK end-to-end OK')
"
```
Expected: `GTK end-to-end OK` 출력, 예외 없음

- [ ] **Step 3: 기능적 스모크 테스트 (ATX 신규 지원 확인)**

```bash
python -c "
from scm_wip_diff.format_detect import parse_wip_file
from scm_wip_diff.comparator import compare_stage_summary, compare_lots
from scm_wip_diff.report import build_variance_report, build_highlighted_today_file, derive_output_paths
import tempfile, os, shutil

y_fmt, yesterday = parse_wip_file('tests/fixtures/260723 ATX WIP.xlsx')
t_fmt, today = parse_wip_file('tests/fixtures/260724 ATX WIP.xlsx')
assert y_fmt == t_fmt == 'ATX'

stage_summary = compare_stage_summary(yesterday, today)
lot_diff = compare_lots(yesterday, today)
assert len(lot_diff['changed_lots']) == 81
assert '완료' not in stage_summary

with tempfile.TemporaryDirectory() as d:
    report_path = os.path.join(d, 'report.xlsx')
    build_variance_report(stage_summary, lot_diff, report_path, today.column_labels, today.value_number_format)
    today_copy = os.path.join(d, 'today.xlsx')
    shutil.copyfile('tests/fixtures/260724 ATX WIP.xlsx', today_copy)
    highlighted_path = os.path.join(d, 'highlighted.xlsx')
    build_highlighted_today_file(today_copy, lot_diff, highlighted_path, today.sheet_name)
    print('ATX end-to-end OK')
"
```
Expected: `ATX end-to-end OK` 출력, 예외 없음

- [ ] **Step 4: 형식 불일치 스모크 테스트**

```bash
python -c "
from scm_wip_diff.format_detect import parse_wip_file

_, yesterday_fmt_pair = 'GTK', parse_wip_file('tests/fixtures/260721 GTK WIP.xlsx')
_, today_fmt_pair = 'ATX', parse_wip_file('tests/fixtures/260723 ATX WIP.xlsx')
y_fmt = yesterday_fmt_pair[0]
t_fmt = today_fmt_pair[0]
print('yesterday format:', y_fmt, '/ today format:', t_fmt)
assert y_fmt != t_fmt
print('mismatch correctly detected (gui.py would show an error dialog here)')
"
```
Expected: `mismatch correctly detected...` 출력

- [ ] **Step 5: 전체 테스트 스위트 재확인**

Run: `python -m pytest tests/ -v`
Expected: 모두 PASS (pytest 대상은 변경 없음 — gui.py는 자동 테스트 제외 정책 유지)

- [ ] **Step 6: 커밋**

```bash
git add scm_wip_diff/gui.py
git commit -m "feat: wire GUI to auto-detected GTK/ATX parsing with format-mismatch guard"
```

---

## 실행 시 참고사항

- 매일 사용 시 흐름은 변경되지 않는다: `python -m scm_wip_diff.main` 실행 → 어제/오늘 파일 선택 → 비교 실행. 이제 GTK/ATX 파일을 구분해서 고를 필요 없이 같은 방식으로 사용하면 된다.
- ATX의 `KSWIPFT (Test)`, `KSFG (완료)`, `CUST TKW` 시트는 이번 범위 밖이다 (설계서 "범위 밖" 참조). 향후 필요하면 별도 브레인스토밍/설계 사이클을 거친다.
- exe 재빌드가 필요하면 기존과 동일하게 `python -m PyInstaller --onefile --windowed --name GTK_WIP_Diff run.py`를 실행한다 (엔트리 스크립트 `run.py`는 이번 작업으로 변경되지 않음).
