# GTK WIP 일일 비교 리포트 도구 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 매일 전달되는 `<YYMMDD> GTK WIP.xlsx` 두 파일(어제/오늘)을 GUI에서 선택하면, 반도체 공정 진행 컬럼(I~AD)의 변동 내역을 엑셀 리포트로 생성하고 오늘 파일에 변경된 셀을 빨간색으로 하이라이트해주는 Windows 데스크톱 도구를 만든다.

**Architecture:** `parser.py`(다중 리포트가 이어붙은 시트에서 대상 표만 추출) → `comparator.py`(MO+랏번호+디바이스+등장순번 키로 두 스냅샷 비교) → `report.py`(변동리포트 엑셀 생성 + 오늘 파일 하이라이트 복사본 생성) → `gui.py`/`main.py`(tkinter로 위 파이프라인을 연결). 각 모듈은 독립적으로 단위 테스트 가능하도록 순수 함수/데이터클래스로 설계한다.

**Tech Stack:** Python 3.14, openpyxl(엑셀 읽기/쓰기), tkinter(GUI, 표준 라이브러리), pytest(테스트)

**설계서:** `docs/superpowers/specs/2026-07-28-gtk-wip-diff-design.md`

---

## 파일 구조

```
SCM_process/
├── scm_wip_diff/
│   ├── __init__.py
│   ├── parser.py        # 엑셀에서 WIP 표 추출
│   ├── comparator.py     # 두 스냅샷 비교 → 변동 목록
│   ├── report.py         # 변동리포트 엑셀 + 하이라이트 파일 생성
│   ├── gui.py             # tkinter 화면
│   └── main.py            # 진입점
├── tests/
│   ├── conftest.py
│   ├── fixtures/
│   │   ├── 260721 GTK WIP.xlsx
│   │   └── 260722 GTK WIP.xlsx
│   ├── test_parser.py
│   ├── test_comparator.py
│   └── test_report.py
├── requirements.txt
└── .gitignore
```

## 핵심 데이터 구조 (모든 태스크에서 공통으로 사용)

```python
# parser.ParsedWip
@dataclass
class ParsedWip:
    column_labels: dict   # {col_index(int): "Saw" 같은 라벨 str}
    stage_groups: dict    # {"전공정": [9,10,11], "후공정": [12..28], "완료": [29,30]}
    lots: dict            # {(mo,lot,device,occurrence_idx): {col_index: int}}
    rows: dict            # {(mo,lot,device,occurrence_idx): 엑셀 행번호}

# comparator.compare_stage_summary(yesterday, today) 반환값
{
  "전공정": {"yesterday": int, "today": int, "delta": int},
  "후공정": {...},
  "완료": {...},
}

# comparator.compare_lots(yesterday, today) 반환값
{
  "changed_lots": [
    {
      "key": (mo, lot, device, idx),
      "row_in_today": int,
      "changes": [{"col": int, "label": "Saw(I)", "before": int, "after": int}, ...],
    },
    ...
  ],
  "new_lots": [{"key": (mo,lot,device,idx), "row_in_today": int}, ...],
  "removed_lots": [{"key": (mo,lot,device,idx), "row_in_yesterday": int}, ...],
}
```

키의 4번째 요소(occurrence_idx)는 같은 파일 안에서 MO+랏번호+디바이스가 완전히 동일한 행이 여러 개 나올 때(실측 데이터에서 확인됨) 등장 순서로 구분하기 위함이다.

---

### Task 1: 프로젝트 스캐폴딩

**Files:**
- Create: `scm_wip_diff/__init__.py`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`
- Create: `tests/fixtures/260721 GTK WIP.xlsx` (기존 루트 파일 복사)
- Create: `tests/fixtures/260722 GTK WIP.xlsx` (기존 루트 파일 복사)
- Create: `requirements.txt`
- Create: `.gitignore`

- [ ] **Step 1: git 저장소 초기화**

Run:
```bash
git init
```
Expected: `Initialized empty Git repository in .../SCM_process/.git/`

- [ ] **Step 2: 디렉터리 및 빈 패키지 파일 생성**

```bash
mkdir -p scm_wip_diff tests/fixtures docs/superpowers/plans docs/superpowers/specs
touch scm_wip_diff/__init__.py tests/__init__.py
```

- [ ] **Step 3: 테스트에서 프로젝트 루트를 import 경로에 추가하는 conftest.py 작성**

`tests/conftest.py`:
```python
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
```

- [ ] **Step 4: 샘플 엑셀 파일을 테스트 픽스처로 복사**

```bash
cp "260721 GTK WIP.xlsx" "tests/fixtures/260721 GTK WIP.xlsx"
cp "260722 GTK WIP.xlsx" "tests/fixtures/260722 GTK WIP.xlsx"
```

- [ ] **Step 5: requirements.txt 작성**

`requirements.txt`:
```
openpyxl>=3.1
pytest>=8.0
```

- [ ] **Step 6: .gitignore 작성**

`.gitignore`:
```
/*.xlsx
__pycache__/
*.pyc
.pytest_cache/
```

이 패턴은 루트에 매일 새로 놓이는 `<YYMMDD> GTK/ATX WIP.xlsx` 원본 파일은 커밋 대상에서 제외하고, `tests/fixtures/` 아래 고정된 테스트 픽스처는 정상적으로 추적한다.

- [ ] **Step 7: 의존성 설치**

```bash
pip install -r requirements.txt
```
Expected: openpyxl, pytest 설치 완료 메시지

- [ ] **Step 8: 커밋**

```bash
git add scm_wip_diff tests requirements.txt .gitignore docs
git commit -m "chore: scaffold project structure and test fixtures"
```

---

### Task 2: parser - 표 시작 위치(anchor) 찾기

**Files:**
- Create: `scm_wip_diff/parser.py`
- Test: `tests/test_parser.py`

**배경:** 한 시트 안에 서로 다른 5개의 리포트("Report for Assy or Turn-key WIP", Pin count 요약, Hold Wafer Status, PFT WIP, TAP WIP)가 이어붙어 있다. 대상 표는 첫 번째뿐이므로, 고정 행번호 대신 "Report for Assy or Turn-key WIP" 텍스트가 있는 행을 찾아 기준점으로 삼는다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_parser.py`:
```python
import openpyxl
import pytest

from scm_wip_diff.parser import find_report_anchor, ReportFormatError

FIXTURE_260721 = "tests/fixtures/260721 GTK WIP.xlsx"
FIXTURE_260722 = "tests/fixtures/260722 GTK WIP.xlsx"


def test_find_report_anchor_locates_title_row():
    wb = openpyxl.load_workbook(FIXTURE_260721, data_only=True)
    ws = wb[wb.sheetnames[0]]

    anchor = find_report_anchor(ws)

    assert anchor == 2


def test_find_report_anchor_raises_when_title_missing():
    wb = openpyxl.Workbook()
    ws = wb.active

    with pytest.raises(ReportFormatError):
        find_report_anchor(ws)
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_parser.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'scm_wip_diff.parser'` (parser.py가 아직 없음)

- [ ] **Step 3: parser.py에 find_report_anchor 구현**

`scm_wip_diff/parser.py`:
```python
"""Parse GTK WIP Excel reports into structured lot-level data."""

REPORT_TITLE = "Report for Assy or Turn-key WIP"
PROCESS_COL_START = 9   # I
PROCESS_COL_END = 30    # AD


class ReportFormatError(Exception):
    """Raised when the expected WIP report table cannot be located."""


def find_report_anchor(ws, title=REPORT_TITLE):
    for row in ws.iter_rows(min_row=1, max_row=10, max_col=4):
        for cell in row:
            if cell.value and title in str(cell.value):
                return cell.row
    raise ReportFormatError(f"'{title}' 표를 찾을 수 없습니다")
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_parser.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/parser.py tests/test_parser.py
git commit -m "feat: locate WIP report table anchor row"
```

---

### Task 3: parser - 전공정/후공정/완료 컬럼 그룹 추출

**Files:**
- Modify: `scm_wip_diff/parser.py`
- Test: `tests/test_parser.py`

**배경:** 3행(anchor+1행)에 `I3:K3`(전공정), `L3:AB3`(후공정), `AC3:AD3`(완료) 병합 셀이 있다. 컬럼 letter를 하드코딩하지 않고 이 병합 범위를 읽어 그룹을 동적으로 결정한다.

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/test_parser.py`에 추가:
```python
from scm_wip_diff.parser import get_stage_groups


def test_get_stage_groups_matches_merged_header_cells():
    wb = openpyxl.load_workbook(FIXTURE_260721, data_only=True)
    ws = wb[wb.sheetnames[0]]
    anchor = find_report_anchor(ws)

    groups = get_stage_groups(ws, anchor)

    assert groups["전공정"] == [9, 10, 11]
    assert groups["후공정"] == list(range(12, 29))
    assert groups["완료"] == [29, 30]
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_parser.py::test_get_stage_groups_matches_merged_header_cells -v`
Expected: FAIL with `ImportError: cannot import name 'get_stage_groups'`

- [ ] **Step 3: get_stage_groups 구현**

`scm_wip_diff/parser.py`에 추가:
```python
def get_stage_groups(ws, anchor_row):
    group_row = anchor_row + 1
    groups = {}
    for merged in ws.merged_cells.ranges:
        if merged.min_row == group_row:
            label = str(ws.cell(row=group_row, column=merged.min_col).value).strip()
            groups[label] = list(range(merged.min_col, merged.max_col + 1))
    return groups
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_parser.py -v`
Expected: PASS (3 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/parser.py tests/test_parser.py
git commit -m "feat: derive process stage column groups from merged header cells"
```

---

### Task 4: parser - 컬럼 라벨 생성 (2줄 헤더 결합)

**Files:**
- Modify: `scm_wip_diff/parser.py`
- Test: `tests/test_parser.py`

**배경:** 컬럼명이 4행/5행 두 줄에 걸쳐 있다. 예: J열은 4행 "Die" + 5행 "Mount" = "Die Mount"(다이 마운트 공정), K열은 4행 "Wire" + 5행 "Bond" = "Wire Bond"(와이어 본딩). 4행이 비어있으면 5행 텍스트만 사용한다(예: I열은 "Saw").

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/test_parser.py`에 추가:
```python
from scm_wip_diff.parser import build_column_labels


def test_build_column_labels_combines_two_header_rows():
    wb = openpyxl.load_workbook(FIXTURE_260721, data_only=True)
    ws = wb[wb.sheetnames[0]]
    anchor = find_report_anchor(ws)

    labels = build_column_labels(ws, anchor)

    assert labels[9] == "Saw"
    assert labels[10] == "Die Mount"
    assert labels[11] == "Wire Bond"
    assert labels[21] == "Ball Mount"
    assert labels[29] == "TR Stock"
    assert labels[30] == "Non TR Stock"
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_parser.py::test_build_column_labels_combines_two_header_rows -v`
Expected: FAIL with `ImportError: cannot import name 'build_column_labels'`

- [ ] **Step 3: build_column_labels 구현**

`scm_wip_diff/parser.py`에 추가:
```python
def build_column_labels(ws, anchor_row):
    row1 = anchor_row + 2
    row2 = anchor_row + 3
    labels = {}
    for col in range(1, ws.max_column + 1):
        v1 = ws.cell(row=row1, column=col).value
        v2 = ws.cell(row=row2, column=col).value
        v1 = str(v1).strip() if v1 else ""
        v2 = str(v2).strip() if v2 else ""
        labels[col] = f"{v1} {v2}".strip() if v1 else v2
    return labels
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_parser.py -v`
Expected: PASS (4 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/parser.py tests/test_parser.py
git commit -m "feat: build combined column labels from two header rows"
```

---

### Task 5: parser - 전체 시트 파싱 (parse_wip_sheet)

**Files:**
- Modify: `scm_wip_diff/parser.py`
- Test: `tests/test_parser.py`

**배경:** anchor, stage_groups, column_labels를 조합해 데이터 행(7행부터 B열이 빌 때까지)을 읽어 `ParsedWip`을 만든다. MO+랏번호+디바이스가 완전히 동일한 행이 있으면(실측 데이터에서 `('', '1B468601', 'TMP1201D')`가 18,19행에 중복 존재) 등장 순서를 키에 추가해 구분한다.

- [ ] **Step 1: 실패하는 테스트 추가**

`tests/test_parser.py`에 추가:
```python
from scm_wip_diff.parser import parse_wip_sheet


def test_parse_wip_sheet_reads_all_lot_rows_and_stops_before_next_report():
    parsed = parse_wip_sheet(FIXTURE_260721)

    assert len(parsed.lots) == 174


def test_parse_wip_sheet_extracts_process_values_for_first_data_row():
    parsed = parsed = parse_wip_sheet(FIXTURE_260721)
    key = ("SNF08N5623340E", "1B4686", "TMP1200D", 0)

    assert parsed.rows[key] == 7
    assert parsed.lots[key][9] == 0     # Saw
    assert parsed.lots[key][29] == 0    # TR Stock
    assert parsed.lots[key][30] == 24627  # Non TR Stock


def test_parse_wip_sheet_disambiguates_duplicate_keys_by_occurrence():
    parsed = parse_wip_sheet(FIXTURE_260721)

    first_key = ("", "1B468601", "TMP1201D", 0)
    second_key = ("", "1B468601", "TMP1201D", 1)

    assert parsed.rows[first_key] == 18
    assert parsed.rows[second_key] == 19
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_parser.py -k parse_wip_sheet -v`
Expected: FAIL with `ImportError: cannot import name 'parse_wip_sheet'`

- [ ] **Step 3: ParsedWip 데이터클래스와 parse_wip_sheet 구현**

`scm_wip_diff/parser.py` 상단에 import 추가:
```python
from dataclasses import dataclass

import openpyxl
```

파일 끝에 추가:
```python
@dataclass
class ParsedWip:
    column_labels: dict
    stage_groups: dict
    lots: dict
    rows: dict


def parse_wip_sheet(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    anchor = find_report_anchor(ws)
    groups = get_stage_groups(ws, anchor)
    labels = build_column_labels(ws, anchor)
    data_start = anchor + 5

    lots = {}
    rows = {}
    occurrence = {}
    r = data_start
    while ws.cell(row=r, column=2).value is not None:
        a = str(ws.cell(row=r, column=1).value or "").strip()
        b = str(ws.cell(row=r, column=2).value or "").strip()
        c = str(ws.cell(row=r, column=3).value or "").strip()
        base_key = (a, b, c)
        idx = occurrence.get(base_key, 0)
        occurrence[base_key] = idx + 1
        key = (a, b, c, idx)

        values = {}
        for col in range(PROCESS_COL_START, PROCESS_COL_END + 1):
            v = ws.cell(row=r, column=col).value
            values[col] = int(v) if isinstance(v, (int, float)) else 0

        lots[key] = values
        rows[key] = r
        r += 1

    return ParsedWip(column_labels=labels, stage_groups=groups, lots=lots, rows=rows)
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_parser.py -v`
Expected: PASS (7 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/parser.py tests/test_parser.py
git commit -m "feat: parse WIP sheet into ParsedWip with occurrence-disambiguated keys"
```

---

### Task 6: comparator - 단계별 합계 비교

**Files:**
- Create: `scm_wip_diff/comparator.py`
- Test: `tests/test_comparator.py`

- [ ] **Step 1: 합성 데이터로 단위 테스트 작성**

`tests/test_comparator.py`:
```python
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
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_comparator.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'scm_wip_diff.comparator'`

- [ ] **Step 3: comparator.py에 compare_stage_summary 구현**

`scm_wip_diff/comparator.py`:
```python
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
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_comparator.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/comparator.py tests/test_comparator.py
git commit -m "feat: compare stage-level quantity totals between two snapshots"
```

---

### Task 7: comparator - 랏 단위 변경/신규/삭제 감지

**Files:**
- Modify: `scm_wip_diff/comparator.py`
- Test: `tests/test_comparator.py`

**배경:** I~AD 컬럼을 하나씩 비교해 값이 달라진 컬럼만 컬럼 순서(I→AD) 그대로 나열한다. 오늘에만 있는 키는 신규, 어제에만 있는 키는 삭제로 분류한다.

- [ ] **Step 1: 합성 데이터 + 실제 픽스처 테스트 작성**

`tests/test_comparator.py`에 추가:
```python
from scm_wip_diff.comparator import compare_lots


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
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_comparator.py -k compare_lots -v`
Expected: FAIL with `ImportError: cannot import name 'compare_lots'`

- [ ] **Step 3: compare_lots 구현**

`scm_wip_diff/comparator.py` 상단에 import 추가:
```python
from openpyxl.utils import get_column_letter
```

파일 끝에 추가:
```python
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
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_comparator.py -v`
Expected: PASS (5 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/comparator.py tests/test_comparator.py
git commit -m "feat: detect per-lot changed columns and new/removed lots"
```

---

### Task 7.5: comparator - 파일 불일치(잘못된 파일 선택) 경고 판단

**Files:**
- Modify: `scm_wip_diff/comparator.py`
- Test: `tests/test_comparator.py`

**배경:** 설계서 "오류 처리" 항목: 어제/오늘 파일의 랏 구성이 크게 다르면(예: 잘못된 파일을 골랐을 때) 경고만 하고 진행은 허용해야 한다. 겹치는 키의 비율이 낮으면 경고 메시지를 반환하는 순수 함수로 구현한다 (실제 GUI 표시는 Task 11에서 연결).

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_comparator.py`에 추가:
```python
from scm_wip_diff.comparator import check_lot_overlap


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
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_comparator.py -k check_lot_overlap -v`
Expected: FAIL with `ImportError: cannot import name 'check_lot_overlap'`

- [ ] **Step 3: check_lot_overlap 구현**

`scm_wip_diff/comparator.py` 파일 끝에 추가:
```python
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
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_comparator.py -v`
Expected: PASS (8 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/comparator.py tests/test_comparator.py
git commit -m "feat: warn when yesterday/today lot sets barely overlap"
```

---

### Task 8: report - 출력 파일 경로 규칙 (derive_output_paths)

**Files:**
- Create: `scm_wip_diff/report.py`
- Test: `tests/test_report.py`

**배경:** 오늘 파일명(`260722 GTK WIP.xlsx`)에서 날짜(`260722`)를 추출해 리포트 파일명을 만들고, 하이라이트 파일은 원본 파일명 뒤에 `_변동표시`를 붙인다. 둘 다 오늘 파일과 같은 폴더에 저장한다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_report.py`:
```python
import os

from scm_wip_diff.report import derive_output_paths


def test_derive_output_paths_uses_date_prefix_and_same_folder():
    today_path = os.path.join("C:", os.sep, "data", "260722 GTK WIP.xlsx")

    report_path, highlighted_path = derive_output_paths(today_path)

    assert report_path == os.path.join("C:", os.sep, "data", "260722_GTK_WIP_변동리포트.xlsx")
    assert highlighted_path == os.path.join("C:", os.sep, "data", "260722 GTK WIP_변동표시.xlsx")
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_report.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'scm_wip_diff.report'`

- [ ] **Step 3: derive_output_paths 구현**

`scm_wip_diff/report.py`:
```python
"""Generate the variance report workbook and the highlighted today-file copy."""
import os
import re
import shutil

import openpyxl
from openpyxl.styles import PatternFill

RED_FILL = PatternFill(fill_type="solid", fgColor="FFFF0000")
NEW_LOT_FILL = PatternFill(fill_type="solid", fgColor="FFADD8E6")

STAGE_ORDER = ["전공정", "후공정", "완료"]


def derive_output_paths(today_path):
    folder = os.path.dirname(today_path)
    basename = os.path.basename(today_path)
    name, ext = os.path.splitext(basename)
    match = re.match(r"^(\d{6})", basename)
    date_prefix = match.group(1) if match else name
    report_path = os.path.join(folder, f"{date_prefix}_GTK_WIP_변동리포트.xlsx")
    highlighted_path = os.path.join(folder, f"{name}_변동표시{ext}")
    return report_path, highlighted_path
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_report.py -v`
Expected: PASS (1 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/report.py tests/test_report.py
git commit -m "feat: derive output file paths from today's filename"
```

---

### Task 9: report - 변동리포트 엑셀 생성 (build_variance_report)

**Files:**
- Modify: `scm_wip_diff/report.py`
- Test: `tests/test_report.py`

**배경:** 요약/변동랏/신규랏/삭제랏 4개 시트를 가진 새 엑셀 파일을 만든다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_report.py`에 추가:
```python
import openpyxl

from scm_wip_diff.report import build_variance_report

STAGE_SUMMARY = {
    "전공정": {"yesterday": 100, "today": 50, "delta": -50},
    "후공정": {"yesterday": 10, "today": 110, "delta": 100},
    "완료": {"yesterday": 0, "today": 0, "delta": 0},
}

LOT_DIFF = {
    "changed_lots": [
        {
            "key": ("m1", "l1", "d1", 0),
            "row_in_today": 7,
            "changes": [{"col": 9, "label": "Saw(I)", "before": 347638, "after": 0}],
        },
    ],
    "new_lots": [{"key": ("m2", "l2", "d2", 0), "row_in_today": 8}],
    "removed_lots": [{"key": ("m3", "l3", "d3", 0), "row_in_yesterday": 9}],
}


def test_build_variance_report_writes_expected_sheets(tmp_path):
    output_path = tmp_path / "report.xlsx"

    build_variance_report(STAGE_SUMMARY, LOT_DIFF, str(output_path))

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

    new_ws = wb["신규랏"]
    assert new_ws["B2"].value == "l2"

    removed_ws = wb["삭제랏"]
    assert removed_ws["B2"].value == "l3"
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_report.py::test_build_variance_report_writes_expected_sheets -v`
Expected: FAIL with `ImportError: cannot import name 'build_variance_report'`

- [ ] **Step 3: build_variance_report 구현**

`scm_wip_diff/report.py` 파일 끝에 추가:
```python
def build_variance_report(stage_summary, lot_diff, output_path):
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

    changed_ws = wb.create_sheet("변동랏")
    changed_ws.append(["MO", "랏번호", "디바이스", "변경컬럼", "어제값", "오늘값"])
    for lot in lot_diff["changed_lots"]:
        mo, lot_no, device, _ = lot["key"]
        for change in lot["changes"]:
            changed_ws.append([mo, lot_no, device, change["label"], change["before"], change["after"]])

    new_ws = wb.create_sheet("신규랏")
    new_ws.append(["MO", "랏번호", "디바이스"])
    for lot in lot_diff["new_lots"]:
        mo, lot_no, device, _ = lot["key"]
        new_ws.append([mo, lot_no, device])

    removed_ws = wb.create_sheet("삭제랏")
    removed_ws.append(["MO", "랏번호", "디바이스"])
    for lot in lot_diff["removed_lots"]:
        mo, lot_no, device, _ = lot["key"]
        removed_ws.append([mo, lot_no, device])

    wb.save(output_path)
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_report.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/report.py tests/test_report.py
git commit -m "feat: build variance report workbook with summary/changed/new/removed sheets"
```

---

### Task 10: report - 오늘 파일 하이라이트 복사본 생성 (build_highlighted_today_file)

**Files:**
- Modify: `scm_wip_diff/report.py`
- Test: `tests/test_report.py`

**배경:** 원본 오늘 파일을 건드리지 않고 새 파일로 복사한 뒤, 값이 바뀐 셀은 빨간색, 신규 랏 행 전체는 파란색으로 칠한다.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_report.py`에 추가:
```python
import hashlib
import shutil

from scm_wip_diff.report import build_highlighted_today_file

FIXTURE_260722_PATH = "tests/fixtures/260722 GTK WIP.xlsx"


def _file_hash(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def test_build_highlighted_today_file_marks_cells_without_touching_original(tmp_path):
    today_copy = tmp_path / "260722 GTK WIP.xlsx"
    shutil.copyfile(FIXTURE_260722_PATH, today_copy)
    original_hash = _file_hash(today_copy)

    lot_diff = {
        "changed_lots": [
            {
                "key": ("TNT0896622324", "1B3C75", "TMP1202D", 0),
                "row_in_today": 21,
                "changes": [
                    {"col": 9, "label": "Saw(I)", "before": 347638, "after": 0},
                    {"col": 10, "label": "Die Mount(J)", "before": 316800, "after": 664438},
                ],
            },
        ],
        "new_lots": [{"key": ("MNS08M6622326", "1B3C76", "TMP1230", 0), "row_in_today": 22}],
        "removed_lots": [],
    }

    output_path = tmp_path / "260722 GTK WIP_변동표시.xlsx"
    build_highlighted_today_file(str(today_copy), lot_diff, str(output_path))

    assert _file_hash(today_copy) == original_hash

    wb = openpyxl.load_workbook(str(output_path))
    ws = wb[wb.sheetnames[0]]
    assert ws.cell(row=21, column=9).fill.fgColor.rgb == "FFFF0000"
    assert ws.cell(row=21, column=10).fill.fgColor.rgb == "FFFF0000"
    assert ws.cell(row=21, column=1).fill.fgColor.rgb != "FFFF0000"
    assert ws.cell(row=22, column=1).fill.fgColor.rgb == "FFADD8E6"
    assert ws.cell(row=22, column=38).fill.fgColor.rgb == "FFADD8E6"
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_report.py::test_build_highlighted_today_file_marks_cells_without_touching_original -v`
Expected: FAIL with `ImportError: cannot import name 'build_highlighted_today_file'`

- [ ] **Step 3: build_highlighted_today_file 구현**

`scm_wip_diff/report.py` 파일 끝에 추가:
```python
def build_highlighted_today_file(today_path, lot_diff, output_path):
    shutil.copyfile(today_path, output_path)
    wb = openpyxl.load_workbook(output_path)
    ws = wb[wb.sheetnames[0]]

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
Expected: PASS (3 passed)

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/report.py tests/test_report.py
git commit -m "feat: highlight changed cells and new-lot rows in a copy of today's file"
```

---

### Task 11: GUI 연결 (gui.py, main.py)

**Files:**
- Create: `scm_wip_diff/gui.py`
- Create: `scm_wip_diff/main.py`

**배경:** 파일 선택 → 비교 실행 → 미리보기 표시 → 리포트/하이라이트 파일 저장까지 연결하는 얇은 레이어. tkinter는 표준 라이브러리라 별도 설치가 필요 없다. 이 태스크는 GUI라 자동 테스트 대신 수동 스모크 테스트로 검증한다(설계서의 테스트 방침 참조).

- [ ] **Step 1: gui.py 작성**

`scm_wip_diff/gui.py`:
```python
"""tkinter GUI wiring parser -> comparator -> report together."""
import tkinter as tk
from tkinter import filedialog, messagebox

from scm_wip_diff.comparator import check_lot_overlap, compare_lots, compare_stage_summary
from scm_wip_diff.parser import ReportFormatError, parse_wip_sheet
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
            yesterday = parse_wip_sheet(y_path)
            today = parse_wip_sheet(t_path)
        except ReportFormatError as e:
            messagebox.showerror("파일 형식 오류", str(e))
            return

        overlap_warning = check_lot_overlap(yesterday, today)
        if overlap_warning:
            messagebox.showwarning("확인 필요", overlap_warning)

        stage_summary = compare_stage_summary(yesterday, today)
        lot_diff = compare_lots(yesterday, today)

        report_path, highlighted_path = derive_output_paths(t_path)
        try:
            build_variance_report(stage_summary, lot_diff, report_path)
            build_highlighted_today_file(t_path, lot_diff, highlighted_path)
        except PermissionError:
            messagebox.showerror("저장 실패", "파일이 열려있어 저장할 수 없습니다")
            return

        self._show_preview(stage_summary, lot_diff)
        messagebox.showinfo("완료", f"저장 완료:\n{report_path}\n{highlighted_path}")

    def _show_preview(self, stage_summary, lot_diff):
        self.result_text.delete("1.0", tk.END)
        lines = ["[단계별 합계]"]
        for stage in ("전공정", "후공정", "완료"):
            s = stage_summary[stage]
            lines.append(f"{stage}: {s['yesterday']:,} -> {s['today']:,} ({s['delta']:+,})")
        lines.append("")
        lines.append(
            f"변경된 랏: {len(lot_diff['changed_lots'])}건 / "
            f"신규: {len(lot_diff['new_lots'])}건 / "
            f"삭제: {len(lot_diff['removed_lots'])}건"
        )
        lines.append("")
        lines.append("[변경 상세]")
        for lot in lot_diff["changed_lots"]:
            _, lot_no, device, _ = lot["key"]
            changes_str = ", ".join(
                f"{c['label']} {c['before']:,}->{c['after']:,}" for c in lot["changes"]
            )
            lines.append(f"{lot_no}/{device}: {changes_str}")
        self.result_text.insert("1.0", "\n".join(lines))
```

- [ ] **Step 2: main.py 작성**

`scm_wip_diff/main.py`:
```python
import tkinter as tk

from scm_wip_diff.gui import App


def main():
    root = tk.Tk()
    root.title("GTK WIP 일일 비교")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: 수동 스모크 테스트**

Run:
```bash
python -m scm_wip_diff.main
```
확인 사항:
1. 창이 뜨는지
2. "찾아보기"로 `tests/fixtures/260721 GTK WIP.xlsx`(어제), `tests/fixtures/260722 GTK WIP.xlsx`(오늘) 선택 가능한지
3. "비교 실행" 클릭 시 에러 없이 미리보기 텍스트가 채워지는지 (전공정/후공정/완료 합계, 변경 19건/신규 1건/삭제 2건이 보여야 함)
4. `tests/fixtures/` 폴더에 `260722_GTK_WIP_변동리포트.xlsx`와 `260722 GTK WIP_변동표시.xlsx`가 생성됐는지
5. 생성된 `260722 GTK WIP_변동표시.xlsx`를 엑셀로 열어 21행 I,J열이 빨간색인지, 22행이 파란색인지 육안 확인

- [ ] **Step 4: 스모크 테스트로 생성된 파일 정리 (테스트 픽스처 폴더를 깨끗하게 유지)**

```bash
rm "tests/fixtures/260722_GTK_WIP_변동리포트.xlsx" "tests/fixtures/260722 GTK WIP_변동표시.xlsx"
```

- [ ] **Step 5: 커밋**

```bash
git add scm_wip_diff/gui.py scm_wip_diff/main.py
git commit -m "feat: wire tkinter GUI to parser/comparator/report pipeline"
```

---

### Task 12: 전체 테스트 스위트 최종 확인

**Files:** 없음 (검증 전용)

- [ ] **Step 1: 전체 테스트 스위트 실행**

Run: `python -m pytest tests/ -v`
Expected: 모든 테스트 PASS (parser 7개 + comparator 8개 + report 3개 = 18개)

- [ ] **Step 2: tests/fixtures/ 폴더에 부산물이 남지 않았는지 확인**

Run: `ls tests/fixtures/`
Expected: `260721 GTK WIP.xlsx`, `260722 GTK WIP.xlsx` 두 개만 존재

- [ ] **Step 3: 최종 상태 확인 및 커밋 (필요 시)**

```bash
git status
```
Expected: `nothing to commit, working tree clean` (Task 11에서 이미 커밋되었으므로 추가 커밋 불필요할 가능성 높음)

---

## 실행 시 참고사항

- 매일 사용 시: `python -m scm_wip_diff.main` 실행 → 어제/오늘 파일 선택 → 비교 실행 → 오늘 파일과 같은 폴더에 리포트 2종 자동 저장.
- `ATX WIP.xlsx` 등 다른 파일명 패턴은 이번 구현 범위 밖이다 (설계서 "범위 밖" 참조).
- MO가 비어있다가 나중에 채워지는 랏은 신규+삭제로 이중 표시될 수 있다는 알려진 한계가 있다 (설계서 "알려진 한계" 참조).
