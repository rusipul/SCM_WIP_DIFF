# 저장 폴더 선택 + 디바이스별 단계 수량 표시 설계

## 배경 및 목적

기존 GTK/ATX WIP 일일 비교 도구는 두 가지 제약이 있었다:
1. 변동리포트/하이라이트 파일이 항상 "오늘 파일"과 같은 폴더에 자동 저장되어, 사용자가 원하는 위치에 저장할 수 없었다.
2. GUI 미리보기와 변동리포트의 "요약" 시트가 전공정/후공정/완료 단계별 **전체 합계**만 보여줘서, 어떤 디바이스가 실제로 변동에 기여했는지 알 수 없었다.

이번 작업은 이 두 가지를 개선한다.

## 1. 저장 폴더 선택 기능

- GUI에 "저장 폴더:" 항목(Entry + "찾아보기" 버튼)을 추가한다. "찾아보기"는 `filedialog.askdirectory()`를 사용한다.
- "오늘 파일"을 선택(`choose_today`)할 때마다 저장 폴더 입력란을 그 파일이 있는 폴더로 자동으로 채운다. 사용자가 이후 "찾아보기"로 다른 폴더를 직접 고르면 그 값이 유지되며(오늘 파일을 다시 선택하기 전까지), 다시 오늘 파일을 선택하면 그 폴더로 다시 자동 채워진다.
- "비교 실행" 시 저장 폴더 입력란이 비어있으면(사용자가 지운 경우 등) 오늘 파일이 있는 폴더로 안전하게 대체한다.
- `report.py`의 `derive_output_paths(today_path)`를 `derive_output_paths(today_path, output_folder)`로 변경한다. 파일명 자체(날짜/회사 토큰 추출, `_변동표시`/`_변동리포트` 접미사 등)는 지금처럼 `today_path`의 파일명에서 그대로 뽑되, 폴더 부분만 `os.path.dirname(today_path)` 대신 인자로 받은 `output_folder`를 사용한다.

## 2. 디바이스별 단계 수량 — 데이터 구조

**핵심 이슈**: GTK의 랏 키는 `(MO, 랏번호, 디바이스, 순번)`, ATX의 랏 키는 `(웨이퍼랏, 디바이스, 컨트롤랏, 순번)`으로 "디바이스"가 위치한 인덱스가 서로 다르다(GTK=2번째, ATX=1번째). 이는 이전에 `key_labels` 필드로 헤더 라벨 문제를 해결했던 것과 동일한 종류의 포맷 차이다.

**`ParsedWip`에 필드 추가**: `device_key_index: int` — 키 튜플에서 디바이스명이 몇 번째 위치인지. GTK 파서(`parser.py`)는 `2`, ATX 파서(`atx_parser.py`)는 `1`을 명시적으로 채운다.

**`comparator.py`에 새 함수 추가**:
```python
def compare_stage_summary_by_device(yesterday, today):
    """반환: {디바이스명: {"전공정": {"yesterday":N,"today":N,"delta":N}, "후공정": {...}, ...}}"""
```
- 어제/오늘 각 랏의 키에서 `device_key_index`로 디바이스명을 뽑아 스테이지 컬럼 합계를 디바이스별로 집계한다.
- 어제 또는 오늘 파일에 존재하는 모든 디바이스를 포함한다 (변동이 없는 디바이스도 포함).
- 결과 딕셔너리는 디바이스명 가나다순/알파벳순으로 정렬된 순서로 채운다 (Python 3.7+ dict는 삽입 순서를 유지하므로, 정렬된 순서로 삽입하면 됨).
- `compare_stage_summary`와 동일하게 실제 데이터에 존재하는 단계만 포함한다 (ATX는 "완료" 제외).

## 3. 리포트 / GUI 표시

**`report.py` — `build_variance_report`에 7번째 인자 추가**:
```python
def build_variance_report(
    stage_summary, lot_diff, output_path, column_labels, value_number_format, key_labels, device_summary
):
```
- "요약" 시트의 기존 내용(단계별 합계 표, 빈 줄, 변경/신규/삭제 랏 수) 아래에 빈 줄을 하나 더 넣고 "[디바이스별 단계 수량]" 섹션 헤더를 추가한다.
- 그 아래에 헤더 행 `["디바이스", "{stage1} 어제", "{stage1} 오늘", "{stage1} 증감", "{stage2} 어제", ...]`을 만든다. 반복 대상 단계 목록은 `device_summary`가 아니라 이미 계산되어 있는 `stage_summary`의 키(`STAGE_ORDER` 순서로, 이미 존재하는 단계만 담고 있음)를 그대로 사용한다 — `device_summary`가 비어 있어도(디바이스가 하나도 없는 극단적 경우) 헤더 행 자체는 항상 올바르게 만들어지도록 하기 위함이다. 그 아래에 디바이스별로 한 행씩 채운다.
- 기존 "요약" 시트에 이미 적용된 서식(굵은 헤더, 숫자 천단위 포맷, 열 너비 자동조정, 상단 고정)을 이 섹션에도 동일하게 적용한다. 다만 상단 고정(freeze_panes)은 시트 전체에 하나만 적용되므로 기존 "A2" 고정을 그대로 유지한다.
- `device_summary`가 빈 딕셔너리인 경우(예: 두 파일 모두 랏이 하나도 없는 극단적 상황) 헤더만 쓰고 데이터 행은 없이 넘어간다.

**`gui.py` 변경**:
- 저장 폴더 입력란 추가 (1번 참고).
- `run_compare`가 `compare_stage_summary_by_device(yesterday, today)`도 호출해 `device_summary`를 얻고, `build_variance_report` 호출 시 7번째 인자로 전달한다.
- `_show_preview`에 "[디바이스별 단계 수량]" 섹션을 추가한다. 디바이스마다 한 줄로 `{디바이스명}: {단계1} {어제:,}->{오늘:,}({증감:+,}), {단계2} ...` 형태로 표시한다 (가나다순, `device_summary` 순회 순서 그대로).

## 4. 오류 처리

- 저장 폴더가 존재하지 않거나 쓰기 권한이 없는 경우: 기존과 동일하게 `build_variance_report`/`build_highlighted_today_file` 호출 시 발생하는 예외(`PermissionError`, `FileNotFoundError` 등)를 `gui.py`의 기존 에러 처리 경로(개별 `PermissionError` 처리 + 바깥쪽 `except Exception`)가 그대로 흡수한다. 새로운 예외 처리를 추가하지 않는다.

## 5. 테스트 방침

- `parser.py`/`atx_parser.py`: `device_key_index`가 각각 `2`, `1`로 올바르게 채워지는지 실제 픽스처로 검증하는 단위 테스트 추가.
- `comparator.py`: 합성 데이터로 디바이스별 집계가 정확한지(여러 랏이 같은 디바이스인 경우 합산되는지, 어제/오늘 한쪽에만 있는 디바이스도 포함되는지, 가나다순 정렬되는지) 검증. 실제 GTK/ATX 픽스처로 최소 1개 디바이스의 합계가 독립적으로 재계산한 값과 일치하는지 검증.
- `report.py`: 합성 데이터로 "요약" 시트에 디바이스별 섹션이 올바른 헤더/값으로 추가되는지 검증. `derive_output_paths`는 `output_folder` 인자를 명시적으로 넘겼을 때 그 폴더가 사용되는지 검증 (기존 GTK 테스트들은 폴더 인자만 추가하고 기대값은 동일하게 유지).
- `gui.py`: 기존 정책대로 자동 테스트 대상에서 제외, 수동/기능적 스모크 테스트로 확인.

## 범위 밖 (Out of Scope)

- `column_labels`/`value_number_format`/`key_labels`/`device_summary`를 하나의 객체로 묶는 리팩터링은 이번 작업에 포함하지 않는다 (이전 리뷰에서 지적된 별도의 기술 부채로 남겨둠).
- 저장 폴더 선택값을 앱 재시작 후에도 기억하는 기능(설정 파일 영속화)은 이번 범위에 포함하지 않는다 — 세션 내에서만 동작.
- GUI 미리보기 텍스트 위젯에 스크롤바를 추가하는 것은 이번 범위에 포함하지 않는다 (디바이스 수가 많아 목록이 길어질 수 있으나, tkinter Text 위젯은 스크롤바 없이도 키보드/마우스 휠로 스크롤 가능).
