# ATX WIP 지원 확장 설계

## 배경 및 목적

기존 GTK WIP 일일 비교 도구를 ATX(다른 협력사) WIP 리포트에도 확장 적용한다. 사용자는 동일한 GUI에서 어제/오늘 파일을 선택하면, 파일이 GTK 형식이든 ATX 형식이든 자동으로 인식되어 동일한 방식(변동리포트 엑셀 + 하이라이트된 오늘 파일)으로 결과를 받아볼 수 있어야 한다.

## 대상 파일 구조 (실제 샘플 분석 결과: `260723 ATX WIP.xlsx`, `260724 ATX WIP.xlsx`)

- 파일명 패턴: `<YYMMDD> ATX WIP.xlsx` (GTK와 동일한 날짜 접두사 관례)
- GTK와 달리 **완전히 분리된 시트 4개**로 구성됨: `KSWIPAY`(조립 공정 WIP, 이번 지원 대상), `KSWIPFT (Test)`(테스트 큐/상태 목록), `KSFG (완료)`(완료/출하 로그), `CUST TKW`(고객 전용, 현재 비어있음)
- **시트 이름이 날짜마다 달라짐**: `260723` 파일은 `KSWIPAY (PKG)`, `260724` 파일은 `KSWIPAY` (접미사 없음). 따라서 정확한 이름 매칭이 아니라 **"KSWIPAY로 시작하는 시트"** 접두사 매칭으로 찾아야 한다.
- 대상 시트(`KSWIPAY`) 구조: 1행에 병합 헤더(전공정/후공정 그룹 라벨), 2행에 실제 컬럼명, 3행부터 데이터.
- 컬럼 구성 (2행 기준, 1-indexed):
  - A: RSOD / B: CUSTOMER / C: DEVICE / D: CUSDEVICE / E: WAFERLOT / F: CONTROLLOT / G: LOTTYPE / H: DIEPART / I: STARTTIME / J: PONUMBER / K: BONDINGNOREV / L: PKG_DESCRIPTION
  - M(13)~V(22): 전공정 (UNISSUE, Grind, Saw, Die_Bond, Wire_Bond, 3/O, Q/I, D/C, C/C, DC/I)
  - W(23): FE Total (전공정 소계, 개별 공정 비교 대상 아님)
  - X(24)~AT(46): 후공정 (Ftape, Molding, MarkOut, M/K, Deflash, WJ, PMC, D/D, IRReFlow, Plating, Etching, VMIPlating, P/B, DDTF, Sorting, F/S, ISO, 2D, B/A, Singulation, FVI, FVI QA, Packing)
  - AU(47): BE Total (후공정 소계) / AV(48): Total / AW(49): WO NO / AX(50): PART NO
- **병합 셀 범위가 불안정함**: `260723` 파일은 후공정 병합이 `X1:AT1`(BE Total 제외)인데 `260724` 파일은 `X1:AU1`(BE Total 포함)으로 바뀌어 있음. GTK처럼 병합 셀로 그룹 경계를 판별하면 하루 사이에도 틀릴 수 있으므로, **대신 2행의 헤더 텍스트로 경계를 판별한다**: `UNISSUE` 컬럼부터 `FE Total` 컬럼 직전까지가 전공정, `Ftape` 컬럼부터 `BE Total` 컬럼 직전까지가 후공정.
- 이 표에는 GTK의 "완료" 그룹에 해당하는 컬럼이 없다 (완료/출하는 별도 `KSFG` 시트에서 관리되는 구조로 보임 — 이번 범위 밖).
- **중복 키 이슈**: WAFERLOT+DEVICE+CONTROLLOT이 완전히 동일한 행이 실제로 존재함 (예: 세 행이 A~M열까지 완전히 동일하고 수량 컬럼만 22.98/28.61/34.48로 다름). GTK와 동일하게 **등장 순번을 키에 추가**해 구분한다: `(WAFERLOT, DEVICE, CONTROLLOT, occurrence_index)`.
- 데이터 종료 조건: 3행부터 시작해 A열(RSOD)이 비는 첫 행에서 종료 (GTK와 달리 이후에 다른 표가 이어붙지 않고 시트가 그대로 끝남 — 별도 리포트 블록 제외 로직 불필요).
- **값 타입이 GTK와 다름**: GTK는 정수(예: 24627)였으나 ATX는 소수(예: 22.98)를 사용한다. `int()`가 아닌 `float()`로 저장해야 한다.
- 실제 두 파일로 계산한 결과: 어제(723) 125개 랏, 오늘(724) 141개 랏, 공통 116개, 신규 25개, 삭제 9개, 변경 81개.

## 아키텍처

```
scm_wip_diff/
├── parser.py         # 기존 GTK 파서 (변경 없음, ParsedWip에 필드 2개 추가는 아래 참고)
├── atx_parser.py       # 신규: ATX 파서
├── format_detect.py    # 신규: 파일이 GTK/ATX 중 무엇인지 판별
├── comparator.py       # 변경 없음 (ParsedWip 구조만 소비하므로 그대로 재사용)
├── report.py            # 소폭 변경 (아래 참고)
├── gui.py               # 변경 (판별 로직 연결, 형식 불일치 처리)
└── main.py              # 변경 없음
```

`comparator.py`의 `compare_stage_summary`/`check_lot_overlap`은 `ParsedWip`(column_labels, stage_groups, lots, rows)만 알면 동작하므로 수정이 필요 없다. 다만 실제 코드를 재확인한 결과 `compare_lots`와 `report.py`의 `build_variance_report`는 GTK 전용 상수 `PROCESS_COL_START`/`PROCESS_COL_END`(9~30, GTK의 I~AD)를 하드코딩해서 비교 대상 컬럼 범위를 정하고 있었다. ATX는 컬럼 범위가 다르고(13~22, 24~46) 중간에 소계 컬럼(23열)으로 끊겨 있어 이 상수를 그대로 쓸 수 없다.

이를 바로잡기 위해 두 함수를 다음과 같이 일반화한다:
- `compare_lots`는 비교할 컬럼 목록을 고정 상수 대신 `yesterday.stage_groups`와 `today.stage_groups`에 실제로 포함된 컬럼 인덱스의 합집합에서 동적으로 계산한다 (GTK는 9~30 연속, ATX는 13~22+24~46처럼 중간이 끊겨도 동일하게 동작).
- 이 계산된 컬럼 목록을 `compare_lots`의 반환값에 `process_columns`로 함께 담아, `report.py`가 같은 목록을 다시 계산하지 않고 그대로 받아 쓰도록 한다.
- `report.py`는 `scm_wip_diff.parser`의 `PROCESS_COL_START`/`PROCESS_COL_END` import와 모듈 레벨 `PROCESS_COLS` 상수를 제거하고, `build_variance_report`가 `lot_diff["process_columns"]`를 사용하도록 바꾼다.

이 정리로 `comparator.py`/`report.py`가 애초 의도대로 `ParsedWip` 구조에만 의존하고 GTK 고유 상수에 더 이상 의존하지 않게 된다 (기존 GTK 동작은 결과적으로 동일하게 유지됨 — 9~30 범위가 전공정+후공정+완료 그룹의 합집합과 정확히 일치하기 때문).

또한 `compare_stage_summary`는 현재 `STAGE_ORDER = ["전공정", "후공정", "완료"]`를 무조건 3개 다 순회해서, 데이터에 없는 그룹도 `{"yesterday": 0, "today": 0, "delta": 0}`으로 채워 넣는다. ATX는 "완료" 그룹 자체가 없으므로 이대로 두면 GUI/리포트에 "완료: 0 -> 0 (+0)"처럼 실제로 존재하지 않는 단계가 마치 변동 없는 단계인 것처럼 표시되어 오해를 준다. `compare_stage_summary`가 `yesterday.stage_groups`(와 `today.stage_groups`)에 실제로 존재하는 단계만 결과에 포함하도록 고친다 (GTK는 3개 다 존재하므로 동작 변화 없음). `report.py`의 `build_variance_report`는 이미 `if stage in stage_summary:`로 존재 여부를 확인하고 있어 추가 수정이 필요 없지만, `gui.py`의 `_show_preview`는 `("전공정", "후공정", "완료")`를 고정으로 순회하고 있어 `stage_summary`에 실제로 있는 키만 순회하도록 고쳐야 한다.

## `format_detect.py`

```python
def detect_format(path) -> str  # "GTK" | "ATX"
```
워크북을 열어: 시트 이름 중 `KSWIPAY`로 시작하는 것이 있으면 `"ATX"`. 그렇지 않고 첫 번째 시트에서 GTK의 `find_report_anchor`가 성공하면 `"GTK"` (기존 `parse_wip_sheet`가 항상 첫 번째 시트만 보는 것과 동일한 가정). 둘 다 아니면 `ReportFormatError`.

## `atx_parser.py`

- `KSWIPAY`로 시작하는 시트를 이름 접두사로 찾는다 (정확한 이름이 아님).
- 2행에서 `UNISSUE`, `FE Total`, `Ftape`, `BE Total` 텍스트가 있는 컬럼 위치를 찾아 전공정/후공정 그룹 경계를 계산한다. 넷 중 하나라도 못 찾으면 `ReportFormatError`로 실패한다 (조용히 빈 그룹으로 넘어가지 않음 — GTK에서 이미 적용한 "무음 실패 방지" 원칙과 동일).
- 컬럼 라벨은 2행 텍스트를 그대로 사용한다 (GTK처럼 2줄 헤더 결합이 필요 없음).
- 데이터는 3행부터 A열이 빌 때까지 읽으며, 각 행을 `(WAFERLOT, DEVICE, CONTROLLOT, occurrence_index)` 키로 `ParsedWip.lots`에 저장한다. 값은 `float(v) if isinstance(v, (int, float)) else 0.0`으로 정규화한다.
- `parse_atx_wip_sheet(path)`는 GTK의 `parse_wip_sheet`와 동일하게 `ParsedWip`을 반환한다.

## `ParsedWip` 확장 (GTK/ATX 공용)

```python
@dataclass
class ParsedWip:
    column_labels: dict
    stage_groups: dict
    lots: dict
    rows: dict
    sheet_name: str           # 신규: 실제로 매칭된 시트 이름
    value_number_format: str  # 신규: "#,##0"(GTK, 정수) 또는 "#,##0.00"(ATX, 소수)
```
`build_highlighted_today_file`이 지금은 `wb.sheetnames[0]`(첫 번째 시트)을 무조건 사용하는데, ATX는 시트 이름이 날짜마다 바뀔 뿐 첫 번째 시트인 것은 두 샘플에서 일관되게 유지되지만, 이 가정에 의존하는 것은 위험하다. `sheet_name`을 명시적으로 전달받도록 시그니처를 `build_highlighted_today_file(today_path, lot_diff, output_path, sheet_name)`로 바꾼다. GTK 쪽 호출부(`gui.py`)도 `today.sheet_name`을 전달하도록 함께 수정한다.

`value_number_format`은 `report.py`가 하드코딩된 `"#,##0"` 대신 이 값을 사용해, GTK(정수)는 기존처럼 소수점 없이, ATX(소수)는 `"#,##0.00"`로 표시하도록 한다.

## 파일명 규칙 일반화 (`report.py`)

`derive_output_paths`가 현재 `"GTK"`를 하드코딩하고 있어 ATX 파일에 적용하면 `260723_GTK_WIP_변동리포트.xlsx`처럼 잘못된 이름이 나온다. 파일명에서 날짜와 함께 회사 토큰도 추출하도록 정규식을 `^(\d{6})\s+(\S+)\s+WIP`로 확장한다: 매칭되면 `{날짜}_{회사}_WIP_변동리포트.xlsx`, 매칭 안 되면 기존처럼 파일명 전체를 사용한다. 이 회사 토큰 추출은 표시용 라벨일 뿐이며, GTK/ATX 판별 자체는 여전히 `format_detect.py`의 파일 내용 기준 판별을 사용한다 (파일명에 의존하지 않음).

## GUI 변경 (`gui.py`)

`parse_wip_sheet`를 직접 호출하는 대신, 다음과 같은 순서로 처리하는 작은 dispatcher 함수를 통해 호출한다:
1. `format_detect.detect_format(path)`로 어제/오늘 파일 각각의 형식을 판별
2. 형식에 맞는 파서(`parser.parse_wip_sheet` 또는 `atx_parser.parse_atx_wip_sheet`) 호출
3. 어제/오늘 형식이 다르면 비교를 중단하고 안내: "어제 파일은 {A} 형식, 오늘 파일은 {B} 형식입니다. 같은 회사 파일끼리 비교해주세요."

## 오류 처리

- 파일이 GTK/ATX 어느 쪽으로도 판별되지 않으면 `ReportFormatError`로 안내 후 중단
- 어제/오늘 파일 형식 불일치 시 위 GUI 섹션의 전용 메시지로 중단
- ATX 표에서 필수 헤더 라벨(`UNISSUE`, `FE Total`, `Ftape`, `BE Total`)을 찾지 못하면 `ReportFormatError`

## 테스트 방침

- `tests/fixtures/`에 `260723 ATX WIP.xlsx`, `260724 ATX WIP.xlsx` 추가
- `atx_parser.py`: 실제 픽스처로 시트 접두사 탐지, 헤더 라벨 기반 그룹 판별, 등장순번 중복 키 처리, 데이터 종료 조건, float 값 추출을 각각 단위 테스트
- `format_detect.py`: GTK 파일 → `"GTK"`, ATX 파일 → `"ATX"`, 빈 워크북 → `ReportFormatError`
- `comparator.py`/`report.py`: 기존 테스트는 수정 없이 그대로 통과해야 함 (회귀 확인). ATX 실제 픽스처를 사용한 통합 테스트를 추가해 신규 25건/삭제 9건/변경 81건을 검증
- `gui.py`: 형식 불일치 시 에러 메시지 노출 여부는 GTK와 동일하게 수동 스모크 테스트로 확인 (자동 테스트 대상 아님)

## 범위 밖 (Out of Scope)

- ATX의 `KSWIPFT (Test)`(테스트 큐/상태 목록), `KSFG (완료)`(완료/출하 로그), `CUST TKW`(고객 전용, 현재 데이터 없음) 시트는 이번 작업에 포함하지 않는다. 세 시트 모두 구조가 서로 다르고 "랏별 공정 컬럼 수량"과는 다른 성격의 데이터(상태값, 로그성 기록)라 별도 설계가 필요하다.
- ATX의 "완료" 단계 추적은 이번 범위에 포함하지 않는다 (KSWIPAY 표 자체에 완료 개념이 없음).
