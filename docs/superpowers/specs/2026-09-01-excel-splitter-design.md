# Excel Table 분할기 설계

## 1. 목적

Windows에서 실행되는 Python 기반 GUI 프로그램을 만든다. 사용자는 `.xlsx` 통합문서에서 시트 하나와 그 시트에 있는 정식 Excel Table(`ListObject`)의 컬럼 하나를 선택한다. 프로그램은 선택 컬럼의 고유값마다 결과 통합문서를 하나씩 생성한다. 결과 통합문서에는 선택한 시트만 남고, 선택한 Table에는 해당 고유값을 가진 데이터 행만 남는다.

이 문서에서 “원본 보존”은 바이너리 동일성을 뜻하지 않는다. 결과 파일에서 의도적으로 제거되는 다른 시트, 대상이 아닌 Table 데이터 행, 기존 Table 필터 조건을 제외하고, 선택 시트의 자체 포함 요소를 데스크톱 Excel이 저장하는 범위에서 유지한다는 뜻이다. 삭제한 시트를 참조하던 요소의 손상은 별도 계약으로 다룬다.

## 2. 범위

### 포함

- Windows 데스크톱 GUI
- `.xlsx` 입력 및 `.xlsx` 출력
- 원본 통합문서의 워크시트 목록 표시
- 선택 워크시트의 정식 Excel Table 개수 검사
- Table 컬럼 목록 표시
- 선택 컬럼의 고유값 및 값별 데이터 행 개수 표시
- 빈 셀을 독립된 분류로 처리
- 고유값별 결과 통합문서 생성
- 출력 폴더 선택, 파일명 미리보기, 기존 파일 덮어쓰기 확인
- Excel에서 열 수 있는 DRM 보호 `.xlsx` 원본
- DRM이 제거된 일반 `.xlsx` 결과 파일
- 단일 `.exe` 배포

### 제외

- `.xls`, `.xlsm`, `.xlsb` 및 CSV
- Excel에서 해제할 수 없는 DRM 또는 열기 암호·편집 암호가 필요한 파일
- 한 시트에 Table이 없거나 둘 이상인 경우의 분할
- 여러 Table 또는 여러 컬럼을 조합한 분류
- 외부 데이터, SharePoint, QueryTable 또는 Data Model에 연결된 Table의 분할
- Microsoft Excel이 설치되지 않은 PC에서의 실제 분할
- Excel Viewer만을 이용한 분할
- 원본 파일의 직접 수정

## 3. 실행 환경과 기술 선택

- GUI: Python 표준 라이브러리 `tkinter`
- Excel 연동: `pywin32`의 COM 자동화
- 패키징: `PyInstaller`의 단일 창 실행 파일
- 대상 운영체제: 64비트 Windows 10 또는 Windows 11
- 실행 전제: 정식 데스크톱 Microsoft Excel 2016 이상 설치
- Excel 비트: 32비트 또는 64비트

프로그램은 `DispatchEx("Excel.Application")`로 별도 Excel COM 인스턴스를 생성한다. 사용자가 이미 실행 중인 Excel 인스턴스를 조회하지 않으며, 프로그램이 만든 인스턴스만 종료한다. Excel 연동은 late binding을 사용하여 개발 PC에서 생성된 COM 캐시에 의존하지 않게 한다.

`openpyxl`은 테스트용 `.xlsx` fixture 생성에는 사용할 수 있지만 실제 결과 생성에는 사용하지 않는다. 결과 생성은 데스크톱 Excel이 파일을 열고 수정하고 저장하게 하여 Excel 문서 요소의 손실 위험을 줄인다.

### 3.1 DRM 입력과 읽기 성능

DRM 원본은 파일 시스템에서 ZIP/XML로 직접 읽지 않는다. GUI는 앱 수명 동안 하나의 전용 COM command worker를 유지하며, 이 worker만 source `Workbook`과 source용 `Excel.Application`을 소유한다. worker는 GUI 시작 후 Excel 인스턴스를 미리 생성하지만 DRM 원본은 사용자가 선택한 뒤에만 연다.

원본 선택 시 파일 서명을 기록하고 `Workbooks.Open(ReadOnly=True)`을 정확히 한 번 호출한다. DRM이 Excel에서 해제된 열린 통합문서를 다음 단계가 재사용한다.

- 원본 선택은 열린 통합문서에서 워크시트 이름만 읽는다.
- 워크시트 선택은 Table 개수, 종류, 행 수 및 컬럼 이름만 검사한다. Table 아래 콘텐츠 검사는 이 단계에서 실행하지 않는다.
- 미리보기는 같은 통합문서에서 전체 안전성 검사를 한 번 실행하고 분류 스냅샷을 만든다.
- source 경로가 바뀌거나 디스크 서명이 달라지거나 COM 신뢰를 잃으면 열린 통합문서를 닫고 선택 상태를 무효화한다.
- GUI 종료 시 command worker가 source 통합문서, source Excel 인스턴스, COM apartment 순서로 정리된다. COM proxy는 이 worker 밖으로 전달하지 않는다.

원본 선택 시 SHA-256·크기·수정 시각을 모두 기록한다. 워크시트 선택과 미리보기에서는 같은 열린 Workbook을 사용하면서 크기와 수정 시각만 빠르게 재검사하고, 실제 분할의 `SaveAs` 직전에 SHA-256까지 다시 계산한다. GUI가 표시된 직후 source Excel 인스턴스는 백그라운드에서 미리 준비하여 첫 원본 선택의 `DispatchEx` 시작 비용을 숨긴다.

DRM 해제와 최초 `Workbooks.Open` 비용은 피할 수 없지만 한 workflow에서 반복하지 않는다. source Excel 인스턴스는 사용자가 실행한 Excel과 분리하며, 프로그램은 자신이 만든 인스턴스만 종료한다.

### 3.2 미리보기 데이터 읽기

분류 열은 `ListColumn.DataBodyRange.Value2`를 한 번 읽어 Python 값 배열로 분리한다. 단일 셀일 때의 scalar와 여러 셀일 때의 2차원 배열을 동일한 행 sequence로 정규화한다. canonical key, 최초 등장 순서 및 원본 1-based `ListRow` 인덱스는 기존 계약을 유지한다.

셀의 표시 문자열 `Text`는 모든 행에서 읽지 않는다. `Value2`로 canonical key를 만든 뒤 각 key가 처음 등장한 행에서만 `Text`를 읽는다. 오류, bool, 숫자, 텍스트, 빈값의 fallback label 규칙은 기존 classifier와 동일하다.

Table 아래 범위 검사는 미리보기에서 한 번만 수행한다. 값·수식·병합·하이퍼링크처럼 range 단위로 판정 가능한 속성은 먼저 bulk COM 호출로 확인한다. 혼합 Style, legacy note 및 threaded comment처럼 bulk 결과가 모호한 항목만 제한적으로 셀 단위 검사한다. Excel 버전이 지원하지 않는 `CommentThreaded` 속성은 직접 호출하지 않으며, 지원되는 collection 또는 명시적인 미지원 결과로 처리한다.

### 3.3 비보호 master와 병렬 출력

분할 실행 직전에 디스크의 source 서명을 미리보기 서명과 다시 비교한다. 일치할 때 source용 command worker가 열린 DRM 통합문서를 출력 폴더의 UUID 임시 경로에 Excel `.xlsx` 형식으로 `SaveAs`한다. source 경로에는 저장하지 않는다.

저장된 master는 다음 조건을 모두 만족해야 한다.

- 일반 ZIP 기반 `.xlsx` package로 열 수 있다.
- 필수 workbook package entry를 읽을 수 있다.
- source 스냅샷의 워크시트, Table, 컬럼 및 행 수 식별자와 일치한다.
- 열기 암호와 쓰기 암호가 없다.

검사에 실패하면 master를 삭제하고 결과 생성을 시작하지 않는다. 이 프로그램에서 “비보호”는 일반 ZIP 기반 `.xlsx`이고 파일 열기·쓰기 암호 없이 Excel로 다시 열 수 있다는 운영 기준이다. DRM 제품이나 조직 정책이 이 기준의 파일 생성을 막으면 프로그램이 이를 우회하지 않고 “비보호 master를 만들 수 없습니다”라고 알린다. source 통합문서는 master 생성 성공 또는 실패 후 저장하지 않고 닫는다.

비보호 master가 준비되면 coordinator는 기본 두 개의 output worker를 실행한다. 각 worker는 자기 thread에서 COM을 초기화하고 별도 `DispatchEx("Excel.Application")` 인스턴스를 만들며, 할당된 target을 순차 처리한 뒤 같은 thread에서 `Quit`과 COM 해제를 수행한다. master와 target 경로만 worker에 전달하고 COM proxy는 공유하지 않는다.

- 분류가 두 개 이하이면 worker 한 개를 사용한다.
- 분류가 세 개 이상이면 worker 두 개를 사용한다.
- 각 target은 master의 별도 UUID 복사본을 열고 서로 다른 최종 경로에 배치한다.
- 제외 행 수가 많은 target부터 worker에 배정하여 마지막 작업 편중을 줄인다.
- 완료 결과는 미리보기 target 순서로 재정렬한다.
- progress callback은 coordinator만 호출하며 완료 수는 0부터 전체 수까지 단조 증가한다.

개별 target의 파일 복사·배치 `OSError`는 해당 분류 실패로 기록하고 다른 target을 계속한다. COM identity, open, save 또는 close 신뢰 오류는 전역 stop을 설정하여 새 target 배정을 중단한다. 이미 Excel에서 편집 중인 target은 강제 취소하지 않고 닫기와 임시 파일 정리를 마친다. 완료된 파일, 개별 실패 및 시작하지 못한 target을 포함한 부분 결과를 GUI에 표시한다.

행 삭제는 원본 인덱스를 기준으로 내림차순 처리한다. 연속된 제외 인덱스는 하나의 contiguous block으로 묶어 Excel 호출 수를 줄일 수 있지만, 정식 Excel 통합 테스트에서 Table 크기·아래 콘텐츠·도형 보존이 확인된 경우에만 block 삭제를 활성화한다. 검증 전 기본 동작은 행별 `ListRow.Delete`다.

### 3.4 성능 관측

로그에는 시트 조회, Table 조회, bulk 스냅샷 생성과 전체 분할의 elapsed time을 기록한다. 워크시트 수, Table 행 수, 컬럼 수, 유니크 분류 수와 성공·실패 수를 함께 기록한다. 로그에는 셀 값, 분류 label 및 파일 암호를 남기지 않는다.

## 4. 사용자 흐름

GUI는 한 화면에서 위에서 아래로 진행한다.

1. 사용자가 원본 `.xlsx` 파일을 선택한다.
2. 프로그램이 통합문서의 워크시트 목록을 표시한다.
3. 사용자가 워크시트를 선택한다.
4. 프로그램이 선택 워크시트의 `ListObjects.Count`를 검사한다.
   - 0개이면 “정식 Excel Table이 없습니다”를 알리고 이후 입력을 비활성화한다.
   - 2개 이상이면 “이 시트에는 Table이 2개 이상 있어 처리할 수 없습니다”를 알리고 이후 입력을 비활성화한다.
   - 정확히 1개이면 그 Table의 컬럼 목록을 표시한다.
5. 사용자가 컬럼 하나를 선택한다.
6. 프로그램이 고유값과 값별 행 개수를 표 형태로 표시한다. 빈 셀은 GUI에서 `∅ (빈 셀)`로 표시한다.
7. 사용자가 파일명 패턴과 출력 폴더를 지정한다. 패턴 기본값은 `%_분할`이고 출력 폴더 기본값은 원본 파일의 폴더다.
8. 프로그램이 생성될 파일명 목록을 미리 보여준다.
9. 기존 파일과 충돌하면 전체 충돌 목록을 표시하고 덮어쓰기 승인을 한 번 받는다.
10. 사용자가 분할을 실행하면 완료 수 기반 진행률과 가장 최근 완료된 분류를 표시한다.
11. 완료 후 성공 파일과 실패 원인을 요약한다.

파일을 다시 선택하거나 시트를 변경하면 그 아래 단계의 상태와 미리보기를 모두 초기화한다. Excel 작업은 백그라운드 worker thread에서 실행하고 GUI 갱신은 Tkinter의 메인 스레드에서만 수행한다. 분할 중에는 모든 입력과 창 닫기를 비활성화하고 취소 기능은 제공하지 않는다. 현재 결과 파일의 중간 상태를 안전하게 되돌릴 수 없는 COM 작업을 임의로 중단하지 않기 위한 정책이다.

## 5. 입력 검사

- 확장자는 대소문자와 관계없이 `.xlsx`여야 한다.
- 원본 경로가 존재하고 일반 파일이어야 한다.
- 원본 경로와 출력 폴더에 Excel이 경로에서 지원하지 않는 대괄호 `[` 또는 `]`가 있으면 경로 변경을 요청한다.
- Excel이 사용자 권한으로 해제할 수 있는 DRM 원본만 지원한다. 열기 암호 또는 편집 암호가 필요한 통합문서는 지원하지 않는다.
- 통합문서 구조가 보호되어 다른 시트를 삭제할 수 없으면 분할을 시작하지 않는다.
- 선택 워크시트 또는 Table이 보호되어 데이터 행을 삭제할 수 없으면 분할을 시작하지 않는다.
- 선택 워크시트가 숨김 또는 매우 숨김 상태이면 분할을 시작하지 않고 이유를 표시한다. 결과 통합문서에는 이 워크시트만 남아야 하므로 최소 한 개의 표시 워크시트가 필요하다.
- Table의 `SourceType`이 일반 범위인 `xlSrcRange`가 아니거나 외부 연결이 있으면 분할을 시작하지 않는다.
- Table 열 범위의 바로 아래부터 워크시트의 마지막 사용 행까지 독립 셀 값, 수식, 메모, 하이퍼링크, 병합 영역 또는 기본값이 아닌 셀 스타일이 있으면 분할을 시작하지 않는다. `ListRow.Delete`가 아래 셀을 위로 이동시키기 때문이다.
- Table에 데이터 행이 없거나 선택 컬럼에 분류할 행이 없으면 결과 파일을 만들지 않고 안내한다.
- Excel 데스크톱 앱을 시작할 수 없으면 설치 필요 메시지를 표시한다.

파일 열기와 저장 과정에서만 발견할 수 있는 COM 오류를 제외한 모든 검사는 결과 파일을 만들기 전에 끝낸다. 읽기 검사에는 원본을 읽기 전용으로 열고 외부 링크를 갱신하지 않는다.

## 6. 분류 스냅샷과 값 규칙

컬럼을 선택한 시점에 원본의 SHA-256, 파일 크기, 수정 시각, 워크시트·Table·컬럼 식별자, 전체 `ListRows.Count`와 각 원본 행의 분류 키를 스냅샷으로 저장한다. 각 행 기록은 `(원본 ListRow 인덱스, canonical key, 표시 문자열)`이다. 실행 직전에 원본의 세 파일 식별값을 다시 계산하며 하나라도 바뀌면 미리보기를 폐기하고 재검사를 요구한다.

분류 소속은 이 최초 스냅샷으로 고정한다. 결과 파일에서 다른 시트를 삭제하거나 Table 행을 삭제하여 수식 결과가 달라져도 분류 키를 다시 계산하지 않는다. 결과마다 스냅샷에서 제거 대상의 원본 행 인덱스를 구하고 내림차순으로 삭제한다.

canonical key는 COM `Value2`의 계산 결과와 오류 여부를 다음 순서로 정규화한다.

| 셀 상태 | canonical key | 예시 |
| --- | --- | --- |
| 실제 빈 셀 또는 수식 결과 `""` | `("blank",)` | 두 상태는 같은 분류 |
| Excel 오류 | `("error", 오류 코드)` | `#N/A`와 `#VALUE!`는 다른 분류 |
| 논리값 | `("bool", true/false)` | 숫자 `1`과 다름 |
| 숫자 | `("number", 정규화된 숫자)` | `1`과 `1.0`은 같음 |
| 텍스트 | `("text", 원문)` | 대소문자와 앞뒤 공백을 구분 |

날짜와 시간은 `Value2`의 Excel 직렬 숫자로 분류하므로 같은 직렬값은 셀 표시 형식이 달라도 같은 분류다. 표시 형식 때문에 빈 것처럼 보이는 숫자는 빈 분류가 아니다. 표의 계산 열과 수식 셀도 스냅샷 시점의 계산 결과를 사용한다. 그 밖의 COM 값 형식은 추측해 변환하지 않고 지원하지 않는 값 오류로 처리한다.

GUI와 파일명의 대표 문자열은 각 분류가 처음 등장한 셀의 `Text`를 사용한다. `Text`가 `####`이거나 비어 있지만 blank가 아닌 경우에는 텍스트 원문, Excel 오류 표기, `TRUE`/`FALSE`, 또는 숫자 15자리 유효숫자 문자열 순으로 대체한다. 집계와 파일 생성 순서는 원본에서 처음 등장한 순서다. 현재 Table 필터는 집계에서 무시하며 모든 `ListRows`를 대상으로 한다.

## 7. 파일명 규칙

사용자가 입력하는 패턴에서 `%`는 분류 값 자리표시자다. 파일명은 다음 순서로 결정한다.

1. 패턴에 `%`가 하나 이상 있는지 검사한다. 패턴이 대소문자와 무관하게 `.xlsx`로 끝나면 확장자를 떼어 내고, `.xls`, `.xlsm`, `.xlsb`로 끝나면 지원하지 않는 확장자 오류로 처리한다. 그 밖의 마침표는 이름의 일부다.
2. 모든 `%`를 대표 문자열로 치환한다. 빈 분류는 빈 문자열로 치환한다.
3. Windows 또는 Excel에서 문제가 되는 `\ / : * ? " < > | [ ]`를 `_`로 치환하고 `.xlsx`를 붙인다.
4. 파일명 끝의 공백과 마침표를 제거한다.
5. 확장자를 제외한 이름이 `CON`, `PRN`, `AUX`, `NUL`, `COM1`~`COM9`, `LPT1`~`LPT9`와 대소문자 무관하게 같으면 앞에 `_`를 붙인다.
6. 정리 후 이름이 비거나 확장자만 남으면 패턴 오류로 처리한다.
7. Windows의 대소문자 무관 비교로 이름이 겹치면 두 번째부터 확장자 앞에 ` (2)`, ` (3)`을 붙인다.
8. 접미사까지 적용한 전체 절대 경로 또는 run 디렉터리 안의 Excel 임시 경로가 218자를 넘으면 패턴 또는 출력 폴더 수정을 요청한다.
9. 최종 경로가 원본 경로와 대소문자 무관하게 같으면 실행을 차단한다.

예를 들어 `%_가나다`와 값 `데이터1`은 `데이터1_가나다.xlsx`, 빈 분류는 `_가나다.xlsx`가 된다. 오류값과 서로 다른 표시 문자열이 문자 정리 후 같은 이름이 되더라도 7단계에서 충돌을 해소한다.

기존 결과 파일이 하나라도 있으면 분할 전에 전체 목록을 보여준다. 사용자가 승인한 경우에만 교체하며, 승인하지 않으면 어떤 결과 파일도 만들지 않는다. 승인 시 기존 충돌 파일의 크기, 수정 시각과 SHA-256을 기록한다. 게시 시 신규 target은 Windows의 no-clobber `os.rename`으로 생성한다. 승인된 기존 target은 같은 폴더의 고유 recovery 경로로 먼저 원자적으로 이동하고 그 파일의 서명을 다시 확인한 뒤, 임시 결과를 no-clobber rename으로 게시한다. 서명이 다르거나 제3자가 target을 선점하면 어느 파일도 덮지 않으며, 자동 복구할 수 없는 경우 recovery 경로를 오류에 표시한다.

## 8. 결과 생성 알고리즘

분할 작업은 다음 순서로 실행한다.

1. source command worker가 원본의 SHA-256, 크기, 수정 시각을 미리보기 스냅샷과 대조한다.
2. 열린 DRM 원본에서 외부 링크·연결을 갱신하지 않고 이벤트를 비활성화한 상태로 UUID master 경로에 `FileFormat=51`, 빈 `Password`와 빈 `WriteResPassword`, `ReadOnlyRecommended=False`, `AddToMru=False`를 지정하여 `SaveAs`한다.
3. master가 일반 ZIP package이고 `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`를 읽을 수 있는지 검사한다. 손상된 ZIP entry가 하나라도 있으면 실패한다.
4. master를 Excel에서 빈 암호로 다시 열 수 있고 워크시트 이름, 선택 Table 이름, 선택 컬럼 이름과 `ListRows.Count`가 스냅샷과 같은지 검사한다. 분류값은 재분류하지 않는다. 같은 열린 source에서 저장한 master이므로 행 순서는 그대로여야 하며, 불일치하면 전체 작업을 중단한다.
5. source 통합문서를 저장하지 않고 닫는다. source Excel 인스턴스는 다음 workflow의 사전 준비 비용을 줄이기 위해 앱 종료 때까지 유지한다.
6. target이 없으면 output worker를 만들지 않는다. target이 1~2개면 worker 한 개, 3개 이상이면 worker 두 개를 만든다.
7. coordinator가 `row_count - group.count`가 큰 target부터 동적 queue에 넣으며 동률은 원래 target 순서를 따른다. stop 확인과 target claim은 같은 scheduling lock 안에서 수행한다.
8. 각 output worker는 자기 thread에서 COM을 초기화하고 `DispatchEx`로 숨겨진 별도 Excel 인스턴스를 시작한다. `Visible=False`, `DisplayAlerts=False`, `EnableEvents=False`, `ScreenUpdating=False`, `AskToUpdateLinks=False`를 설정한다.
9. worker는 stop event를 확인한 뒤 target 하나를 가져와 master를 최종 파일과 같은 출력 volume의 run 전용 임시 디렉터리에 UUID `.xlsx`로 복사한다.
10. 임시 복사본을 쓰기 가능 상태로 열고 워크시트·Table·컬럼 식별자와 `ListRows.Count`를 스냅샷과 대조한다.
11. 기존 Table 필터 조건을 해제한다. 헤더의 필터 기능은 유지하되 결과 파일을 열면 남은 모든 행이 보이게 한다.
12. 스냅샷에서 현재 분류가 아닌 원본 행 인덱스를 구하고 `ListRows`를 인덱스 내림차순으로 삭제한다. 삭제 중에는 셀 값을 다시 읽어 분류하지 않는다.
13. 선택 워크시트를 제외한 모든 `Workbook.Sheets` 항목을 삭제한다.
14. `ListRows.Count`가 스냅샷의 해당 분류 행 수와 같고 선택 시트만 남았는지 확인한다.
15. 통합문서를 저장하고 닫은 뒤 신규 target은 no-clobber `os.rename`으로 게시한다. 기존 target은 고유 recovery 경로로 원자 claim하고 recovery의 승인 시점 서명을 재확인한 다음 임시 결과를 no-clobber rename으로 게시한다. 실패 시 target이 비어 있을 때만 recovery를 되돌리며, 이미 점유되었으면 recovery 파일을 보존하고 경로를 보고한다.
16. worker는 target 결과를 coordinator queue에 전달한다. coordinator만 완료 progress를 올리며 성공, 개별 실패와 fatal을 발생시킨 target은 완료 수에 포함한다. fatal 시 이미 실행 중인 target의 후속 결과도 수집하고, 시작하지 못한 target은 `unstarted`로 분리하여 완료 수에 포함하지 않는다.
17. 개별 파일 `OSError`는 다음 target을 계속한다. workbook identity, open, save 또는 close 신뢰 오류가 발생하면 stop event를 설정하며 해당 target은 재시도하지 않는다.
18. 모든 in-flight target이 종료되면 coordinator가 worker를 join하고 결과를 원래 target 순서로 정렬한다.
19. master, run 전용 임시 디렉터리와 그 안의 target 임시 파일은 성공·개별 실패·전역 중단 모두 `finally`에서 삭제한다. 삭제는 짧게 재시도하며, 끝내 실패하면 원래 처리 오류와 평문 잔존 경로를 함께 표시한다. 사용자 target을 claim한 recovery 파일은 run 디렉터리 밖에 두어 자동 복구가 불가능한 충돌이나 비정상 종료에서도 사용자 데이터가 삭제되지 않게 한다.

역순 인덱스 삭제는 행 삭제에 따른 인덱스 이동을 피하며, 수식 재계산이나 삭제 시트 참조 손상과 관계없이 최초 미리보기의 분류 소속을 유지한다. Excel의 저장·재계산 정책은 원본 통합문서 설정을 따른다. 프로그램은 외부 링크나 연결의 새로 고침을 시작하지 않는다.

## 9. 보존 계약과 장애 안전성

| 대상 | 계약 |
| --- | --- |
| 원본 파일 | 저장하지 않으며 작업 전후 SHA-256과 수정 시각이 같아야 한다. |
| 결과의 시트 | 선택 워크시트 하나만 남는다. 차트 시트를 포함한 나머지 `Sheets`는 삭제한다. |
| Table | 헤더, 이름, 스타일, 합계 행과 계산 열 정의를 유지한다. 현재 필터 조건만 의도적으로 해제하고 해당 분류의 행을 원본 순서로 남긴다. |
| 선택 시트의 자체 포함 요소 | 제거된 Table 행에 속하지 않는 셀 값·수식·서식, 조건부 서식, 데이터 유효성 검사, 병합, 열 너비, 행 높이, 도형·차트의 내용과 위치, 인쇄 설정을 유지한다. |
| 삭제 시트 의존 요소 | 삭제 시트를 참조하는 수식, 이름, 차트 계열, 유효성 검사, 피벗 및 연결은 Excel이 `#REF!` 또는 연결 손실로 바꿀 수 있다. 이는 “선택 시트만 유지”의 불가피한 결과로 허용하며 실행 전 고정 경고를 표시한다. |
| Table 의존 수식 | 행 제거에 따른 계산 결과 변경은 허용하지만 수식 정의는 Excel이 행 삭제에 맞게 조정한 결과를 유지한다. |
| Table 아래 콘텐츠 | 입력 검사에서 독립 사용 셀을 차단한다. 도형과 차트의 위치·크기는 삭제 전 스냅샷과 대조하고 달라졌으면 복원한 뒤 저장한다. |

각 결과 파일의 배치는 원자적이지만 전체 배치는 원자적이지 않다. 파일별 처리 실패 시 그 임시 파일을 삭제하고 다음 분류를 계속한다. COM 인스턴스 종료, 통합문서 손상 의심, 원본 스냅샷 불일치처럼 Excel 세션의 신뢰를 잃는 오류는 새 작업 배정을 중단한다. 이미 실행 중인 다른 worker는 강제 중단하지 않고 정리하며, 완성된 결과는 롤백하지 않고 성공 목록에 남긴다. 비보호 master와 run 임시 파일은 결과 파일과 달리 영구 산출물이 아니며 모든 정상·예외 경로에서 삭제 대상이다.

기존 파일 교체 실패 시 기존 파일을 유지해야 한다. `try/finally`에서 열린 통합문서 닫기, Excel 종료, COM 해제, COM thread 종료를 수행한다. 사용자 Excel 인스턴스, 열려 있는 다른 통합문서, 클립보드는 사용하지 않는다. 완료 후 성공 파일 수, 실패한 분류와 원인, 출력 폴더를 표시하며 일부 성공을 전체 성공으로 표현하지 않는다.

## 10. 오류 처리

사용자 메시지는 기술 예외 대신 해결 가능한 원인을 우선 표시한다.

- Excel 미설치 또는 시작 실패
- 원본 파일 없음, 접근 권한 없음, 다른 프로그램에 의한 잠금
- 지원하지 않는 확장자, Excel에서 해제할 수 없는 DRM 또는 암호가 필요한 통합문서
- DRM 정책 때문에 비보호 master를 만들 수 없는 통합문서
- Table 0개 또는 2개 이상
- 외부 데이터에 연결된 Table 또는 Table 아래 독립 콘텐츠
- 보호된 통합문서 구조, 워크시트 또는 Table
- 숨겨진 대상 워크시트
- 유효하지 않은 파일명 패턴 또는 지나치게 긴 경로
- 출력 파일 잠금 또는 덮어쓰기 실패
- COM 호출 실패, Excel 저장 실패

예상 가능한 오류는 도메인 예외로 변환하여 GUI에 간결하게 표시한다. 진단 로그는 `%LOCALAPPDATA%\ExcelSplitter\logs\excel-splitter.log`에 최대 1MB 파일 3개로 회전 저장한다. 원본·출력 경로와 기술 예외는 기록하지만 셀 값, 고유값 목록, 수식 내용은 기록하지 않는다.

## 11. 모듈 구조

```text
src/excel_splitter/
  __init__.py
  app.py              # 진입점과 의존성 조립
  controller.py       # GUI 상태와 service 호출 조율
  gui.py              # Tkinter 화면 및 UI 상태 전이
  models.py           # 워크시트, Table, 컬럼, 분류, 결과 모델
  ports.py            # gateway/service/progress 인터페이스
  source_session.py   # 장수명 source COM worker와 DRM 원본 수명주기
  snapshot_reader.py  # bulk Value2 정규화와 분류 스냅샷 생성
  parallel_writer.py  # 최대 두 output worker의 queue, progress와 부분 결과 조율
  excel_gateway.py    # COM 검사, 비보호 master 요청, 복사본 편집 및 정리
  classifier.py       # 값 정규화, 고유값 집계, 행 비교 키
  naming.py           # 패턴 검증, 파일명 정리, 충돌 해소
  split_service.py    # 사전 검사와 분할 작업 조율
  errors.py           # 사용자 표시 가능한 도메인 오류
tests/
  unit/
  integration/
scripts/
  build.ps1
  smoke_test.ps1
```

GUI는 COM 객체를 직접 다루지 않는다. `split_service`는 인터페이스를 통해 Excel gateway를 사용하므로 단위 테스트에서는 가짜 gateway를 주입할 수 있다. COM 객체는 생성된 worker thread 밖으로 전달하지 않는다.

## 12. 테스트 전략

### 개발 PC 자동 테스트

현재 개발 PC에는 Excel Viewer만 있으므로 실제 Excel COM 통합 테스트는 실행하지 않는다.

- 파일명 패턴 치환, 빈값 치환, 금지 문자와 예약 이름 정리
- 파일명 및 원본 경로 충돌 처리
- canonical key와 고유값 개수 집계
- 빈 셀과 빈 문자열 수식 결과의 통합
- 숫자 `1`/`1.0`, 논리값/숫자, 날짜 직렬값, 대소문자, 공백, 오류 코드 경계 사례
- 스냅샷 행 인덱스 기반 역순 삭제 계획과 원본 변경 감지
- Table 개수와 보호 상태 검사 결과 처리
- 가짜 Excel gateway를 이용한 성공, 부분 실패, 정리 동작
- source command queue의 thread affinity, 단일 open, source 변경 무효화와 종료 정리
- bulk `Value2`의 scalar/2차원 배열 정규화와 유니크 key별 최초 `Text` 조회
- 비보호 master ZIP 검증, 필수 package entry, SaveAs 실패와 임시 파일 정리
- worker 수 0/1/2 경계, 최대 동시성, 동적 queue, stop event와 결과 순서 복원
- out-of-order 완료에서도 단조 progress, 부분 성공과 미시작 target 요약
- GUI 상태 전이와 worker 이벤트 처리

### 정식 Excel PC 통합 테스트

고정 fixture 통합문서로 다음을 검증한다.

- 원본에는 여러 워크시트가 있고 선택 시트에는 Table이 정확히 한 개 있다.
- 결과에는 선택 시트만 존재한다.
- 각 결과 Table에는 해당 분류의 행만 존재한다.
- 빈 셀 분류 파일이 생성된다.
- 타 시트 참조 수식, 행 삭제에 따라 값이 변하는 수식, 계산 열과 volatile 수식도 최초 행 스냅샷의 분류를 유지한다.
- 자체 포함 셀 값·수식·서식, 조건부 서식, 데이터 유효성 검사, 열 너비, 행 높이, 병합, 이름 정의, 도형·차트 위치와 인쇄 설정이 보존 계약과 일치한다.
- 삭제 시트 의존 요소는 허용된 손상으로만 달라지고 외부 링크는 자동 갱신되지 않는다.
- 일반 Table 아래 콘텐츠가 있는 입력과 외부 연결 Table을 사전 차단한다.
- 기존 필터 상태와 관계없이 전체 행을 분류하고 결과에서는 모든 잔존 행이 보인다.
- 기존 파일 충돌 승인과 거절이 각각 동작한다.
- 승인 후 새로 생긴 충돌 파일을 덮어쓰지 않는다.
- 개별 파일 실패는 다음 분류를 계속하고 COM 세션 실패는 전체 작업을 중단한다.
- 실패 후 프로그램이 시작한 Excel 프로세스가 남지 않는다.
- smoke test가 기록한 원본 파일의 해시와 수정 시간이 작업 전후로 동일하다.
- DRM 원본은 source Excel에서 한 번만 열리고 워크시트·컬럼·미리보기가 같은 열린 통합문서를 재사용한다.
- master와 결과 파일은 DRM이 없는 일반 `.xlsx`로 열리며 master와 임시 파일은 성공·실패 후 남지 않는다.
- worker 두 개가 서로 다른 Excel 인스턴스에서 동시에 target을 처리하고 사용자 Excel에는 영향이 없다.
- 직렬/병렬 결과의 시트, Table 행, 보존 요소와 파일명이 동일하다.

실제 Excel 통합 테스트를 통과하기 전에는 “원본 보존 검증 완료”로 표시하지 않는다.

## 13. 패키징과 배포

`PyInstaller`로 콘솔 없는 단일 실행 파일 `ExcelSplitter.exe`를 만든다. 먼저 one-folder 빌드로 의존성을 검증한 뒤 one-file 빌드를 만든다. 빌드 스크립트는 고정된 진입점, 고정된 의존성 버전과 필요한 숨은 import를 선언한다. “반복 가능한 빌드”는 같은 소스와 잠금 파일에서 기능이 같은 실행 파일을 다시 만들 수 있다는 뜻이며 바이너리 해시 동일성을 요구하지 않는다.

배포물에는 다음을 포함한다.

- `ExcelSplitter.exe`
- 사용자 설명서
- 정식 Microsoft Excel 데스크톱 앱이 필요하다는 실행 조건
- 지원 파일 형식과 알려진 제한
- 버전 정보

릴리스 후보는 정식 Excel이 설치된 Windows PC에서 smoke test를 통과한 뒤 배포한다.

## 14. 기술 근거

- Microsoft의 [Excel `ListObject` 문서](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject)는 정식 Table과 `ListRows` 객체 모델을 정의한다.
- [`ListRow.Delete` 문서](https://learn.microsoft.com/en-us/office/vba/api/excel.listrow.delete)는 Table 행 셀을 삭제하고 아래 셀을 위로 이동시킨다고 명시하므로 Table 아래 콘텐츠를 사전 검사한다.
- [`Workbooks.Open` 문서](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)는 `UpdateLinks=0`이 외부 링크를 갱신하지 않는다고 정의한다.
- [openpyxl 문서](https://openpyxl.readthedocs.io/en/stable/tutorial.html)는 기존 파일을 열고 저장할 때 지원하지 않는 도형이 손실될 수 있다고 경고하므로 실제 결과 편집에 사용하지 않는다.
- [PyInstaller 문서](https://pyinstaller.org/en/stable/operating-mode.html)는 Python 인터프리터를 묶은 단일 실행 파일과 Python 미설치 실행을 설명한다.
- Python의 [`os.rename` 문서](https://docs.python.org/3/library/os.html#os.rename)는 Windows에서 대상 경로가 존재하면 `FileExistsError`를 발생시키므로 같은 volume의 no-clobber 게시에 사용한다.
- Microsoft의 [Excel 사양 및 제한](https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)은 경로를 포함한 파일명 길이를 218자로 제한한다.
- `pywin32`의 [`DispatchEx` 구현](https://github.com/mhammond/pywin32/blob/main/com/win32com/client/__init__.py)은 `CoCreateInstanceEx`를 호출하며, Microsoft는 [Excel을 다중 인스턴스 COM 서버로 설명](https://learn.microsoft.com/ko-kr/previous-versions/office/troubleshoot/office-developer/use-visual-c-automate-run-program-instance)한다.

## 15. 완료 조건

- 사용자가 `.xlsx`를 선택하고 워크시트, 단일 Table의 컬럼을 순서대로 선택할 수 있다.
- 선택 컬럼의 모든 고유값과 행 개수가 표시되며 빈 셀도 한 분류로 집계된다.
- 파일명 미리보기가 실제 출력 이름과 일치한다.
- 각 결과에는 선택 워크시트만 있고 Table에는 해당 분류 행만 있다.
- 수식 재계산 여부와 관계없이 최초 미리보기의 행별 분류 소속이 유지된다.
- 자체 포함 문서 요소와 허용되는 삭제 시트 의존 손상이 보존 계약과 일치한다.
- 원본 파일은 수정되지 않는다.
- 기존 파일은 승인 없이 덮어쓰지 않는다.
- 실패 시 원인과 부분 성공 여부가 명확히 표시된다.
- DRM 원본은 workflow당 한 번만 열리고 분류가 3개 이상이면 output worker 두 개를 사용한다.
- 결과 파일은 비보호 일반 `.xlsx`이며 비보호 master와 임시 파일이 작업 후 남지 않는다.
- 단위 테스트가 통과한다.
- 정식 Excel PC의 통합 및 smoke test가 통과한다.
- `ExcelSplitter.exe`가 Python 설치 없이 실행된다.
