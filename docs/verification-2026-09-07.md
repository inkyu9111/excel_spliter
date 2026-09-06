# Excel File Toolkit 실사용 검증 기록

검증일: 2026-09-07, 한국시간. 초기 배포 빌드는 오전 3시 12분에 확인했으며, 아래 아이콘·성능 후속 작업의 빌드를 오전 8시 34분에 다시 확인했다. 후속 실제 Excel 검사는 오전 8시 58분까지 완료했다. 회사 자료 없이 임시 합성 `.xlsx`만 사용했다.

## 판정

기존 Split과 Merge가 모두 구현되어 있음을 코드·기존 테스트·README에서 확인했다. 초기 작업은 진행 막대와 화면 크기·경로 표시를 수정했다. 후속 작업에서는 아이콘을 적용하고 병합의 조건부 서식 검사 한 곳만 최적화했다. 분할 처리·병렬 작업자·공유 Excel 처리 코드는 변경하지 않았다.

필터를 해제한 정식 Excel Table의 분할 및 같은 열 구조의 단일 시트 Table 병합은 실제 Excel 통합 검사에서 통과했다. 15개 파일 조건부 서식 검사와 최종 배포 빌드도 통과했다. 아래 지원 범위와 미검증 항목을 확인하고 업무 파일의 복사본으로 첫 실행을 해야 한다.

## 저장소와 환경

- 저장소: `inkyu9111/excel_spliter`, 기준 커밋 `1bbedbf76ce9db85d377d2550ce841496d2c862a`. 저장소 이름을 유지했다.
- 작업 위치: `C:\dev\inkyu9111\excel_spliter`. 환경에서 기존 로컬 작업본을 찾지 못해 별도 복제본에서 작업했다. 다른 작업본의 사용자 변경은 가져오거나 덮어쓰지 않았다.
- Windows 11, Python 3.12.14 64비트, 데스크톱 Excel 16.0 build 20326.0.
- 저장소의 고정 의존성 그대로 사용: pywin32 312, pytest 8.4.2, PyInstaller 6.22.2. 새 런타임 의존성·플러그인·서버를 추가하지 않았다.
- `my-skills`의 `risk-routed-development`, `skill-authoring`을 커밋 `ef708f853565a13903d566e23b513e892b0f52c7` 기준으로 갱신했다. 기존 내용도 최신본과 같았다. 기존 설치본은 `C:\dev\inkyu9111\toolkit-verification\skills-backup-20260907`에 보관했다.

## 최소 수정

| 파일 | 변경과 이유 |
|---|---|
| `src/excel_splitter/gui.py` | 전체량을 모르는 진행 막대의 최대값 1과 15ms 타이머로 한 프레임이 전 범위를 이동하던 원인을 수정. 최대값 100, 50ms로 변경하고 반복 상태 통지가 애니메이션을 0으로 되돌리지 않게 했다. |
| `src/excel_splitter/toolkit_gui.py` | 네 탭의 진행 값과 상태 문구를 분리. 기본 창 900×680, 최소 740×520을 유지하며 여백을 줄임. 긴 경로는 입력칸 아래에 전체 문자열을 줄바꿈하고 병합 목록에는 가로 스크롤과 선택 파일의 전체 경로를 표시. |
| `tests/unit/test_toolkit_gui.py` | 실제 Tk 타이머·탭 간 상태 분리·실제 백그라운드 작업의 중간 오류·재실행·작은 창과 긴 한글 경로의 회귀 검사 추가. |
| `tests/unit/test_compare_service.py`, `test_etc_service.py` | Windows 심볼릭 링크 생성 권한 오류 1314에 한해서 명시적으로 skip. 다른 오류는 그대로 실패 처리. |
| `scripts/check_split_merge_excel.py` | 실제 Excel에서 독립적인 고정 기대값으로 분할→병합과 원본·기존 출력 보호를 확인하는 실행 가능한 합성 검사 추가. |
| `scripts/check_merge_excel.py` | 기존 검사에서 Excel 종료 전에 COM 객체 참조를 정리하고, 합성 조건부 서식 생성에 위치 인수를 사용. 기존 데이터·서식 기대값은 유지. |
| `README.md`, `scripts/smoke_test.ps1`, `.gitignore` | 사용법·필터 제한·검사 명령 기록. 수동 검사 안내에서 도형·차트 보존에 관한 잘못된 설명 수정. editable 설치 산출물 `*.egg-info/` 제외. |

## 실행 명령과 결과

저장소 루트에서 실행한다. 이 환경의 셸 명령에는 `RTK.md`에 따라 `rtk proxy`를 붙였다. 데스크톱 Excel과 Tk 검사는 로그인한 Windows 사용자 세션에서 실행했다.

```powershell
rtk proxy .\.venv\Scripts\python.exe -m pytest tests/unit -q -rs -p no:cacheprovider --basetemp C:\dev\inkyu9111\toolkit-verification\pytest-final
rtk proxy .\.venv\Scripts\python.exe -u scripts/check_split_merge_excel.py
rtk proxy .\.venv\Scripts\python.exe -u scripts/check_merge_excel.py
rtk proxy powershell -NoProfile -ExecutionPolicy Bypass -File scripts\build.ps1 -PythonExe C:\dev\inkyu9111\excel_spliter\.venv\Scripts\python.exe
```

`--basetemp`에는 재검사마다 새로운 빈 경로를 사용한다. 실제 Excel 검사들은 순서대로 실행한다. 검사 중에는 다른 프로그램에서 복사·붙여넣기를 하지 않는다.

| 검사 | 결과 |
|---|---|
| 수정 전 단위 테스트 | 심볼릭 링크 4건 제외 후 373 passed. |
| 진행 막대·오류 상태·레이아웃 관련 최종 대상 검사 | 59 passed. 수정 전에는 실제 타이머 이동폭과 다른 탭의 진행·상태 오염을 재현했고 수정 후 통과. |
| `check_split_merge_excel.py` | **통과, 종료 코드 0, 555.3초.** 아래 분할·병합·잠금 검사와 필터 제한 확인 포함. |
| 전체 단위 테스트 | **387 passed, 4 skipped, 11.95초, 종료 코드 0.** skip은 현재 Windows 계정의 심볼릭 링크 생성 권한 부족 4건. |
| 기존 `check_merge_excel.py` | **통과, 종료 코드 0, 142.1초.** 값·중복·계산값·셀별 표시 형식·합계·필터/숨김 행·표 밖 콘텐츠·원본 보존·15개 파일의 단일 `$H:$H` 조건부 서식 확인. |
| 기존 배포 빌드 | **통과, 종료 코드 0, 41.7초.** 고정 의존성·전체 테스트 재실행(387 passed, 4 skipped, 11.89초)·onedir 시작·onefile 생성·내장 pythoncom DLL·최종 EXE pywin32 self-test 확인. |

### 심볼릭 링크 4건 재검사 (오전 7시 45분)

사용자 요청으로 권한 부족 검사만 다시 실행했다. 비교와 시트 정리에서 각각 ① 출력이 원본을 가리키는 실제 심볼릭 링크인 경우, ② 출력이 없는 파일을 가리키는 끊어진 심볼릭 링크인 경우를 검사한다. 두 경우 모두 Excel 처리 전에 거부하고 원본을 보존해야 한다.

```powershell
rtk proxy .\.venv\Scripts\python.exe -m pytest tests/unit/test_compare_service.py tests/unit/test_etc_service.py -k symlink -vv -rs -p no:cacheprovider --basetemp C:\dev\inkyu9111\toolkit-verification\pytest-symlink-recheck-0739
```

결과는 **4 skipped, 69 deselected, 0.19초**였다. 일반 사용자 토큰에 `SeCreateSymbolicLinkPrivilege`가 없으며 링크 생성에서 Windows 오류 1314가 발생한다. 관리자 프로세스로 재검사하도록 UAC를 요청했으나 Windows가 `The operation was canceled by the user`를 반환해 관리자 검사는 실행되지 않았다. 이 네 건은 여전히 미검증이다.

다음 전용 명령은 네 검사만 실행하고 로그·JUnit XML·`result.json`을 검사 폴더에 남긴다. **4 passed와 0 skipped**가 모두 확인돼야 성공으로 처리한다. 실행에는 사용자의 Windows UAC 승인이 필요하다.

```powershell
rtk proxy powershell -NoProfile -ExecutionPolicy Bypass -File C:\dev\inkyu9111\toolkit-verification\symlink-recheck-20260907\run-admin.ps1
```

**후속 작업 — 사용자 요청으로 추후 진행**

- [ ] Windows 관리자 승인이 가능한 환경에서 위 전용 명령으로 심볼릭 링크 검사 4건을 재실행한다.
- [ ] **4 passed, 0 skipped**와 원본 보존을 확인한 뒤 로그·JUnit XML 경로와 결과를 이 문서에 추가한다. 그때까지 네 건은 미검증으로 유지한다.

로그는 `C:\dev\inkyu9111\toolkit-verification`에 저장한다. 초기 실행에서는 샌드박스의 임시 폴더·Tcl 접근 제한과 COM 로그인 세션 오류가 있었다. 정상 사용자 세션의 검사와 구분하며, 초기 실패를 통과 건수에 포함하지 않는다.

실제 Excel의 초기 반복 세션 검사에서는 `Workbooks` 접근 시 `0x800706B5` COM 인터페이스 오류와 검증 재열기 중 Python 접근 위반(`0xC0000005`)이 발생했다. 추적에서 Excel 세션 종료 이후 남아 있는 COM 참조가 원인 후보로 좁혀졌지만 모든 초기 오류의 원인을 확정한 것은 아니다. 별도 검사 복사본에서 참조를 정리하자 첫 병합 결과 검증을 통과했고, 그 다음 합성 조건부 서식 생성에서 `DISP_E_PARAMNOTOPTIONAL`이 드러났다. 검사 코드의 참조 정리와 인수 전달을 보정했다. 제품의 공유 Excel 처리 코드는 변경하지 않았다.

다른 PC에서의 COM 연결 안정성까지 보장하지 않는다. 오류가 발생한 실행은 실패로 취급하고 오류 상세·로그를 확인해야 한다. 초기 실패 로그는 `check_merge_excel-first-failure.log`, 충돌 추적은 `trace-merge.log`, 대조 실험은 `check-merge-lifetime.log`에 보관했다.

## 실제 데이터·파일 보존 확인

새 통합 검사는 Excel로 만든 한글·공백 경로의 6행 자료를 사용한다. 세 그룹의 크기는 3·2·1이며 완전히 동일한 중복 행, 빈 분류값, 숫자, `=literal` 문자열, 표 수식, 셀 서식, 수동 숨김 행과 다른 시트 참조를 포함한다.

| 요구사항 | 확인 내용 |
|---|---|
| 행 누락·중복·값 변형 | 각 분할 결과의 모든 셀을 고정 기대값과 비교. 병합 6행의 정확한 순서·값 및 원본과 행별 중복 횟수 비교. 기존 중복 행은 그대로 유지. |
| 한글·공백 경로 | 실제 한글 폴더와 파일명으로 생성·분할·병합·재열기 통과. |
| 빈 시트·빈 Table | 표 없는 시트 거부, 0행 Table의 Split 거부, 0행 Table을 첫 입력으로 한 Merge 성공. |
| 머리글 불일치·다중 시트 | 서로 다른 머리글, 여러 시트가 있는 Merge 입력 거부. |
| 같은 파일 중복 선택 | Merge 미리보기에서 거부. |
| 결과 파일 재입력 | 이전 병합 결과를 새 출력 경로로 재입력하면 8행으로 합쳐짐. 같은 결과 경로로 원본을 덮어쓰려 하면 거부. |
| 잠금·덮어쓰기 | Windows 파일 핸들로 삭제·교체를 막은 출력 사용. Merge 기존 결과 보존, Split은 잠긴 그룹만 실패하고 다른 결과 유지. 승인 없는 덮어쓰기는 거부. |
| 미리보기 이후 변경 | 복제한 합성 입력을 변경하자 실행 거부. 출력 미생성. |
| 원본·임시 파일 | 성공과 예상 오류 이후 입력 SHA-256·크기·수정시간 동일. 기존 잠긴 출력 동일. 임시 파일 정리 확인. |
| 진행 막대 | 네 탭 × 전체량 유무의 8가지 중간 오류에서 타이머 정지·0 초기화·늦게 도착한 이벤트 무시·탭 복구·다음 작업 완료 확인. |

## 확인된 지원 범위와 제한

- **Split**: 선택한 시트의 정식 Excel Table 하나를 분류한다. 다른 시트는 결과에서 삭제된다. Table 내부 수식과 시험한 셀 서식·열 너비는 유지됐다. 삭제된 다른 시트를 참조하는 수식이 `#REF!`가 되는 것도 실제 확인했다. 참조 복구는 지원하지 않는다.
- **활성 필터**: 필터가 적용된 표의 행 삭제를 Excel이 거부하는 사례를 재현했다. 해당 검사는 실패를 예상하고 원본 보존·임시 파일 정리를 확인한 것이다. 필터 적용 상태의 분할을 지원하거나 성공했다고 판정한 것이 아니다. 분할 전에 필터를 해제해야 한다. 수동 숨김 행은 정상 분할에 포함됐다.
- **Merge**: 입력마다 시트 하나·Table 하나, 머리글 이름과 순서가 같아야 한다. 목록 순서로 이어 붙이며 분할 이전의 교차 행 순서를 복원하거나 중복 행을 제거하지 않는다. Table 수식은 계산값으로 고정된다. 첫 파일의 레이아웃·Table 스타일을 사용하고 각 데이터 셀의 서식을 복사한다.
- Table 밖의 다른 파일 콘텐츠, 외부 참조 복구, 임의의 여러 시트 통합은 지원하지 않는다. Split에서 선택 시트의 도형·차트·메모·댓글은 삭제된다. 보호 시트·외부 연결 Table 등 상세 제한은 README를 따른다.
- 일부 Split 결과가 실패해도 성공한 결과는 남는다. 전체 성공으로 오해하지 말고 결과별 성공·실패 목록을 확인해야 한다.
- 강제 종료·전원 차단의 원자성이나 임시 파일 정리를 보장하지 않는다. 출력 폴더에 평문 임시 파일이 남을 수 있다.

## 이번 환경에서 확인하지 못한 항목

- 회사 DRM, 실제 민감 자료, 보안 프로그램이 적용된 업무 PC에서의 실행.
- 모든 Excel 버전, 네트워크 공유·OneDrive 동기화·장시간 대용량 작업·시트 행 한도 근처의 실제 성능.
- 화면 배율별 사람이 보는 최종 외관 및 수동 전체 업무 흐름. 실제 Tk의 크기·배치·줄바꿈·타이머를 검사했지만 화면 캡처 도구는 승인 대기 시간 초과로 완료하지 못했다.
- 심볼릭 링크 관련 단위 테스트 4개. 현재 Windows 계정의 생성 권한 부족으로 skip이며 통과가 아니다. 하드 링크와 일반 경로 검사는 별도로 실행한다.
- Compare·Etc의 별도 실제 Excel 검사 스크립트는 이번 분할·병합 검증 범위에서 실행하지 않았다. 관련 단위 테스트와 공통 UI의 네 탭 오류 복구 검사는 전체 테스트에 포함됐다.
- 같은 스레드에서 COM을 반복 초기화하며 서비스를 연속 호출하는 별도 사용 방식의 안정성. 후속 검사에서 재연결 오류를 재현했으며, 공유 Excel 코드를 수정하거나 해결됐다고 보고하지 않았다.

## 아이콘·성능 후속 작업

사용자가 선택한 초록색 문서·렌치 PNG를 `src/excel_splitter/assets/app-icon.png`에 원본 바이트 그대로 보관했다. 16·24·32·48·64·128·256px의 DIB 형식 `app.ico`를 만들어 EXE 리소스와 Tk 창에 함께 적용했다. PNG 압축 ICO는 Tk에서 다른 픽셀로 읽히는 것을 재현해 사용하지 않았다. 런타임 의존성은 추가하지 않았다.

실제 Tk 창, 빌드한 폴더형 EXE와 단일 EXE의 작은·큰 창 아이콘을 Windows API로 읽어 같은 크기의 ICO 픽셀과 대조했다. EXE에 내장된 일곱 크기의 아이콘 데이터도 ICO와 정확히 일치했다. 이 자동 검사는 사람이 화면 배율별 최종 외관을 확인하는 검사를 대신하지 않는다.

조건부 서식이 없는 일반 UTF-8 워크시트를 전체 DOM으로 만들고 workbook·styles·table·theme를 읽던 비용을 줄였다. 조건부 서식 또는 DTD가 있거나 인코딩 선언이 다르면 기존 분석을 유지한다. UTF-16·ISO-8859-1·소문자 UTF-8 선언·엔터티로 생성한 서식도 회귀 검사에 포함했다. 원본 서명 검사, 파일 복사, 결과 검증·게시와 Split의 병렬 처리 방식은 그대로다.

| 후속 검사 | 확인 결과 |
|---|---|
| 실제 Tk 아이콘 및 기존 UI·진입점 | 46 passed. 아이콘 설정이 없을 때 실패하고 DIB ICO 적용 후 통과. |
| 최종 전체 단위 테스트 | **397 passed, 4 skipped, 28.20초.** skip 네 건은 앞서 기록한 심볼릭 링크 생성 권한 부족이며 추후 검사 대상이다. |
| 최종 배포 빌드 | **종료 코드 0, 89.7초.** 의존성·전체 테스트·onedir 시작·onefile·pywin32·모든 EXE 아이콘 리소스·실제 실행 창 아이콘·정상 종료 확인. |
| 최종 `check_split_merge_excel.py` | **종료 코드 0, 342.3초.** 위 분할→병합의 고정 기대값·중복·수식·서식·빈 표·검증 오류·덮어쓰기·잠금·재입력·미리보기 이후 변경·원본 보존·임시 정리를 재확인. 활성 필터 거부는 지원 제한 검사로 구분. |
| 기존 동일 스레드 `check_merge_excel.py` 재실행 | **실패, 종료 코드 1, 12.8초.** 미리보기·실행 이후 결과를 다시 열려고 새 Excel 세션의 `Workbooks`에 접근할 때 `0x800706B5` 재발. 초기 검증의 통과 기록과 구분하며 이번 실행은 통과로 세지 않는다. |
| COM 수명 대조 실험 | 검사 호출 스레드의 COM 초기화를 전체 실행 동안 유지하면 기존 기대값과 15파일 규칙 검사 **통과, 종료 코드 0, 85.0초**. COM 수명이 이 재현에 영향을 준다는 근거이며 공유 Excel 코드가 수정됐거나 원인이 완전히 해결됐다는 뜻은 아니다. |
| GUI 작업 스레드 대조 실험 | 바깥 COM 초기화 없이, 기존 미리보기·실행 메서드를 작업마다 새 스레드에서 호출한 검사 **통과, 종료 코드 0, 143.2초**. 기존 모든 데이터·CF·원본 보존 기대값 유지. |
| 보정 후 최종 `check_merge_excel.py` | **통과, 종료 코드 0, 140.9초.** GUI와 동일한 작업별 새 스레드에서 값·순서·중복·수식 계산값·셀 서식·합계·필터/숨김 행·표 밖 콘텐츠·원본 보존과 15파일의 단일 `$H:$H` 규칙을 재확인. |
| 조건부 서식 없는 20,000×8셀 합성 ZIP | 5회 교차 측정 중앙값 **3.2518초 → 0.0147초**, ZIP 내부 읽기 **5 → 1회**. 이전·현재 코드의 고정 기대 판정 및 원본 바이트 보존 통과. |
| 같은 크기의 조건부 서식 있는 ZIP | 중앙값 **3.0796초 → 3.7163초**, ZIP 읽기 5회로 동일. 이 표본은 더 느렸으며 이 경로의 속도 향상을 주장하지 않는다. 규칙 판정과 원본 보존은 동일했다. |
| 기존 코드·원본 보존 | Git 비교에서 SplitService·controller·Excel gateway·source session·parallel writer와 기존 `dist/ExcelSplitter.exe` 변경 없음. |

실행 명령은 저장소 루트 기준이다.

```powershell
python scripts/check_executable_icon.py dist/ExcelFileToolkit.exe
python scripts/benchmark_merge_cf.py
python scripts/benchmark_merge_cf.py --baseline C:\dev\inkyu9111\toolkit-verification\icon-performance-20260907-0809\start\src\excel_splitter\merge_conditional_formats.py
```

성능 스크립트의 baseline 인수에는 비교할 신뢰하는 이전 모듈 파일을 지정한다. 합성 ZIP만 생성하며 Excel·파일 복사·전체 병합 시간을 측정하는 스크립트가 아니다.

벤치마크 로그는 `benchmark-merge-cf-final.log`이다. 시간은 PC 부하에 따라 변하며 위 수치는 조건부 서식 지원 판정 단계의 측정값이다. 전체 병합·분할의 처리 속도나 조건부 서식이 있는 파일의 성능이 개선됐다고 해석하지 않는다. 비교 탭의 키 표시 문자열 생성은 10만 행에서 약 69ms 절감 후보였으나 이번 최소 수정에 포함하지 않았다.

`check_merge_excel.py`는 GUI의 실제 계약처럼 미리보기와 실행마다 새 작업 스레드를 사용하는 검사로 맞췄다. `ThreadPoolExecutor`를 호출마다 새로 만들고 `.result()`로 실패를 전달한다. 바깥 COM 초기화는 추가하지 않았고 기존 데이터·서식·원본 보존 assert를 모두 유지했다. 별도의 범위 검토에서 GUI 호출 방식과의 일치 및 예외 전달을 확인했다. 이 보정은 같은 스레드의 반복 COM 사용 문제를 해결한 제품 변경이 아니다. 원래 실패는 `check-merge-same-thread-failure.log`, 두 대조 실험은 `run-merge-apartment-probe.log`와 `run-merge-gui-workers-probe.log`에 보존했다.

후속 로그는 `C:\dev\inkyu9111\toolkit-verification\icon-performance-20260907-0809`에 보관한다. 첫 빌드의 아이콘 검사에서는 검사 코드가 지원되지 않는 PyHANDLE context manager를 사용해 실패했다. 명시적인 `try/finally` 핸들 해제로 보정한 뒤 두 EXE와 전체 빌드를 다시 실행해 통과했다. 초기 실패 로그도 `build-first-icon-probe-failure.log`에 보존했다.

검증한 새 `dist/ExcelFileToolkit.exe`는 14,698,745바이트이며 SHA-256은 `84c8937c004b8935d3bc633f4c3c294542acf8fae95de645c0cbd1276208b35b`이다. 이전 `ExcelSplitter.exe`는 보존하며 새 프로그램은 `ExcelFileToolkit.exe`를 실행한다.

## 사용법과 실사용 직전 확인

1. `dist\ExcelFileToolkit.exe`를 실행한다. 소스에서는 `.\.venv\Scripts\python.exe -m excel_splitter`로 실행한다.
2. 원본 복사본과 새 출력 폴더를 준비한다. Excel에서 열리는 `.xlsx`인지 확인하고 원본·출력 파일을 닫는다.
3. 분할은 정식 Table 하나가 있는 시트와 분류 열을 선택한다. 활성 필터를 해제하고 `%_분할` 같은 패턴으로 미리보기의 그룹별 행 수와 출력 이름을 확인한다.
4. 병합은 동일한 머리글·순서의 단일 시트 파일을 두 개 이상 추가한다. 원하는 목록 순서를 정하고 입력과 다른 결과 경로로 미리보기를 실행한다.
5. 실행 중 다른 프로그램의 복사·붙여넣기를 피한다. 오류가 나면 성공·실패 목록과 기존 파일을 먼저 확인하고 새 출력 위치에서 다시 시도한다.
6. 첫 업무 실행에서는 입력 총행 수와 출력 총행 수, 중요 금액 합계·문자값·수식·필요 서식을 Excel에서 대조한다. Split→Merge는 행 순서가 바뀔 수 있으므로 단순 좌표 비교만으로 누락을 판정하지 않는다.
7. 다른 시트 참조, DRM 유지, 모든 도형·수식 보존이 필수인 문서라면 위 제한을 충족하지 않으므로 적용하지 않는다.

수정된 소스와 실행 가능한 검사는 저장소에 있으며, 기존의 추적 중인 `dist/ExcelSplitter.exe`는 덮어쓰지 않았다.
