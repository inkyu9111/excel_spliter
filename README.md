# Excel Splitter

Excel Splitter는 한 Excel Table의 분류 컬럼을 기준으로 `.xlsx` 통합문서를 여러 파일로 나누는 Windows 데스크톱 도구입니다. 결과 파일에는 선택한 워크시트만 남습니다.

## 입력과 사용 조건

- 원본은 `.xlsx` 파일이어야 합니다.
- 선택 워크시트에는 정식 Excel Table이 정확히 하나 있어야 합니다.
- 데스크톱 Microsoft Excel 2016 이상이 설치된 64비트 Windows 10/11이 필요합니다.
- 출력 폴더는 존재하고 쓰기 가능해야 하며, 원본과 같은 경로를 결과로 지정할 수 없습니다.
- 외부 데이터에 연결된 Table과 Table 아래의 독립 콘텐츠가 있는 입력은 지원하지 않습니다.

파일명 패턴에는 분류 값 자리표시자 `%`가 하나 이상 있어야 합니다. 기본 패턴은 `%_분할`입니다. 모든 `%`가 분류 값으로 바뀌며 빈 분류 값은 빈 문자열이므로 `%_분할`은 `_분할.xlsx`가 됩니다. Windows에서 사용할 수 없는 문자는 `_`로 정리되고, 같은 이름은 ` (2)`, ` (3)` 접미사로 구분됩니다.

미리보기에서 분류, 행 수, 출력 파일명을 확인한 뒤 분할합니다. 기존 파일 충돌은 전체 목록을 한 번에 승인해야 덮어씁니다. 개별 그룹 실패는 다음 그룹 처리를 막지 않으며 완료 창에 성공 파일과 각 실패 원인이 함께 표시됩니다. 작업 전체 예외가 발생하면 성공으로 표시하지 않습니다.

> 선택한 시트를 제외한 다른 시트를 삭제합니다. 삭제되는 시트에 의존하는 수식, 이름, 차트, 유효성 검사 및 연결이 손상될 수 있습니다.

보존 범위와 허용되는 변경의 전체 계약은 [설계 문서](docs/superpowers/specs/2026-09-01-excel-splitter-design.md)를 참고하세요.

## 개발과 빌드

고정된 개발 의존성을 설치한 Python 3.12 환경에서 단위 테스트를 실행합니다.

```powershell
& C:\path\to\python.exe -m pytest tests/unit -q
```

단일 실행 파일은 다음 스크립트로 빌드합니다. 스크립트는 고정 의존성 버전과 단위 테스트를 확인하고 one-folder 시작 검사를 거쳐 `dist\ExcelSplitter.exe`를 만듭니다.

```powershell
.\scripts\build.ps1 -PythonExe C:\path\to\python.exe
```

정식 Excel PC에서는 고정 fixture와 수동 보존 체크리스트를 생성하는 smoke test를 실행합니다.

```powershell
.\scripts\smoke_test.ps1 -ExePath .\dist\ExcelSplitter.exe -WorkbookPath C:\temp\excel-splitter-fixture.xlsx
```
