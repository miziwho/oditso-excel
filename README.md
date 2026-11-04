# Excel 업무도구 v1.1

*[English README](README.en.md)*

Excel 2016+ 에서 자주 쓰는 작업 **39가지**를 단축키 하나로 실행하는 매크로 모음입니다.
파워쿼리 없이 열 피벗 해제, 다중 VLOOKUP, 두 목록 대조 같은 작업을 키 하나로 합니다.

<!-- 스크린샷 자리 — docs/images/ 에 이미지를 추가한 뒤 여기에 삽입 -->

## 설치 방법 (3단계)

1. Excel 창을 모두 닫습니다.
2. `설치파일.xlsm` 을 열고 **[콘텐츠 사용]** → **[설치하기]** 를 누릅니다.
3. Excel 을 완전히 종료한 뒤 다시 실행하고, `Ctrl+Shift+M` 으로 설치를 확인합니다.

> 파일은 [릴리스 페이지](../../releases/latest)에서 `설치파일.xlsm` 과 `PERSONAL.XLSB` 를 받거나,
> 이 저장소의 `dist/` 폴더에서 받을 수 있습니다.

**설치 프로그램이 보안 정책으로 막혔다면** 받은 `PERSONAL.XLSB` 를
`%APPDATA%\Microsoft\Excel\XLSTART` 폴더에 직접 넣어도 똑같이 동작합니다.
자세한 절차는 사용설명서의 「손으로 넣기」 항목에 있습니다.

## 처음이라면 이 다섯 개만

| 키 | 기능 |
|---|---|
| `Ctrl+Shift+M` | **업무도구 메뉴** — 나머지는 다 잊어도 됩니다 |
| `Ctrl+Shift+X` | 쉼표 형식 |
| `Ctrl+Shift+A` | 필터 켜기 / 끄기 |
| `Ctrl+Shift+W` | 행/열 자동맞춤 |
| `Ctrl+Shift+K` | 단축키 목록을 시트에 붙이기 |

39개를 외울 필요 없습니다. `Ctrl+Shift+M` 하나면 전부 메뉴에서 고를 수 있습니다.

## 단축키 39개

`Ctrl+Shift` 29개 + `Ctrl+Alt` 10개입니다.

### 편집

| 키 | 기능 | 되돌리기 |
|---|---|---|
| `Ctrl+Shift+O` | 새 창 열기 (같은 파일을 창 두 개로) | — |
| `Ctrl+Shift+Q` | 새 창으로 복사 (현재 시트를 새 통합문서로) | — |
| `Ctrl+Shift+X` | 쉼표 형식 (음수·0 모양은 Ctrl+Shift+M 에서 설정) | `Ctrl+Z` |
| `Ctrl+Shift+Y` | 텍스트 나누기 - 일반으로 (텍스트 숫자를 숫자로) | `Ctrl+Z` |
| `Ctrl+Shift+I` | 텍스트 나누기 - 텍스트로 | `Ctrl+Z` |
| `Ctrl+Shift+U` | 표 테두리 (머리글 굵게 + 세로 구분선) | `Ctrl+Z` |
| `Ctrl+Shift+W` | 행/열 자동맞춤 | `Ctrl+Z` |
| `Ctrl+Shift+S` | 첫 행 고정 / 해제 (머리글 자동 인식) | 다시 누르면 풀림 |
| `Ctrl+Shift+H` | 시트 일괄정리 (모든 시트 A1으로, 확대 100%) | — |
| `Ctrl+Alt+Z` | 가운데 정렬 / 기본 (토글) | `Ctrl+Z` |
| `Ctrl+Alt+S` | 선택 영역의 가운데로 (병합 없이 걸쳐 보이기) | `Ctrl+Z` |
| `Ctrl+Alt+X` | 서식 복사 / 붙여넣기 (한 키로 번갈아) | `Ctrl+Z` |
| `Ctrl+Alt+F` | 값으로 찾기 (수식 결과 기준, 찾은 칸 모두 선택) | — |

### 채우기 · 선택

| 키 | 기능 | 되돌리기 |
|---|---|---|
| `Ctrl+Shift+1` | 수식 아래로 채우기 (표 끝까지) | `Ctrl+Z` |
| `Ctrl+Shift+2` | 인덱스 열 삽입 (No, =ROW()-머리글행) | `Ctrl+Z` |
| `Ctrl+Shift+3` | 표 끝까지 선택 (빈 칸 무시) | — |
| `Ctrl+Shift+4` | 위 칸 값으로 빈 칸 채우기 (병합 푼 표 정리) | `Ctrl+Z` |
| `Ctrl+Alt+A` | 빠른 채우기 (Ctrl+E 업그레이드, 범위 지정 없이 표 끝까지) | `Ctrl+Z` |
| `Ctrl+Alt+Q` | 자동 채우기 (1,2,3 / 월,화,수 패턴을 표 끝까지) | `Ctrl+Z` |
| `Ctrl+Alt+E` | 머리글 위 행 제거 (내보내기 전 정리) | `Ctrl+Z` |

### 데이터

| 키 | 기능 | 되돌리기 |
|---|---|---|
| `Ctrl+Alt+B` | 열 머리글을 행 데이터로 (1월 2월 3월 열 -> 한 열로) | — |
| `Ctrl+Shift+L` | 다중 VLOOKUP (같은 값끼리 묶어 펼치기 · 그룹화) | `Ctrl+Z` |
| `Ctrl+Shift+V` | 기준 열 중복 (열 골라서 색칠 · 그룹번호 · 새 시트로 추출) | `Ctrl+Z` |
| `Ctrl+Alt+D` | 두 열 비교 (목록끼리 대조 · 같은 행끼리 비교) | `Ctrl+Z` |
| `Ctrl+Shift+D` | 중복행 색칠 (Ctrl로 골라 잡으면 값끼리 비교) | `Ctrl+Z` |
| `Ctrl+Shift+C` | 셀 배경 지우기 (선택 범위의 채우기 색 제거) | `Ctrl+Z` |
| `Ctrl+Shift+J` | 피벗테이블 만들기 (새 시트, 범위 자동) | — |
| `Ctrl+Shift+T` | 클래식 피벗 정리 (부분합 제거, 되돌리기 불가) | **되돌릴 수 없음** |
| `Ctrl+Shift+G` | SUBTOTAL 합계 (머리글 위 칸에서, 다시 누르면 취소) | `Ctrl+Z` |
| `Ctrl+Alt+G` | SUBTOTAL 개수 (채워진 칸 수, 필터 따라감) | `Ctrl+Z` |

### 필터

| 키 | 기능 | 되돌리기 |
|---|---|---|
| `Ctrl+Shift+A` | 필터 켜기 / 끄기 | 다시 누르면 풀림 |
| `Ctrl+Shift+Z` | 선택한 셀 값으로 필터 (다시 누르면 그 열만 해제) | 다시 누르면 풀림 |
| `Ctrl+Shift+F` | 고급필터 실행 (다시 누르면 해제) | 다시 누르면 풀림 |
| `Ctrl+Shift+R` | 고급필터 조건 열 세팅 (머리글 자동 복사) | `Ctrl+Z` |

### 내보내기 · 기타

| 키 | 기능 | 되돌리기 |
|---|---|---|
| `Ctrl+Shift+E` | 시트 빼서 텍스트 저장 (txt → dat 변환용) | — |
| `Ctrl+Shift+N` | 파일 경로 복사 (전체 경로를 클립보드로) | — |
| `Ctrl+Shift+P` | 암호 설정 (열기 암호, 기본값 1234) | — |
| `Ctrl+Shift+K` | 이 목록 붙이기 / 지우기 | `Ctrl+Z` |
| `Ctrl+Shift+M` | 업무도구 메뉴 - 기억 안 나면 이것만 누르세요 | — |

이 외에 **저장된 파일을 닫을 때 열기 암호가 없으면 자동으로 묻는** 기능이 있습니다.

## 되돌리기

| | 개수 | |
|---|---|---|
| `Ctrl+Z` 로 되돌아감 | 23개 | |
| 다시 누르면 풀림 | 4개 | 필터 계열 토글 |
| 되돌릴 것이 없음 | 11개 | 셀 선택 · 창 열기 · 경로 복사 등 |
| **되돌릴 수 없음** | **1개** | **`Ctrl+Shift+T` 클래식 피벗 정리** |

- 데이터가 바뀌는데 되돌아가지 않는 것은 **`Ctrl+Shift+T` 하나뿐**입니다. 실행 전 저장을 권합니다.
- `Ctrl+Z` 는 **바로 다음에** 눌러야 합니다. 중간에 다른 셀을 건드리면 Excel 이 등록을 지웁니다.
- 새 시트를 만드는 기능(`Ctrl+Alt+B`, `Ctrl+Shift+J`, `Ctrl+Shift+L`·`Ctrl+Shift+V` 의 새 시트 추출)은
  시트 탭을 우클릭해 삭제하면 됩니다.

## 주의사항

- 기존에 `PERSONAL.XLSB`(개인용 매크로 통합 문서)를 쓰고 있다면 설치 시 교체됩니다.
  **v1.1 설치 프로그램은 기존 파일을 자동으로 백업합니다** —
  `%APPDATA%\Microsoft\Excel\PERSONAL_백업_날짜.XLSB`
- Excel 기본 단축키 중 `Ctrl+Shift+1` `3` `4`(표시 형식), `Ctrl+Shift+A` `F` `P` 를 덮어씁니다.
- `Ctrl+Shift+B` 는 비워 두었습니다. 일부 PC 에서 다른 상주 프로그램이 이 키를 먼저 가져갑니다.
- 설정(쉼표 형식 모양 · 암호 확인 · 기본 암호)은 `PERSONAL.XLSB` 안에 저장됩니다.
  레지스트리를 쓰지 않으므로 그 파일만 옮기면 설정이 따라갑니다.

## 제거 방법

`설치파일.xlsm` 에서 `Alt+F8` → `PERSONAL_제거` 를 실행합니다.
직접 넣은 경우에는 `XLSTART` 폴더의 `PERSONAL.XLSB` 를 지우면 됩니다.

## 동작 환경

- Excel 2016 이상, Windows, 32/64bit 공통
- 외부 참조 · DLL 없음

## 사용설명서

- [`docs/Excel_업무도구_v1.1_사용설명서.html`](docs/Excel_업무도구_v1.1_사용설명서.html)
- [`docs/Excel_Work_Tools_v1.1_Guide.html`](docs/Excel_Work_Tools_v1.1_Guide.html) — 영문판

설치 · 39개 기능 설명 · 캡처 · 자주 묻는 질문이 한 장에 들어 있습니다.
영문판은 영어로 먼저 열리고, 사이드바의 **한국어** 단추로 한글판으로 돌아갑니다.

## 저장소 구조

```
oditso-excel/
├─ src/        VBA 소스 (CP949 인코딩 — UTF-8로 변환 금지)
│   ├─ mod업무도구.bas          PERSONAL.XLSB 본체
│   ├─ ThisWorkbook_코드.txt    ThisWorkbook 에 붙여넣는 코드
│   ├─ modInstall.bas           설치파일.xlsm 안의 설치 로직
│   └─ en/                      영문판
│       ├─ mod_WorkTools.bas        본체
│       ├─ ThisWorkbook_code.txt    ThisWorkbook 에 붙여넣는 코드
│       ├─ modInstall.bas           설치 로직
│       └─ INSTALL.txt              영문 설치 안내
├─ docs/       HTML 사용설명서 (한글판, 영문판)
└─ dist/       배포본 (설치파일.xlsm, PERSONAL.XLSB)
```

## 라이선스

[CC BY-NC-ND 4.0](LICENSE) — 저작자 표기: **oditso (oditso.com)**

## 링크

- https://oditso.com
- 도움이 되셨다면 ☕ [커피 한 잔](https://www.oditso.com/donate)
