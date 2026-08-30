# Excel 업무도구 v1.0

Excel 2016+ 에서 자주 쓰는 작업 23가지를 `Ctrl+Shift+` 단축키 하나로 실행하는 개인용 매크로 모음입니다.

<!-- 스크린샷 자리 — docs/images/ 에 이미지를 추가한 뒤 여기에 삽입 -->

## 설치 방법 (3단계)

1. Excel 창을 모두 닫습니다.
2. `설치파일.xlsm` 을 열고 **[콘텐츠 사용]** → **[설치하기]** 를 누릅니다.
3. Excel 을 완전히 종료한 뒤 다시 실행하고, `Ctrl+Shift+M` 으로 설치를 확인합니다.

> 파일은 [릴리스 페이지](../../releases/latest)에서 `설치파일.xlsm` 과 `PERSONAL.XLSB` 를 받거나,
> 이 저장소의 `dist/` 폴더에서 받을 수 있습니다.

## 단축키 23개

### 편집

| 키 | 기능 | `Ctrl+Z` |
|---|---|---|
| `Ctrl+Shift+O` | 새 창 열기 | — |
| `Ctrl+Shift+Q` | 새 창으로 복사 (현재 시트를 새 통합문서로) | — |
| `Ctrl+Shift+X` | 쉼표 형식 (음수 빨강 △, 0은 `-`) | ✓ |
| `Ctrl+Shift+Y` | 텍스트 나누기 - 일반으로 | ✓ |
| `Ctrl+Shift+I` | 텍스트 나누기 - 텍스트로 | ✓ |
| `Ctrl+Shift+U` | 표 테두리 | ✓ |
| `Ctrl+Shift+W` | 행/열 자동맞춤 | ✓ |
| `Ctrl+Shift+S` | 첫 행 고정 / 해제 | — |
| `Ctrl+Shift+H` | 시트 일괄정리 (모든 시트 A1으로) | — |

### 데이터

| 키 | 기능 | `Ctrl+Z` |
|---|---|---|
| `Ctrl+Shift+J` | 피벗테이블 만들기 (새 시트) | — |
| `Ctrl+Shift+T` | 클래식 피벗 정리 | **✗** |
| `Ctrl+Shift+D` | 중복행 색칠 (보이는 행만) | ✓ |
| `Ctrl+Shift+C` | 셀 배경 지우기 (보이는 셀만) | ✓ |
| `Ctrl+Shift+G` | SUBTOTAL 합계 (다시 누르면 취소) | ✓ |

### 필터

| 키 | 기능 | `Ctrl+Z` |
|---|---|---|
| `Ctrl+Shift+A` | 필터 켜기 / 끄기 | — |
| `Ctrl+Shift+Z` | 선택한 셀 값으로 필터 (다시 누르면 해제) | — |
| `Ctrl+Shift+F` | 고급필터 실행 (다시 누르면 해제) | — |
| `Ctrl+Shift+R` | 고급필터 조건 열 세팅 | ✓ |

### 내보내기 · 기타

| 키 | 기능 | `Ctrl+Z` |
|---|---|---|
| `Ctrl+Shift+E` | 시트 빼서 텍스트 저장 (txt → dat 변환용) | — |
| `Ctrl+Shift+N` | 파일 경로 복사 | — |
| `Ctrl+Shift+P` | 암호 설정 (기본값 1234) | — |
| `Ctrl+Shift+K` | 단축키 목록 붙이기 / 지우기 | ✓ |
| `Ctrl+Shift+M` | **업무도구 메뉴 — 기억 안 나면 이것만** | — |

이 외에 **저장된 파일을 닫을 때 열기 암호가 없으면 자동으로 묻는** 기능이 있습니다.

## 되돌리기 안내

- 위 표에서 `✓` 표시된 **10개** 기능은 `Ctrl+Z` 로 되돌릴 수 있습니다.
- **클래식 피벗 정리(`Ctrl+Shift+T`)만 되돌릴 수 없습니다.** 실행 전 파일을 저장해 두세요.

## 주의사항

- 기존에 `PERSONAL.XLSB`(개인용 매크로 통합 문서)를 쓰고 있다면 **백업 없이 교체됩니다.**
  설치 전 반드시 백업하세요.
- `Ctrl+Shift+` 조합 23개를 점유합니다.
- 저장된 파일을 닫을 때 열기 암호가 없으면 자동으로 묻습니다.
  `Ctrl+Shift+M` → `닫을 때 암호 확인` 에서 끌 수 있습니다.

## 제거 방법

`설치파일.xlsm` 에서 `Alt+F8` → `PERSONAL_제거` 를 실행합니다.

## 동작 환경

- Excel 2016 이상, Windows, 32/64bit 공통
- 외부 참조 · DLL 없음

## 사용설명서

- [`docs/Excel_업무도구_v1.0_사용설명서.html`](docs/Excel_업무도구_v1.0_사용설명서.html)

## 저장소 구조

```
oditso-excel/
├─ src/    VBA 소스 (CP949 인코딩 — UTF-8로 변환 금지)
│   ├─ mod업무도구.bas          PERSONAL.XLSB 본체
│   ├─ ThisWorkbook_코드.txt    ThisWorkbook 에 붙여넣는 코드
│   └─ modInstall.bas           설치파일.xlsm 안의 설치 로직
├─ docs/   HTML 사용설명서
└─ dist/   배포본 (설치파일.xlsm, PERSONAL.XLSB)
```

## 라이선스

[CC BY-NC-ND 4.0](LICENSE) — 저작자 표기: **oditso (oditso.com)**

## 링크

- https://oditso.com
