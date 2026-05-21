# AutoDWG 도면 검토 자동화 시스템 — CLAUDE.md

## 프로젝트 개요

실시설계 도서 납품 전, **도면목록표(LIST DWG)** 와 **개별 캐드 도면(DWG 수백 장)** 을
자동으로 교차 검토하여 컬러 엑셀 리포트를 생성하는 Windows 단독 실행(EXE) 앱.

- **현재 버전:** v6.8
- **진입점:** `app.py` (단일 파일, 88 KB)
- **최종 산출물:** `dist/DWGChecker.exe` (PyInstaller 단일 EXE)
- **설정 파일 저장 경로:** `%APPDATA%\AutoDWG_Checker\<블록명>.json` (ROI 좌표)
- **로그 파일:** `%APPDATA%\AutoDWG_Checker\autodwg.log`

---

## 파일 구조

```
dwg-checker-main/
├── app.py              # 메인 소스 (GUI + 전체 로직)
├── SET_ROI.lsp         # AutoCAD LISP — ROI 좌표 설정 및 JSON 저장
├── build.bat           # EXE 빌드 스크립트 (venv\Scripts\pyinstaller 사용)
├── build.spec          # PyInstaller 스펙 (customtkinter/tkinterdnd2 asset 포함)
├── requirements.txt    # Python 의존성
├── readme.html         # 사용 설명서 (이미지 base64 내장, 단독 배포 가능)
├── readme.md           # 사용 설명서 Markdown 버전
├── KNOWN_ISSUES.md     # 수정 보류 중인 제약사항 8개 문서화
├── image/              # readme.html/md에 삽입된 스크린샷 원본
└── venv/               # Python 가상환경 (빌드 전 의존성 설치 필요)
```

---

## 기술 스택

| 역할 | 라이브러리 |
|---|---|
| GUI 프레임워크 | `customtkinter` + `tkinter` |
| 드래그 앤 드롭 | `tkinterdnd2` |
| DXF 파싱 | `ezdxf` |
| DWG → DXF 변환 | ODA File Converter (외부 EXE, `C:\Program Files\ODA`) |
| 데이터 처리 | `pandas` |
| 엑셀 출력 | `openpyxl` |
| EXE 패키징 | `PyInstaller` |

---

## app.py 구조 (섹션 순서)

```
[0] 로깅 설정 / 전역 상수
[1] 공통 유틸리티
    - _도면번호_패턴, _축척_패턴, _뷰_축척_타입_패턴, _동_패턴 (정규식)
    - GLOBAL_IGNORE_HEADERS (헤더 노이즈 제거 목록)
    - _clean_text_from_headers(), _도면번호_세척(), _merge_title_char_runs()
    - _extract_dong_from_title(), _extract_group_from_title()
    - _spatial_reconstruct_num_str() (공간 좌표 기반 하이픈 복원)
[2] ROI / 좌표 변환 엔진
    - load_roi_config() — JSON 로드 (cp949/utf-8/euc-kr 순 시도)
    - _oda_환경_설정() — ODA EXE 경로 자동 탐색 및 ezdxf 연결
    - _find_도곽_blocks() — 정확일치 우선, 부분일치 폴백
    - _roi_to_abs() — 비율 ROI → 절대 좌표 변환 (회전 도곽 보정 포함)
[3] DXF 파싱 엔진
    - _extract_texts_in_roi() — 지정 ROI 내 TEXT/MTEXT/ATTRIB 추출
    - _extract_view_symbols() — 원+우측수평선 구조 뷰심볼 감지
    - _parse_scale_text() — A1/A3 축척 분리 파싱
[4] 도면목록표 파싱
    - _parse_list_dwg() — 다단(Multi-Column) ROI 순차 스캔
    - _apply_spatial_indent_inheritance() — 들여쓰기 동 정보 상속
[5] 개별 도면 처리
    - _process_single_dwg() — 단일 DWG 분석 (traceback 포함 예외 로깅)
    - run_check() — ProcessPoolExecutor 멀티프로세싱 오케스트레이션
[6] 엑셀 리포트 생성
    - _write_report() — 시트1(목록표 검토) + 시트2(뷰심볼 검토) 컬러 출력
[7] GUI
    - AutoDWGApp (CTk 메인 윈도우)
    - GUILogHandler — after(0, ...) 스레드 안전 로그 핸들러
```

---

## 데이터 흐름

```
AutoCAD → SET_ROI.lsp → %APPDATA%\AutoDWG_Checker\<블록명>.json
                                        ↓
DWG 파일들 → ODA File Converter → 임시 DXF → ezdxf 파싱
                                        ↓
              _find_도곽_blocks() → ROI 좌표 변환 → 텍스트/뷰심볼 추출
                                        ↓
              도면목록표 파싱 결과와 pandas merge (키: 도면번호)
                                        ↓
              openpyxl 컬러 리포트 → 도면검토리포트_최종.xlsx
```

---

## 주요 설계 결정 및 제약

### ROI(JSON) 구조
SET_ROI.lsp가 저장하는 JSON 형식:
```json
{
  "base_w": 594.0,        // 도곽 원본 너비 (mm)
  "base_h": 420.0,        // 도곽 원본 높이 (mm)
  "num_roi":   [x1, x2, y1, y2],  // 도면번호 영역 (비율 0~1)
  "title_roi": [x1, x2, y1, y2],  // 도면명 영역
  "scale_roi": [x1, x2, y1, y2],  // 축척 영역
  "list_rois": [[...], [...], ...], // 도면목록표 다단 (Column별 배열)
  "view_symbol_roi": [x1, x2, y1, y2]  // 뷰심볼 탐색 영역 (null이면 생략)
}
```
- 모든 좌표는 도곽 삽입점 기준 **비율값** (0.0 ~ 1.0 범위)
- `view_symbol_roi`가 `null`이면 시트 2(뷰심볼 검토)가 생성되지 않음

### 뷰심볼 인식 조건
`_extract_view_symbols()`에서 뷰심볼로 인정되려면:
- 원(Circle) + 우측 수평 LINE 구조
- 우측 연장 ≥ 원 반지름 × 2
- 좌측 연장 < 우측 연장 × 0.5 (비대칭, 구조 십자선 제외용)

좌측/양쪽 대칭/박스/다이아몬드형 뷰심볼은 미감지 → KNOWN_ISSUES #3

### 도면번호 패턴
`_도면번호_패턴` 정규식 기준:
- 영문 prefix 최대 5자 (`{0,4}`)
- `AA-000-000` 형식, 무한 체인 지원 (`AA-000-000-000`)
- CAD에서 한 글자씩 분리 저장된 경우 `_spatial_reconstruct_num_str()`로 복원

### 멀티프로세싱 취소
`threading.Event` + `ProcessPoolExecutor.shutdown(cancel_futures=True)` 조합.
취소 버튼(⏹) 클릭 → `_cancel_event.set()` → 풀 강제 종료.

### ODA 미설치 시
`sys.exit()` 대신 `askretrycancel()` retry loop — 사용자가 설치 후 재시도 가능.

---

## 빌드 방법

```bat
:: 1. 가상환경 의존성 설치 (최초 1회)
venv\Scripts\pip install -r requirements.txt

:: 2. EXE 빌드
build.bat
:: → dist\DWGChecker.exe 생성
```

- `build.spec`에 `SET_ROI.lsp`가 **포함되어 있지 않음** → 배포 시 별도 동봉 필요 (KNOWN_ISSUES #7)
- UPX 압축 활성화 (`upx=True`), 콘솔 창 없음 (`console=False`)

---

## 코딩 컨벤션

- **변수명:** 직관적인 영어 사용. 도메인 용어는 한글 허용 (예: `도면번호`, `블록명`)
- **주석:** 각 처리 단계마다 목적 주석 작성
- **코드 출력:** 수정 시 전체 함수/섹션을 완성된 형태로 출력 (스니펫 지양)
- **스코프:** 명시적으로 요청된 기능만 수정 — 테스트 코드, 파일 분리, 변수 rename, 폴더 정리 불필요
- **에러 처리:** `_process_single_dwg`에서는 반드시 `traceback.format_exc()` 포함 로깅

---

## 알려진 제약사항 요약 (KNOWN_ISSUES.md 참조)

| # | 항목 | 영향도 |
|---|---|---|
| 1 | ROI 의존성 (JSON 필수) | 도곽 바뀌면 재설정 필요 |
| 2 | 도곽이 LINE/PLINE이면 인식 0건 | 블록/XREF 도곽만 지원 |
| 3 | 뷰심볼 인식 조건 까다로움 | 비표준 형태 미감지 |
| 4 | ODA 설치 경로 고정 | 기본값 경로만 자동 탐색 |
| 5 | 중복 도면번호 두 번째 행 소실 | 경고 로그는 출력됨 |
| 6 | ROI 미리보기 없음 | 시행착오 시 캐드 왕복↑ |
| 7 | LISP 파일 EXE 미포함 | 배포 시 별도 동봉 필요 |
| 8 | 동/그룹 휴리스틱 강함 | 순수 한글 첫 토큰 = 그룹 판정 |
