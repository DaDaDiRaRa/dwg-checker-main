# AutoDWG 도면 검토 자동화 시스템 — CLAUDE.md

## 프로젝트 개요

실시설계 도서 납품 전, **도면목록표(LIST DWG)** 와 **개별 캐드 도면(DWG 수백 장)** 을
자동으로 교차 검토하여 컬러 엑셀 리포트를 생성하는 Windows 단독 실행(EXE) 앱.

- **현재 버전:** v6.8
- **진입점:** `app.py` (단일 파일, ~1600줄)
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
├── build.spec          # PyInstaller 스펙 (git 미추적 — .gitignore 등록)
├── requirements.txt    # Python 의존성
├── readme.html         # 사용 설명서 (이미지 base64 내장, 단독 배포 가능)
├── readme.md           # 사용 설명서 Markdown 버전
├── KNOWN_ISSUES.md     # 수정 보류 중인 제약사항 8개 문서화
├── image/              # readme.html/md에 삽입된 스크린샷 원본
└── venv/               # Python 가상환경 (빌드 전 의존성 설치 필요, git 미추적)
```

> `build.spec`, `build/`, `cfg_list.txt`, `xcfg.txt`는 `.gitignore`에 등록되어 git에서 추적되지 않는다.

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

## app.py 구조 (실제 함수명 기준)

```
[0] 로깅 설정 / 전역 상수
    - _setup_file_logger()
    - 리포트_이름, ODA_DOWNLOAD_URL

[1] 공통 유틸리티
    정규식: _도면번호_패턴, _축척_패턴, _뷰_축척_타입_패턴, _동_패턴
    상수:   GLOBAL_IGNORE_HEADERS, CATEGORY_KEYWORDS
    함수:
    - _clean_text_from_headers()   — 헤더 노이즈 제거
    - _extract_dong_from_title()   — 동(棟) 정보 추출
    - _extract_group_from_title()  — 그룹(분류) 정보 추출
    - _도면번호_세척()              — 대소문자 정규화, 대시 복원
    - _spatial_reconstruct_num_str() — 공간 좌표 기반 하이픈 복원
    - _merge_title_char_runs()     — 한 글자씩 분리된 도면명 합치기
    - _축척_텍스트_정리()           — 축척 문자열 → "1/N" 표준화
    - _정리문자열()                 — 공백 정규화
    - _expand_title_keywords()     — 쉼표 축약형 도면명 확장 (입,단면도 → 입면도/단면도)
    - _title_contains_view()       — 뷰심볼 도면명이 도곽 도면명에 포함되는지 확인
    - _extract_drawing_number()    — 텍스트에서 도면번호 추출

[2] CAD 로드 / 도곽 탐색
    - load_roi_config()            — JSON 로드 (cp949/utf-8/euc-kr 순 시도)
    - _oda_환경_설정()              — ODA EXE 경로 자동 탐색 및 ezdxf 연결
    - _cad_로드()                  — DXF 직접 읽기 또는 ODA를 통한 DWG 변환
    - _find_도곽_blocks()          — 정확일치 우선, 부분일치 폴백
    - _get_safe_point()            — ATTRIB 정렬점 안전 추출
    - _텍스트_데이터_추출()          — TEXT/MTEXT/ATTRIB/ATTDEF 단일 엔티티 파싱
    - _collect_layout_texts()      — 레이아웃 전체 텍스트 수집 (INSERT 재귀 포함)
    - _parse_xref_original()       — XREF 원본 파일에서 고정 텍스트 암기
    - _transform_xref_texts()      — XREF 텍스트 WCS 좌표 변환

[3] 뷰심볼 + 축척 파싱
    - _extract_view_symbols()      — 원+우측수평선(TEXT/MTEXT) 방식 +
                                     INSERT 블록 ATTRIB 방식 두 경로로 뷰심볼 감지
    - _clean_title_only()          — 도면명에서 축척 문자열 제거
    - _extract_scale_smart()       — A1/A3 축척 쌍 추출
                                     (헤더 위치 기반 최근접 매칭, 목록표/개별도면 모드 분기)

[4] 도면목록표 파싱
    - extract_dwg_list_table()     — 다단(Multi-Column) ROI 순차 스캔
                                     1차 패스: 행별 도면번호/도면명/그룹 후보 계산
                                     2차 패스: 연속 동일 후보 → 그룹 아닌 도면명 일부로 처리

[5] 개별 도면 처리
    - _process_single_dwg()        — 단일 DWG 분석
                                     내부 클로저 get_data_in_roi()로 ROI 텍스트 추출
                                     traceback 포함 예외 로깅, doc 은 finally에서 해제
    - extract_dwg_data_multiprocess() — ProcessPoolExecutor 멀티프로세싱 오케스트레이션

[6] 엑셀 리포트 생성
    - _build_view_sheet()          — 시트2(뷰심볼 검토) 컬러 출력
    - build_report()               — 시트1(목록표 검토) 생성 후 wb.save()
                                     view_df가 있으면 시트2 추가

[7] GUI
    - GUILogHandler                — after(0, ...) 스레드 안전 로그 핸들러
    - AutoDWGApp                   — ctk.CTk + TkinterDnD.DnDWrapper 결합 메인 윈도우
    - _ensure_oda_installed()      — ODA 미설치 시 askretrycancel() retry loop
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
              도면목록표 파싱 결과와 pandas merge (키: 도면번호 정규화 KEY)
                                        ↓
              openpyxl 컬러 리포트 → 도면검토리포트_최종.xlsx
                                   시트1: 목록표 검토
                                   시트2: 뷰심볼 검토 (view_symbol_roi 설정 시)
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
- ROI는 항상 **Master(도면목록표) 도곽** 기준으로 저장됨

### Master / Slave 도곽 이름 분리

GUI 체크박스("개별 도면의 도곽 이름이 다른 경우 체크")로 활성화.

- **Master 블록**: 도면목록표 스캔 + ROI JSON 키 — `extract_dwg_list_table()` 사용
- **Slave 블록**: 개별 도면 탐색 — `extract_dwg_data_multiprocess()` 사용
- 체크박스 미사용 시 Master = Slave

### 뷰심볼 인식 방식 (두 경로)

`_extract_view_symbols()`에서 다음 두 방식을 모두 시도:

1. **TEXT/MTEXT 방식**: 원(Circle) + 우측 수평 LINE 구조 감지
   - 우측 연장 ≥ 원 반지름 × 2
   - 좌측 연장 < 우측 연장 × 2 (구조 십자선 제외)
   - line_y 위 = 도면명, line_y 아래 = 축척
2. **ATTRIB 방식**: INSERT 블록 내부에 CIRCLE + ATTRIB로 구성된 뷰심볼
   - 블록 정의에 CIRCLE이 있어야 뷰심볼로 간주
   - ATTRIB 삽입점 y 기준으로 위/아래 분리

좌측/양쪽 대칭/박스/다이아몬드형 뷰심볼은 미감지 → KNOWN_ISSUES #3

### 도면번호 패턴

`_도면번호_패턴` 정규식 기준:

- 영문/한글/그리스어 prefix 최대 5자 (`{0,4}`)
- `AA-000-000` 형식, 무한 체인 지원 (`AA-000-000-000`)
- CAD에서 한 글자씩 분리 저장된 경우 `_spatial_reconstruct_num_str()`로 복원

### merge KEY 정규화

`build_report()`에서 LIST와 DWG 도면번호를 합칠 때:

```python
KEY = 도면번호.upper().replace(r"[\s\-_]", "")
```

공백·하이픈·언더스코어를 모두 제거한 문자열로 outer join.

### 멀티프로세싱 취소
`threading.Event` + `ProcessPoolExecutor.shutdown(cancel_futures=True)` 조합.
취소 버튼(⏹) 클릭 → `self.cancel_event.set()` → 잔여 future 취소.
Python 3.8 이하 호환을 위해 `cancel_futures` 인자 없는 경우 폴백 처리.

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
- `build.spec`은 로컬에 존재하지만 `.gitignore`로 git 미추적 (빌드 산출물과 함께 관리)

---

## 코딩 컨벤션

- **변수명:** 직관적인 영어 사용. 도메인 용어는 한글 허용 (예: `도면번호`, `블록명`)
- **주석:** 각 처리 단계마다 목적 주석 작성
- **코드 출력:** 수정 시 전체 함수/섹션을 완성된 형태로 출력 (스니펫 지양)
- **스코프:** 명시적으로 요청된 기능만 수정 — 테스트 코드, 파일 분리, 변수 rename, 폴더 정리 불필요
- **에러 처리:** `_process_single_dwg`에서는 반드시 `traceback.format_exc()` 포함 로깅, `doc`은 `finally`에서 해제

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
