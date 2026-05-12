#role

<FILL: ...> 자리만 덮어쓰면 됩니다 — 구조 자체는 건드리지 마세요. PPT 슬라이드 자동 매핑을 위해 섹션 번호와 키 이름을 고정해두었습니다.
자연어 문장 금지, 키-값 / 태그 / 룰 형식 유지 (예: before_manual_check: 4시간/건).
모르는 항목은 <FILL>을 그대로 두시거나 TBD로 표시. 빈 칸이 있어도 PPT 기획은 진행 가능합니다.
스크린샷은 파일 경로만 적어주시면 됩니다. 실제 이미지는 나중에 한 번에 업로드. 스크린샷 없으면 없는대로 넘어가.

# BRIEF: AutoDWG Cross-Checker

> Windows Desktop App 타입 — 도면목록표 ↔ 개별 DWG 교차검증 자동화
> 채우는 방식: `<FILL: ...>` 자리에 최신화된 내용을 덮어쓰기

---

## 1. META

```yaml
app_name: AutoDWG Cross-Checker
korean_name: 도면 교차검증 자동화 도구
version: v6.8 (+ 2026-05-12 hotfix)
type: Windows Desktop App
target_user: DA / 실시설계 도서 납품 담당 / 인허가 도서 검수자
problem_solved: 도면목록표 ↔ 수백 장 개별 DWG 수기 교차검토 자동화
core_value_proposition: ODA + ezdxf + 멀티프로세싱 + XREF X-Ray 합성으로 누락/오타/축척오류 일괄 탐지
status: 내부 베타
deployment_format: .exe (PyInstaller 단일파일) + Python 소스
maintainer: 김정현 (junghyunk9966@gmail.com)
last_updated: 2026-05-12
```

---

## 2. ARCHITECTURE_LAYERS

```
- Layer 0 (Config):
    - load_roi_config        : %APPDATA%/AutoDWG_Checker/{블록명}.json (cp949/utf-8/euc-kr 폴백)
    - oda_env_setup          : C:/Program Files/ODA/**/ODAFileConverter.exe 자동 탐색 + PATH 등록
    - file_logger            : %APPDATA%/AutoDWG_Checker/autodwg.log (DEBUG)
    - ensure_oda_installed   : 미설치 시 다운로드 페이지 + 재시도 루프

- Layer 1 (Common Utils):
    - regex_engine           : _도면번호_패턴 / _축척_패턴 / _뷰_축척_타입_패턴 / _동_패턴
    - ignore_headers_list    : GLOBAL_IGNORE_HEADERS (한·영 40+종)
    - horizontal_text_merge  : _merge_title_char_runs / _spatial_reconstruct_num_str (한 글자 분리 텍스트 재결합)
    - title_cleaners         : _도면번호_세척 / _축척_텍스트_정리 / _clean_title_only
    - dong_group_extract     : _extract_dong_from_title / _extract_group_from_title

- Layer 2 (DXF/DWG Parsing Core):
    - collect_layout_texts   : TEXT/MTEXT/LINE/LWPOLYLINE/INSERT/ATTDEF 재귀 수집
    - parse_xref_original    : XREF 원본 DWG modelspace 텍스트 추출
    - transform_xref_texts   : scale + rotation + translation 매트릭스 적용
    - get_safe_point         : halign/valign align_point 보정
    - find_도곽_blocks       : 정확 일치 우선 / 부분일치 fallback + 경고 로그

- Layer 3 (Analysis Engine):
    - extract_dwg_list_table : 도면목록표 다단(Multi-Column) ROI 스캔 + 도면번호 anchor
    - process_single_dwg     : 개별 DWG 1장 분석 (자식 프로세스, traceback 회수)
    - multiprocess_pool      : concurrent.futures.ProcessPoolExecutor + progress_cb + cancel_event
    - extract_scale_smart    : A1/A3 라벨-값 거리 기반 페어링
    - extract_view_symbols   : 원+우측 수평선 뷰심볼 + ATTRIB 블록형 동시 인식

- Layer 4 (Report):
    - build_report           : pandas outer merge + 셀별 빨간색 칠 + 상태 컬럼 ("도면명/축척 불일치" 등)
    - excel_highlight_rule   : PatternFill FFFF9999 (불일치) / FFFFD699 (중복) / FFD6E4F7 (헤더)
    - build_view_sheet       : 뷰심볼 검토 시트 (도면명 포함 O/X / 축척 일치 O/X/?)

- Layer 5 (GUI):
    - customtkinter_window   : AutoDWGApp = ctk.CTk + TkinterDnD.DnDWrapper
    - stdout_redirector      : GUILogHandler (logging.Handler → CTkTextbox, after() 마샬링)
    - threading_model        : main(GUI) ↔ worker(run_core_logic) ↔ ProcessPool(자식 N개)
    - drag_and_drop          : tkinterdnd2 DND_FILES (파일/폴더 다중 드롭)
    - progress_and_cancel    : determinate progressbar + threading.Event 취소 신호
```

---

## 3. PARSING_PIPELINE

```
[AutoCAD]
    ↓ SET_ROI LISP macro
[ROI JSON @ %APPDATA%/AutoDWG_Checker/{블록명}.json]
    ↓
[ODA File Converter]
    ↓ DWG → DXF
[ezdxf parser]
    ↓ TEXT / MTEXT / INSERT(ATTRIB) / ATTDEF 재귀 탐색
[XREF X-Ray composite]
    ↓ 좌표 변환 + 가상 합성
[Cross-check engine]
    ↓ Drawing list ↔ Individual DWG (multiprocessing)
[openpyxl report]
    ↓ Red-highlighted mismatches
[Excel output]
```

---

## 4. XREF_TRANSFORM_LOGIC

```
- detect_xref_block       : layout.query("INSERT") + _find_도곽_blocks (정확/부분일치)
- read_xref_original_dxf  : _parse_xref_original (modelspace TEXT/MTEXT/ATTDEF/INSERT 가상엔티티 재귀)
- insert_point_matrix     : ins.dxf.insert → (ix, iy) translation 기준점
- scale_matrix            : abs(ins.dxf.xscale) / abs(ins.dxf.yscale)
- rotation_matrix         : ins.dxf.rotation(deg) → cos/sin (90°/270° 회전 도곽 자동 보정)
- composite_to_host       : (sx·cos − sy·sin + ix, sx·sin + sy·cos + iy) + 높이 yscale 스케일
- view_symbol_unrot       : 뷰심볼 ROI도 동일 역회전 적용 (v6.8 회전도곽 보정)
- block_internal_attribs  : INSERT 내부 ATTRIB / 블록정의 CIRCLE WCS 변환 (뷰심볼 ATTRIB 양식)
```

---

## 5. ROI_DETECTION

```
- block_name_anchor       : Master(목록표) / Slave(개별도면) 블록명 분리, JSON 파일명 = 블록명
- multi_column_scan       : list_rois 배열 → 단(Column)별 별도 비율 박스 (3단 이상 지원)
- row_grouping_algorithm  : y좌표 클러스터링 (높이·0.012 임계) + 도면번호 anchor 기준 sub_line 묶음
- drawing_number_anchor   : "도면번호/DWG.NO/번호" 텍스트 X좌표 평균 → header_num_x
- header_alias_a1_a3      : "A1"/"A3" 라벨 좌표 = header_a1_x / header_a3_x
- hyphen_dash_recovery    : 도면번호 column 내 "-" 단일자 → 인접 행 y좌표로 흡수 (그래픽 대시 복원)
- spatial_indent_inherit  : 들여쓰기 행 → 위 행의 동/그룹 정보 자동 상속
- category_row_skip       : "공통사항/건축도면/기계도면" 등 분류행은 도면번호 없으면 건너뜀
```

---

## 6. CROSS_CHECK_RULES

```
- KEY_normalization        : str.upper().replace(r"[\s\-_]", "") (pandas regex)
- merge_strategy           : outer merge on KEY (pandas.merge, indicator=True)
- mismatch_categories:
    - drawing_number       : KEY 정규화 후 LIST≠DWG → "도면번호 불일치"
    - scale (1/100 등)     : A1/A3 각각 공백 제거 비교 → "축척 불일치"
    - building_block(동)   : _동_패턴 (102동/A동/가동 등) → "그룹 불일치"
    - drawing_title        : replace(" ","") 비교 → "도면명 불일치"
    - view_symbol_title    : 뷰 도면명 단어집합 ⊆ 도곽 도면명 (쉼표축약 확장 포함)
    - view_symbol_scale    : 뷰 A1/A3 ↔ 도곽 A1/A3 매핑 일치
    - duplicate_view_name  : 같은 파일 내 동일 뷰명 2회+ → "중복" (오타 의심, 주황색)
- ignore_headers           : SUBJECT TITLE / PROJECT TITLE / DRAWING NO. / DWG.NO / SHEET NO / TITLE / 도면번호 / 도면명 / 축척(A1) / 축척(A3) / SCALE(A1) / SCALE(A3) / 비고 / REMARK / 사업승인 / 착공 / 견적 / 사용승인 등 40+
- safe_bidirectional_mode  : 단방향 비교만 (DWG 누락 / 목록표 누락 명시) — 어느 쪽도 자동 수정·덮어쓰기 X
```

---

## 7. GUI_FLOW

```
[Step 1] : AutoCAD에서 SET_ROI 명령 → 도곽 블록 선택 → 도면번호/도면명/축척/뷰심볼/단(Column) 박스 드래그 (최초 1회)
[Step 2] : 도면검토기.exe 실행 → ODA 자동 검색 (미설치 시 다시시도 다이얼로그)
[Step 3] : ① 도곽 원본 DWG 드래그/선택 → 블록명 자동입력 + (옵션) 개별도면 도곽명 별도 입력
[Step 4] : ② 도면목록표 DWG / ③ 개별 도면 폴더 드래그/선택 (다중 폴더 가능)
[Step 5] : [검토 시작 →] 클릭 → 결정형 진행률 + 실시간 로그 + [⏹ 취소] 버튼
[Step 6] : 도면검토리포트_최종.xlsx 생성 → 폴더 자동 오픈 (목록표 검토 / 뷰심볼 검토 2시트)
```

---

## 8. SCREENSHOTS

```
- /screenshots/01_main_window.png        : TBD
- /screenshots/02_lisp_in_autocad.png    : TBD
- /screenshots/03_file_selection.png     : TBD
- /screenshots/04_progress_log.png       : TBD
- /screenshots/05_excel_report.png       : TBD
- /screenshots/06_mismatch_highlight.png : TBD
```

---

## 9. PERFORMANCE

```
- avg_processing_time_per_dwg : TBD
- multiprocessing_workers     : os.cpu_count() (ProcessPoolExecutor 기본값)
- max_tested_drawing_count    : TBD
- memory_footprint            : TBD
- before_manual_check         : TBD (참고: 수 시간/건)
- after_auto_check            : 수십초 ~ 수 분/건 (README 기준)
- time_saved_ratio            : TBD %
- error_detection_rate        : TBD %
```

---

## 10. DEPENDENCIES

```
- python              : 3.10+
- ezdxf               : >=1.3.0
- pandas              : >=2.0.0
- openpyxl            : >=3.1.0
- customtkinter       : >=5.2.0
- tkinterdnd2         : >=0.4.0
- ODA File Converter  : C:\Program Files\ODA\<버전>\ODAFileConverter.exe (필수, 기본 경로)
- AutoCAD LISP        : SET_ROI.lsp (프로젝트 루트, AutoCAD APPLOAD or D&D)
- build_tool          : PyInstaller (build.spec, --clean)
- runtime_os          : Windows 10/11 (os.startfile 의존)
```

---

## 11. ERROR_HANDLING

```
- ODA_not_found       : _ensure_oda_installed 루프 (askretrycancel + webbrowser.open + 설치완료 확인)
- corrupt_DWG         : _process_single_dwg 자식 try/except → traceback 메인 logger.debug
- missing_ROI_json    : load_roi_config → None → logger.error + 작업 중단 메시지
- empty_list_table    : logger.warning ("도곽 블록 미발견" / "추출 0건" / "도면번호 중복 N행 발견")
- xref_path_broken    : _parse_xref_original try/except → logger.error + 빈 리스트 반환 (분석은 계속)
- encoding_failure    : load_roi_config cp949 → utf-8 → euc-kr 순차 폴백
- excel_locked        : PermissionError → "엑셀 창을 닫고 다시 실행" 안내
- thread_safety       : GUILogHandler.emit → self.after(0, ...) 메인 스레드 마샬링
- block_name_mismatch : _find_도곽_blocks 정확 일치 0건 → 부분일치 fallback + 매칭블록 경고 출력
- worker_cancel       : ProcessPoolExecutor.shutdown(cancel_futures=True) (Python 3.9+) / wait=False fallback
```

---

## 12. SECURITY_AND_DEPLOYMENT

```
- local_only_execution : 외부 통신 없음 (webbrowser.open으로 ODA 다운로드 페이지만 호출)
- data_residency       : 모든 DWG / DXF / 리포트는 사용자 PC 로컬, 네트워크 업로드 X
- distribution_channel : 사내 공유폴더 EXE 배포 (PyInstaller 단일파일 + SET_ROI.lsp 동봉 필요)
- update_mechanism     : 수동 재배포 (자동 업데이트 채널 없음)
- config_storage       : %APPDATA%/AutoDWG_Checker/ (도곽별 JSON + autodwg.log)
- code_signing         : TBD (build.spec codesign_identity=None)
```

---

## 13. CHANGELOG

```
- v6.0 (X-Ray XREF)             : 외부참조 텍스트 host DWG 좌표계 가상 합성
- v6.5 (GUI Smart Edition)      : customtkinter 윈도우 앱 전환, ATTDEF 레이더, 수평 결합기
- v6.6 (도면목록표 정밀 파싱)   : 공간 좌표 하이픈 복원, 한 글자 분리 텍스트 통합, 무한체인 정규식
- v6.7 (동 정보 지능형 추출)    : Drag&Drop, 동·구역 자동 분리, 들여쓰기 상속, 단일 EXE 빌드
- v6.8 (뷰심볼 검토 고도화)     : 뷰심볼 검토 시트, A1/A3 분리, 회전도곽 보정, 비표준 축척 표기 확장
- v6.8 hotfix (2026-05-12)      : requirements.txt 의존성 정정, GUILogHandler 스레드 안전, _find_도곽_blocks 정확/부분일치, ODA 재시도 루프, 결정형 진행률, 취소 버튼, 옵션 ④→① 통합, 한글 동 패턴 추가, 자식 traceback 회수
```

---

## 14. PRESENTATION_HOOKS

PPT 발표 시 강조 포인트 (각 항목이 슬라이드 1장 후보)

```
- HOOK_1            : 수기 검토 자동화 — 수 시간/건 → 수 분/건 (검토 시간 90%+ 단축)
- HOOK_2            : XREF X-Ray 합성 — 외부참조 도곽 텍스트를 host 좌표계로 변환·합성해 단일 패스 스캔
- HOOK_3            : 안전한 단방향 비교 — 자동 수정·덮어쓰기 0건, 색상 강조만 / 판단은 사람이 함
- BEFORE_AFTER      : 수기 ~수시간/건 vs 자동 ~수분/건 + 누락·오타·축척·동 4종 불일치 셀별 빨간 강조
- TEAM_ASSET_PLAN   : 사내 공유폴더 단일 EXE + SET_ROI.lsp 동봉, 도곽 양식별 ROI JSON 라이브러리 공유
- LIVE_DEMO_SCENARIO: ① SET_ROI 시연 (1분) → ② EXE에 DWG·폴더 D&D (10초) → ③ 진행률 + 취소 시연 → ④ 빨간 셀 클릭해 원본 도면과 비교
- RISK_DISCUSSION   : ROI 의존성(양식 변경 시 재설정) / 도곽 LINE·PLINE 양식 미지원 / A1·A3 외 축척 미지원 / 뷰심볼은 원+우측수평선 한정 (KNOWN_ISSUES.md 참조)
```
