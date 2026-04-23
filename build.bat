@echo off
chcp 65001 > nul
echo.
echo ============================================================
echo  AutoDWG EXE 빌드 스크립트
echo ============================================================
echo.

:: PyInstaller 설치 확인
python -m PyInstaller --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] PyInstaller 가 없습니다. 설치 중...
    pip install pyinstaller
    echo.
)

:: 패키지 경로 자동 탐색
echo [INFO] 패키지 경로 탐색 중...
for /f "tokens=*" %%i in ('python -c "import customtkinter, os; print(os.path.dirname(customtkinter.__file__))"') do set CTK_PATH=%%i
for /f "tokens=*" %%i in ('python -c "import tkinterdnd2, os; print(os.path.dirname(tkinterdnd2.__file__))"') do set DND_PATH=%%i

echo  CustomTkinter : %CTK_PATH%
echo  TkinterDnD2   : %DND_PATH%
echo.

:: 이전 빌드 정리
if exist dist\도면검토기.exe (
    echo [INFO] 이전 EXE 파일 삭제 중...
    del /q dist\도면검토기.exe
)

echo [INFO] 빌드 시작 (수 분 소요될 수 있습니다)...
echo.

pyinstaller --onefile --windowed ^
  --name "도면검토기" ^
  --add-data "%CTK_PATH%;customtkinter" ^
  --add-data "%DND_PATH%;tkinterdnd2" ^
  --hidden-import customtkinter ^
  --hidden-import tkinterdnd2 ^
  --hidden-import PIL ^
  --hidden-import PIL.Image ^
  --hidden-import PIL.ImageTk ^
  --collect-submodules ezdxf ^
  --hidden-import pandas ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.styles ^
  --hidden-import openpyxl.utils ^
  app.py

echo.
if exist dist\도면검토기.exe (
    echo ============================================================
    echo  빌드 완료!  dist\도면검토기.exe 를 배포하세요.
    echo ============================================================
    start "" "dist"
) else (
    echo [ERROR] 빌드 실패. 위의 오류 메시지를 확인하세요.
)

pause
