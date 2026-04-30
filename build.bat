@echo off
chcp 65001 > nul
echo ========================================
echo   AutoDWG 도면검토기 EXE 빌드 시작
echo ========================================

:: 이전 빌드 결과물 정리
if exist "dist\도면검토기.exe" (
    del /f /q "dist\도면검토기.exe"
    echo [정리] 이전 EXE 삭제 완료
)
if exist "build" (
    rmdir /s /q "build"
    echo [정리] build 폴더 삭제 완료
)

echo.
echo [빌드] PyInstaller 실행 중... (수 분 소요)
echo.

venv\Scripts\pyinstaller.exe --clean build.spec

echo.
if exist "dist\도면검토기.exe" (
    echo ========================================
    echo   빌드 성공!
    echo   dist\도면검토기.exe 생성 완료
    echo ========================================
    explorer dist
) else (
    echo ========================================
    echo   빌드 실패. 위 오류 메시지를 확인하세요.
    echo ========================================
)

pause
