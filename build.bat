@echo off
echo ========================================
echo   AutoDWG Checker - EXE Build Start
echo ========================================

:: Remove previous build artifacts
if exist "dist" (
    rmdir /s /q "dist"
    echo [Clean] dist folder removed
)
if exist "build" (
    rmdir /s /q "build"
    echo [Clean] build folder removed
)

echo.
echo [Build] Running PyInstaller... (may take a few minutes)
echo.

venv\Scripts\pyinstaller.exe --clean build.spec

echo.
if exist "dist\DWGChecker.exe" (
    echo ========================================
    echo   BUILD SUCCESS!
    echo   dist\DWGChecker.exe created
    echo ========================================
    explorer dist
) else (
    echo ========================================
    echo   BUILD FAILED. Check errors above.
    echo ========================================
)

pause
