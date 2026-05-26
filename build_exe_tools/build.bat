@echo off
REM build.bat — Build CPAuto.exe using PyInstaller (one-file mode)
REM Can be run from any location; script anchors CWD to project root automatically.
REM Output: dist\CPAuto\  (CPAuto.exe + external runtime files)

setlocal
cd /d "%~dp0\.."

set VENV_DIR=.build_venv
set DIST_DIR=dist\CPAuto

REM ---- Step 1: Create a clean virtual environment -------------------------
echo [BUILD] Creating clean virtual environment...
if exist "%VENV_DIR%" rmdir /S /Q "%VENV_DIR%"
python -m venv "%VENV_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to create virtual environment.
    exit /b 1
)

REM ---- Step 2: Install only required packages -----------------------------
echo [BUILD] Installing required packages...
"%VENV_DIR%\Scripts\pip" install --quiet --upgrade pip
REM Pillow is used internally in openpyxl for fetching image (must install)
"%VENV_DIR%\Scripts\pip" install --quiet PyQt6 lxml openpyxl Pillow pyinstaller
if errorlevel 1 (
    echo [ERROR] Package installation failed.
    exit /b 1
)

REM ---- Step 3: Build with PyInstaller -------------------------------------
echo [BUILD] Running PyInstaller (one-file mode)...
"%VENV_DIR%\Scripts\pyinstaller" build_exe_tools\CPAuto.spec --noconfirm
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    exit /b 1
)

REM ---- Step 4: Assemble distribution folder -------------------------------
echo [BUILD] Assembling distribution folder: %DIST_DIR%\

REM CPAuto.exe is placed directly in dist\ by one-file mode.
REM Move it into our named dist\CPAuto\ folder for clean distribution.
if not exist "%DIST_DIR%" mkdir "%DIST_DIR%"
move /Y "dist\CPAuto.exe" "%DIST_DIR%\CPAuto.exe" >nul 2>&1

REM Copy external runtime files
xcopy /E /I /Y "data" "%DIST_DIR%\data"
xcopy /E /I /Y "assets\CPA_template" "%DIST_DIR%\assets\CPA_template"
xcopy /I /Y "assets\splash_assets\splash.mp4" "%DIST_DIR%\assets\splash_assets\"
xcopy /I /Y "assets\DocGear.ico" "%DIST_DIR%\assets\"

REM Create the output folder (written to at runtime)
if not exist "%DIST_DIR%\output" mkdir "%DIST_DIR%\output"

REM ---- Step 5: Cleanup build artifacts ------------------------------------
echo [BUILD] Cleaning up build artifacts...
rmdir /S /Q build
rmdir /S /Q "%VENV_DIR%"

echo.
echo [BUILD] Done!
echo [BUILD] Distribution folder: %DIST_DIR%\
echo [BUILD] Run: %DIST_DIR%\CPAuto.exe

endlocal
