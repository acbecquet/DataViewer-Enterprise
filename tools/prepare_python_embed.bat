@echo off
REM ============================================================
REM  prepare_python_embed.bat
REM  Downloads Python 3.11 embeddable + installs openpyxl.
REM  Output: release\python\  (included by installer.iss)
REM
REM  Run once before building the installer.
REM  Requires: curl (built-in on Win10+), PowerShell
REM ============================================================

setlocal enabledelayedexpansion

set PY_VER=3.11.9
set PY_URL=https://www.python.org/ftp/python/%PY_VER%/python-%PY_VER%-embed-amd64.zip
set GET_PIP_URL=https://bootstrap.pypa.io/get-pip.py

set SCRIPT_DIR=%~dp0
set ROOT=%SCRIPT_DIR%..
set WORK=%ROOT%\tools\_py_work
set OUT=%ROOT%\release\python

echo.
echo === DataViewer Enterprise - Python Embed Preparation ===
echo Python version : %PY_VER%
echo Output         : %OUT%
echo.

REM Clean previous output
if exist "%OUT%" (
    echo Removing previous %OUT% ...
    rmdir /s /q "%OUT%"
)
if exist "%WORK%" rmdir /s /q "%WORK%"
mkdir "%WORK%"
mkdir "%OUT%"

REM ── 1. Download Python embeddable zip ───────────────────────────────────────
echo [1/5] Downloading Python %PY_VER% embeddable ...
curl -L --progress-bar -o "%WORK%\python_embed.zip" "%PY_URL%"
if errorlevel 1 ( echo ERROR: Download failed. & exit /b 1 )

REM ── 2. Unzip into OUT ───────────────────────────────────────────────────────
echo [2/5] Extracting ...
powershell -NoProfile -Command "Expand-Archive -LiteralPath '%WORK%\python_embed.zip' -DestinationPath '%OUT%' -Force"
if errorlevel 1 ( echo ERROR: Extraction failed. & exit /b 1 )

REM ── 3. Enable site-packages (needed for pip / third-party packages) ──────────
echo [3/5] Enabling site-packages ...
REM The _pth file controls sys.path; uncomment "import site" to allow packages.
set PTH_FILE=
for %%f in ("%OUT%\python3*._pth") do set PTH_FILE=%%f
if not defined PTH_FILE (
    echo ERROR: Could not find python3XX._pth in %OUT%
    exit /b 1
)
echo Found pth file: %PTH_FILE%
REM Replace "#import site" with "import site"
powershell -NoProfile -Command "(Get-Content '%PTH_FILE%') -replace '#import site','import site' | Set-Content '%PTH_FILE%'"

REM ── 4. Bootstrap pip ────────────────────────────────────────────────────────
echo [4/5] Installing pip ...
curl -L --progress-bar -o "%WORK%\get-pip.py" "%GET_PIP_URL%"
if errorlevel 1 ( echo ERROR: get-pip.py download failed. & exit /b 1 )
"%OUT%\python.exe" "%WORK%\get-pip.py" --no-warn-script-location
if errorlevel 1 ( echo ERROR: pip bootstrap failed. & exit /b 1 )

REM ── 5. Install openpyxl ─────────────────────────────────────────────────────
echo [5/5] Installing openpyxl ...
"%OUT%\python.exe" -m pip install openpyxl --no-warn-script-location
if errorlevel 1 ( echo ERROR: openpyxl install failed. & exit /b 1 )

REM ── Strip non-runtime packages (pip, setuptools, wheel, packaging) ──────────
REM    These are only needed during this script and cause corruption issues
REM    when packed by Inno Setup's solid compression.
echo Stripping non-runtime packages ...
for %%d in (pip setuptools wheel packaging _distutils_hack) do (
    if exist "%OUT%\Lib\site-packages\%%d" rmdir /s /q "%OUT%\Lib\site-packages\%%d"
)
for %%f in ("%OUT%\Lib\site-packages\pip-*.dist-info"
            "%OUT%\Lib\site-packages\setuptools-*.dist-info"
            "%OUT%\Lib\site-packages\wheel-*.dist-info"
            "%OUT%\Lib\site-packages\packaging-*.dist-info"
            "%OUT%\Lib\site-packages\distutils-precedence.pth") do (
    if exist "%%f" del /q "%%f"
)
if exist "%OUT%\Scripts" rmdir /s /q "%OUT%\Scripts"

REM ── Repackage as zip (avoids Inno Setup binary file corruption) ─────────────
echo Zipping python bundle ...
set ZIP_OUT=%ROOT%\release\python_bundle.zip
if exist "%ZIP_OUT%" del /q "%ZIP_OUT%"
powershell -NoProfile -Command "Compress-Archive -Path '%OUT%\*' -DestinationPath '%ZIP_OUT%'"
if errorlevel 1 ( echo ERROR: Zip failed. & exit /b 1 )

REM ── Cleanup ─────────────────────────────────────────────────────────────────
rmdir /s /q "%WORK%"

echo.
echo === Done! Bundled Python zip is ready at: %ZIP_OUT%
echo     Now build the app (release\DataViewer.exe) then run the Inno Setup compiler.
echo.
endlocal
