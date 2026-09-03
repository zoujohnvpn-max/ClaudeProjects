@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo  没有找到 Python。请先安装 Python 3.9 或更高版本，
    echo  安装时务必勾选 "Add Python to PATH"。
    echo.
    pause
    exit /b 1
)

set "GUI="
for /f "delims=" %%f in ('dir /b /o-n "pdiv_report_gui_v*.py" 2^>nul') do (
    set "GUI=%%f"
    goto :run
)

echo  同目录下没找到 pdiv_report_gui_v*.py
pause
exit /b 1

:run
python "%GUI%"
if errorlevel 1 pause
