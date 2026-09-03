@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal

rem 先找 py 启动器，再退回 python，兼容只装了其中一种的机器
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)

if not defined PY (
    echo.
    echo  [!] Python not found. Install Python 3.9+ and tick "Add Python to PATH".
    echo  [!] 没有找到 Python。请安装 Python 3.9 或更高版本，
    echo      安装时务必勾选 "Add Python to PATH"。
    echo.
    pause
    exit /b 1
)

rem 自动挑版本号最高的界面脚本
set "GUI="
for /f "delims=" %%f in ('dir /b /o-n "pdiv_report_gui_v*.py" 2^>nul') do (
    set "GUI=%%f"
    goto :run
)

echo.
echo  [!] pdiv_report_gui_v*.py not found in this folder.
echo  [!] 同目录下没找到 pdiv_report_gui_v*.py
echo.
pause
exit /b 1

:run
echo  Starting / 正在启动: %GUI%
%PY% "%GUI%"
if errorlevel 1 pause
endlocal
