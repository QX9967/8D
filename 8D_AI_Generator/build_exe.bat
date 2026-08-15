@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

rem ============================================================
rem  Adayo8D-AI 打包脚本（Win7 兼容）
rem  客户运行环境是 Win7，必须用 Python 3.8 构建。
rem  Python 3.9+ / MSVC 2019+ 工具链会引用 api-ms-win-core-path
rem  (Win8+ 才有的 API set)，在 Win7 上启动即报“丢失 dll”。
rem ============================================================

set "PY38=%LOCALAPPDATA%\Programs\Python\Python38\python.exe"
set "VENV=%~dp0.build_venv38"

if not exist "%PY38%" (
    echo [ERROR] 未找到 Python 3.8.10。
    echo 请下载并安装（选 per-user 安装即可）：
    echo   https://www.python.org/ftp/python/3.8.10/python-3.8.10-amd64.exe
    pause
    exit /b 1
)

if not exist "%VENV%\Scripts\python.exe" (
    echo [..] 创建 Win7 兼容构建环境 .build_venv38 ...
    "%PY38%" -m venv "%VENV%" || goto :err
)

"%VENV%\Scripts\python.exe" -m pip install -r "%~dp0requirements.txt" || goto :err
"%VENV%\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --name Adayo8D-AI --add-data "..\8D模板.pptx;." ai8d.py || goto :err

echo.
echo EXE created: dist\Adayo8D-AI.exe  (Python 3.8 / Win7 兼容)
pause
exit /b 0

:err
echo.
echo [ERROR] 打包失败。
pause
exit /b 1
