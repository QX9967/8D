@echo off
setlocal
set "PYTHON=python"
%PYTHON% -m pip install -r requirements.txt
%PYTHON% -m PyInstaller --noconfirm --clean --onefile --name Adayo8D-AI --add-data "..\8D模板.pptx;." ai8d.py
echo.
echo EXE created: dist\Adayo8D-AI.exe
pause
