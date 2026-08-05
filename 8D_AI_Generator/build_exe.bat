@echo off
setlocal
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --onefile --name Adayo8D-AI --add-data "assets\8D模板.pptx;." ai8d.py
echo.
echo EXE created: dist\Adayo8D-AI.exe
pause
