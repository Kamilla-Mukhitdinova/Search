@echo off
setlocal

cd /d %~dp0

if not exist .venv (
    py -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

pyinstaller --noconfirm --windowed --name SearchApp app\main.py

echo.
echo Build completed.
echo Open dist\SearchApp and run SearchApp.exe
pause
