@echo off
REM Arranca la app con doble clic (usa uv + Python 3.12, instala Flask y openpyxl solo si faltan)
cd /d "%~dp0"
uv run --python 3.12 --with flask --with openpyxl python app.py
pause
