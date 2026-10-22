@echo off
REM Arranca la app con doble clic (usa uv + Python 3.12, instala las librerías solo si faltan)
cd /d "%~dp0"
uv run --python 3.12 --with flask --with openpyxl --with fpdf2 --with python-dotenv python app.py
pause
