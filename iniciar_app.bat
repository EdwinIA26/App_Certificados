@echo off
REM Arranca la app con doble clic (usa uv + Python 3.12, instala las librerías solo si faltan)
cd /d "%~dp0"
REM Espera a que el servidor responda y abre el navegador automáticamente
start "" powershell -NoProfile -WindowStyle Hidden -Command "$u='http://127.0.0.1:5000'; for($i=0;$i -lt 60;$i++){ try { Invoke-WebRequest $u -UseBasicParsing -TimeoutSec 1 | Out-Null; break } catch { Start-Sleep -Seconds 1 } }; Start-Process $u"
uv run --python 3.12 --with flask --with openpyxl --with fpdf2 --with python-dotenv python app.py
pause
