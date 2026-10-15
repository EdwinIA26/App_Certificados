# AGENTS.md

## Contexto

Academia Horizonte: centro de formación, programas técnicos a empresas. Proceso manual Excel/Word; objetivo: app web para certificados por cohorte.
Idioma de trabajo: español sencillo en explicaciones y comentarios.

## Arquitectura (3 piezas)

- Bodega: `Insumos/` — datos fuente, NO modificar.
- Cocina: `app.py` (Python + Flask) — TODA la lógica aquí, NUNCA en la interfaz.
- Salón: HTML servido por Flask — solo muestra, no calcula ni decide.

## Entradas

- `Insumos/Maestro_Estudiantes.xlsx`: Identificacion, Nombre_Completo, Correo, Programa, Cohorte.
- `Insumos/Registro_Evaluaciones.xlsx`: Identificacion, Programa, Modulo, Nota (0-100), Asistencia_Pct, Fecha_Cierre.

## Reglas de negocio

- Agrupar por Identificacion + Programa. Un certificado por estudiante y programa.
- Promedio = suma Notas / cantidad módulos. Asistencia = promedio Asistencia_Pct.
- Aprobación: Promedio >= 70 y Asistencia >= 80. Participación: Promedio < 70 y Asistencia >= 80. Sin certificado: Asistencia < 80. Límites inclusivos.

## Estilo

- Python simple y comentado en español, marcando variables, tipos, condicionales, bucles y funciones.
- Web azul marino y dorado.

## Decisiones técnicas (ver PLAN.md)

- Excel: `openpyxl` (ligera, directa; `pandas` descartado por pesado, `xlrd` por obsoleto).
- Web: `Flask` + Jinja incluido (descartados Django por overkill, FastAPI por API-JSON).
- Estructura: `app.py`, `requirements.txt`, `templates/index.html`, `templates/certificado.html`, `static/style.css`. Lógica solo en `app.py`.
- Llave: `Identificacion + Programa` como texto; divisor dinámico; UTF-8 obligatorio.

## Estado

App construida: `app.py` + `templates/` + `static/` + `requirements.txt` + `iniciar_app.bat`. Arranque: doble clic en `iniciar_app.bat` o `uv run --python 3.12 --with flask --with openpyxl python app.py`. Shell: sintaxis PowerShell (workspace Windows).
