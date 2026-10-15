# PLAN.md — App Certificados Academia Horizonte

Modo Plan: este archivo manda. No escribir código hasta validar cada etapa.

## Orden general

1. Etapa 1 — Validar cruce de datos (sin servidor).
2. Etapa 2 — Backend Flask (`app.py`).
3. Etapa 3 — Interfaz HTML (Salón).

No saltar de etapa sin validación del usuario.

## Etapa 1 — Validar cruce de datos

Qué hacer:
- Leer `Insumos/` solo con `openpyxl` (lectura, sin modificar).
- Cruzar `Maestro × Evaluaciones` por `Identificacion + Programa` (texto, no número).
- Detectar y reportar: maestro sin notas (`304560321`), huérfano sin maestro (`999880777`), abandono con ceros (`111230984`), conteo dinámico de módulos (IA=4, Excel=3).
- Calcular Promedio y Asistencia según `AGENTS.md` y clasificar (Aprobación / Participación / Sin).

Resultado esperado (ya verificado a mano):
- 24 grupos en evaluaciones → 20 Aprobación, 3 Participación, 1 Sin.
- Emitibles reales: 22 (19 + 3), el huérfano no certifica.

Cómo validar:
- Mostrar tabla Id | Programa | n | Prom | Asis | Tipo + lista de los 3 casos borde.
- Usuario confirma números y tratamiento de huérfano/sin-datos antes de seguir.

## Etapa 2 — Backend Flask (Cocina)

Qué hacer:
- Crear `app.py` + `requirements.txt` (`Flask`, `openpyxl`).
- Toda la lógica en `app.py`: carga, agrupación, cálculo, decisión. Nada en HTML.
- Rutas mínimas: `/` (resumen/tabla), `/certificado/<id>/<programa>` (detalle). UTF-8 y tildes correctas.

Cómo validar:
- Arrancar con `python app.py`, abrir `/` y comprobar 22 emitibles y filtros por tipo.
- Revisar los 3 casos borde en pantalla antes de pasar a diseño.

## Etapa 3 — Interfaz (Salón)

Qué hacer:
- `templates/index.html`, `templates/certificado.html` (solo muestran, Jinja sin lógica).
- `static/style.css` azul marino y dorado, legible para imprimir.
- Comentarios Python en español marcando variables, tipos, condicionales, bucles y funciones.

Cómo validar:
- Revisión visual: colores, tildes, impresión de un certificado de cada tipo.
- Cierre: `Insumos/` intacto y `AGENTS.md` al día.
