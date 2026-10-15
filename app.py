# app.py — LA COCINA (toda la lógica vive aquí, NUNCA en el HTML)
# Academia Horizonte: lee los Excel de Insumos/, cruza, calcula y decide.
from pathlib import Path  # LADRILLO: rutas ancladas al propio archivo (tipo Path)
import openpyxl  # LADRILLO: librería para leer Excel .xlsx
from flask import Flask, render_template, request  # LADRILLO: servidor web + páginas

# LADRILLO: variables (cajas con nombre). Tipo Path y texto (str).
# Path(__file__) = ubicación de ESTE archivo, funciona sin importar desde dónde se arranque.
BASE = Path(__file__).resolve().parent
RUTA_MAESTRO = BASE / "Insumos" / "Maestro_Estudiantes.xlsx"
RUTA_EVAL = BASE / "Insumos" / "Registro_Evaluaciones.xlsx"

app = Flask(__name__)  # LADRILLO: variable que guarda la aplicación web


# LADRILLO: función (bloque reutilizable: recibe datos, devuelve un resultado)
def calcular_promedio(valores):
    # LADRILLO: variables y tipos. float = número con decimales, int = entero.
    total = 0.0  # tipo float
    cantidad = 0  # tipo int
    # LADRILLO: bucle (repetir una acción por cada elemento de la lista)
    for v in valores:  # v es tipo float
        total = total + float(v)
        cantidad = cantidad + 1
    # LADRILLO: condicional (pregunta sí/no para decidir)
    if cantidad == 0:
        return 0.0
    return total / cantidad


# LADRILLO: función que decide el certificado según AGENTS.md (límites inclusivos)
def decidir_certificado(promedio, asistencia):
    # LADRILLO: condicionales encadenados (se revisan en orden)
    if asistencia < 80:  # tipo: comparación de float; menos de 80 = sin certificado
        return "Sin certificado"
    if promedio >= 70:  # tipo: comparación; 70 SÍ aprueba
        return "Aprobación"
    return "Participación"  # si no fue Sin ni Aprobación, es Participación


def cargar_datos():
    # Lee la Bodega (solo lectura, jamás modifica Insumos/) y deja todo listo para el Salón.
    # LADRILLO: variables tipo lista (varios valores) y dict (tabla clave -> valor).
    wb_m = openpyxl.load_workbook(RUTA_MAESTRO, data_only=True)
    ws_m = wb_m.active
    maestro = {}  # tipo dict: clave (id, programa) -> datos del estudiante
    # LADRILLO: bucle por filas (saltamos el encabezado con [1:])
    for fila in list(ws_m.iter_rows(values_only=True))[1:]:
        if fila[0] is None:
            continue
        ident = str(fila[0]).strip()  # tipo texto: la ID siempre como texto
        maestro[(ident, str(fila[3]).strip())] = {
            "id": ident,  # tipo texto
            "nombre": str(fila[1]).strip(),  # tipo texto
            "correo": str(fila[2]).strip(),  # tipo texto
            "programa": str(fila[3]).strip(),  # tipo texto
            "cohorte": str(fila[4]).strip(),  # tipo texto
        }
    wb_m.close()

    wb_e = openpyxl.load_workbook(RUTA_EVAL, data_only=True)
    ws_e = wb_e.active
    notas_por_clave = {}  # tipo dict: clave -> lista de notas (float)
    asis_por_clave = {}  # tipo dict: clave -> lista de asistencias (float)
    # LADRILLO: bucle para agrupar por Identificacion + Programa
    for fila in list(ws_e.iter_rows(values_only=True))[1:]:
        if fila[0] is None:
            continue
        clave = (str(fila[0]).strip(), str(fila[1]).strip())  # tipo tupla de textos
        # LADRILLO: condicional (si la clave es nueva, creo sus listas)
        if clave not in notas_por_clave:
            notas_por_clave[clave] = []
            asis_por_clave[clave] = []
        notas_por_clave[clave].append(float(fila[3]))  # tipo float
        asis_por_clave[clave].append(float(fila[4]))  # tipo float
    wb_e.close()

    # LADRILLO: bucle para la tabla principal (cruce interno: solo maestro CON evaluación)
    resultados = []  # tipo lista de dict
    for clave in notas_por_clave:
        # LADRILLO: condicional (si no está en el maestro, es huérfano: va a inconsistencias, no a la tabla)
        if clave not in maestro:
            continue
        promedio = calcular_promedio(notas_por_clave[clave])  # LADRILLO: llamada a función
        asistencia = calcular_promedio(asis_por_clave[clave])  # reutilizamos el promedio para asistencia
        tipo = decidir_certificado(promedio, asistencia)  # LADRILLO: llamada a función
        dato = maestro[clave]
        # LADRILLO: condicional para el color de la fila (verde / azul / rojo)
        if tipo == "Aprobación":
            clase = "aprobacion"
        elif tipo == "Participación":
            clase = "participacion"
        else:
            clase = "sin"
        resultados.append({
            "id": dato["id"],
            "nombre": dato["nombre"],
            "programa": dato["programa"],
            "cohorte": dato["cohorte"],
            "modulos": len(notas_por_clave[clave]),  # tipo int
            "promedio": round(promedio, 2),  # tipo float
            "asistencia": round(asistencia, 2),  # tipo float
            "tipo": tipo,  # tipo texto
            "clase": clase,  # tipo texto (para el color)
        })
    # Orden por programa y nombre (sin lógica en el HTML)
    resultados.sort(key=lambda r: (r["programa"], r["nombre"]))

    # Contadores para el encabezado (se calculan aquí, el HTML solo los muestra)
    contadores = {"Aprobación": 0, "Participación": 0, "Sin certificado": 0}  # tipo dict
    # LADRILLO: bucle + condicionales para contar
    for r in resultados:
        if r["tipo"] == "Aprobación":
            contadores["Aprobación"] = contadores["Aprobación"] + 1
        elif r["tipo"] == "Participación":
            contadores["Participación"] = contadores["Participación"] + 1
        else:
            contadores["Sin certificado"] = contadores["Sin certificado"] + 1

    programas = sorted({m["programa"] for m in maestro.values()})  # tipo lista de textos

    # Las 3 inconsistencias detectadas en la validación
    prom_h = round(calcular_promedio(notas_por_clave.get(("999880777", "Técnico en IA Aplicada"), [0])), 2)
    asis_h = round(calcular_promedio(asis_por_clave.get(("999880777", "Técnico en IA Aplicada"), [0])), 2)
    inconsistencias = [
        {"titulo": "Maestro SIN evaluación (1)",
         "detalle": "Espinoza Leon Javier | 304560321 | Técnico en IA Aplicada: está en el maestro pero no tiene ninguna fila de notas."},
        {"titulo": "Evaluación SIN maestro (1)",
         "detalle": f"ID 999880777 | Técnico en IA Aplicada: tiene 4 módulos (prom {prom_h}, asis {asis_h}) pero no existe en el maestro, no se puede certificar."},
        {"titulo": "Posible abandono con ceros (1)",
         "detalle": "Zuñiga Arias Fernanda | 111230984: notas 88, 90, 0, 0 y asistencias 70, 55, 20, 0 (prom 44.5, asis 36.25). Los ceros cuentan en el promedio."},
    ]
    return resultados, contadores, programas, inconsistencias


# LADRILLO: ruta web (qué página se muestra al visitar una dirección)
@app.route("/")
def index():
    resultados, contadores, programas, inconsistencias = cargar_datos()
    # LADRILLO: variable + condicional para el filtro por programa (llega como ?programa=...)
    programa_actual = request.args.get("programa", "Todos")  # tipo texto
    if programa_actual == "Todos":
        filtrados = resultados
    else:
        filtrados = [r for r in resultados if r["programa"] == programa_actual]
    # El Salón solo muestra: no calcula nada aquí, todo viene preparado.
    return render_template("index.html", resultados=filtrados, contadores=contadores,
                           programas=programas, programa_actual=programa_actual,
                           inconsistencias=inconsistencias)


@app.route("/certificado/<ident>/<path:programa>")
def certificado(ident, programa):
    resultados, _, _, _ = cargar_datos()
    # LADRILLO: bucle + condicional para buscar un certificado
    elegido = None
    for r in resultados:
        if r["id"] == ident and r["programa"] == programa:
            elegido = r
            break
    if elegido is None:
        return "Certificado no encontrado", 404
    return render_template("certificado.html", r=elegido)


if __name__ == "__main__":
    # Arranca el Salón en http://localhost:5000 ( codificación UTF-8 para tildes )
    app.run(host="127.0.0.1", port=5000, debug=False)
