import io
import os
from pathlib import Path

import openpyxl
import pytest

# LADRILLO: las credenciales de prueba se definen ANTES de importar la app.
# Así las pruebas no dependen del archivo .env de la máquina ni del modo demostración.
os.environ["ADMIN_CORREO"] = "admin@prueba.local"
os.environ["ADMIN_CLAVE"] = "clave-admin-prueba"
os.environ["USUARIO_CORREO"] = "usuario@prueba.local"
os.environ["USUARIO_CLAVE"] = "clave-usuario-prueba"
os.environ["SECRET_KEY"] = "clave-secreta-solo-para-pruebas-123456789"

# LADRILLO: se importa después de las variables de entorno a propósito.
from app import app  # noqa: E402

# LADRILLO: rutas de los archivos de fábrica (solo se leen, nunca se modifican).
BASE = Path(__file__).resolve().parent.parent
RUTA_MAESTRO = BASE / "Insumos" / "Maestro_Estudiantes.xlsx"
RUTA_EVALUACIONES = BASE / "Insumos" / "Registro_Evaluaciones.xlsx"


@pytest.fixture
def cliente():
    # LADRILLO: simula un navegador nuevo para cada prueba (sin sesión heredada).
    app.config["TESTING"] = True
    with app.test_client() as navegador:
        yield navegador


def construir_evaluaciones_con_notas_en_40():
    # LADRILLO: copia el Excel de evaluaciones en memoria y cambia SOLO las notas
    # de Camila Rojas Méndez (identificación 101230456) a 40.
    # El archivo original de Insumos/ no se toca.
    libro = openpyxl.load_workbook(RUTA_EVALUACIONES, data_only=True)
    hoja = libro.active
    filas_cambiadas = 0

    # LADRILLO: bucle por las filas, saltando el encabezado.
    for fila in hoja.iter_rows(min_row=2):
        identificacion = str(fila[0].value).strip()  # tipo texto
        if identificacion == "101230456":
            fila[3].value = 40  # columna D = Nota
            filas_cambiadas = filas_cambiadas + 1

    assert filas_cambiadas > 0, "No se encontraron notas de la identificación 101230456."

    # LADRILLO: BytesIO es la caja en memoria donde se guarda la copia modificada.
    buffer = io.BytesIO()
    libro.save(buffer)
    libro.close()
    return buffer.getvalue()


def iniciar_sesion(cliente, correo, clave):
    # LADRILLO: función auxiliar que simula el envío del formulario de entrada.
    return cliente.post("/login", data={"correo": correo, "clave": clave})


def test_sin_sesion_redirige_a_login(cliente):
    # Simula a alguien que abre la página principal sin haber entrado.
    respuesta = cliente.get("/")

    assert respuesta.status_code == 302, "Sin sesión, / debe redirigir."
    assert respuesta.headers["Location"].endswith("/login"), "Debe redirigir a /login."


def test_clave_incorrecta_no_entra(cliente):
    # Simula un intento de entrada con correo correcto y contraseña equivocada.
    respuesta = iniciar_sesion(cliente, "admin@prueba.local", "clave-equivocada")
    texto = respuesta.get_data(as_text=True)

    assert respuesta.status_code == 200, "Debe quedarse en la pantalla de entrada."
    assert "Correo o contraseña incorrectos" in texto, "Debe mostrar el mensaje de error."
    # LADRILLO: además comprobamos que la sesión no quedó abierta.
    assert cliente.get("/").status_code == 302, "No debe dejar entrar al tablero."


def test_usuario_no_puede_cargar(cliente):
    # Simula a una persona con rol "usuario" intentando usar la carga del admin.
    iniciar_sesion(cliente, "usuario@prueba.local", "clave-usuario-prueba")
    respuesta = cliente.post("/cargar", data={"accion": "ver"})

    assert respuesta.status_code == 403, "El rol usuario no puede entrar a /cargar."
    assert "solo para administradores" in respuesta.get_data(as_text=True)


def test_excel_modificado_cambia_contadores(cliente):
    # Simula a un administrador que sube una evaluación modificada de Camila
    # con notas de 40. Eso debe bajar sus notas y pasarla de Aprobación a Participación.
    iniciar_sesion(cliente, "admin@prueba.local", "clave-admin-prueba")

    evaluaciones_modificadas = construir_evaluaciones_con_notas_en_40()
    maestro = RUTA_MAESTRO.read_bytes()

    respuesta = cliente.post(
        "/cargar",
        data={
            "accion": "ver",
            "maestro": (io.BytesIO(maestro), "Maestro_Estudiantes.xlsx"),
            "evaluaciones": (io.BytesIO(evaluaciones_modificadas), "Registro_Evaluaciones.xlsx"),
        },
        content_type="multipart/form-data",
    )
    texto = respuesta.get_data(as_text=True)

    assert respuesta.status_code == 200, "La carga debe procesarse sin error."
    # LADRILLO: se leen los tres contadores tal como los dibuja el HTML.
    assert 'class="tarjeta verde"><strong>18</strong>' in texto, "Deben quedar 18 aprobaciones."
    assert 'class="tarjeta azul"><strong>4</strong>' in texto, "Deben quedar 4 participaciones."
    assert 'class="tarjeta roja"><strong>1</strong>' in texto, "Debe quedar 1 sin certificado."
