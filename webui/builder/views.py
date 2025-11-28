from pathlib import Path
import sys

from django.conf import settings
from django.http import FileResponse, HttpResponse
from django.shortcuts import render, redirect

# 👉 AÑADIMOS LA RAÍZ DEL REPO AL PYTHONPATH
# BASE_DIR = carpeta "webui" (donde está manage.py)
WEBUI_BASE = Path(settings.BASE_DIR)          # ...\dossier-builder\webui
REPO_ROOT = WEBUI_BASE.parent                 # ...\dossier-builder

if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# Ahora ya podemos importar tu script
from word.script_word import run_word
from excel.script_excel import run_excel


def index(request):
    return render(request, "builder/index.html")


def run_word_view(request):
    if request.method != "POST":
        return redirect("index")

    # === AQUÍ VAN LAS VARIABLES QUE ANTES ESTABAN EN EL .BAT ===

    # Ruta a la plantilla base (ajusta si tu carpeta real es distinta)
    base_docx = REPO_ROOT / "plantillas" / "1839 - Informe parcial .docx"

    # JSON y salida: los mismos que tenías en el .bat
    data_json = Path(
        r"C:\Users\ctorresb\OneDrive - Direccion de Impuestos y Aduanas Nacionales de Colombia\CASOS\ACTIVOS\ITAU COLOMBIA S A\ITAU COLOMBIA S A 2020\R - 2020\Auditoria\data.json"
    )
    mapping_file = REPO_ROOT / "configs" / "1839_mapping.json"
    output_docx = Path(
        r"C:\Users\ctorresb\OneDrive - Direccion de Impuestos y Aduanas Nacionales de Colombia\CASOS\ACTIVOS\ITAU COLOMBIA S A\ITAU COLOMBIA S A 2020\R - 2020\F1839 - Informe parcial.docx"
    )

    # Ejecutar tu flujo de generación
    run_word(
        str(base_docx),
        str(data_json),
        str(mapping_file),
        str(output_docx),
    )

    # Si el archivo existe, lo devolvemos como descarga
    if output_docx.exists():
        return FileResponse(
            open(output_docx, "rb"),
            as_attachment=True,
            filename=output_docx.name,
        )

    return HttpResponse(
        "Se ejecutó la generación, pero no se encontró el archivo de salida."
    )


def run_excel_view(request):
    if request.method != "POST":
        return redirect("index")

    # === RUTAS TOMADAS DIRECTAMENTE DE TU .BAT ===

    # Plantilla base (relativa a la raíz del repo)
    base_excel = REPO_ROOT / "Plantillas" / "1811 - VERIFICACION REQUISITOS FORMALES.xlsx"

    # JSON de datos (ruta absoluta)
    data_json = Path(
        r"C:\Users\ctorresb\OneDrive - Direccion de Impuestos y Aduanas Nacionales de Colombia\CASOS\ACTIVOS\BANCO SERFINANZA S.A\R - 2024\Auditoria\data.json"
    )

    # Mapping (relativo al repo)
    mapping_file = REPO_ROOT / "configs" / "1811_mapping.json"

    # Archivo de salida (ruta absoluta)
    output_excel = Path(
        r"C:\Users\ctorresb\OneDrive - Direccion de Impuestos y Aduanas Nacionales de Colombia\CASOS\ACTIVOS\BANCO SERFINANZA S.A\R - 2024\F1811.xlsx"
    )

    # === EJECUCIÓN DIRECTA DEL SCRIPT (SIN .BAT) ===
    run_excel(
        str(base_excel),
        str(data_json),
        str(mapping_file),
        str(output_excel),
    )

    # === DEVOLVER ARCHIVO COMO DESCARGA ===
    if output_excel.exists():
        return FileResponse(
            open(output_excel, "rb"),
            as_attachment=True,
            filename=output_excel.name,
        )

    return HttpResponse(
        "Se ejecutó la generación, pero no se encontró el archivo de salida."
    )