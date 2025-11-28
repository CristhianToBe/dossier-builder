from pathlib import Path
import sys

from django.conf import settings
from django.http import FileResponse, HttpResponse
from django.shortcuts import render, redirect
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile

# 👉 AÑADIMOS LA RAÍZ DEL REPO AL PYTHONPATH
# BASE_DIR = carpeta "webui" (donde está manage.py)
WEBUI_BASE = Path(settings.BASE_DIR)          # ...\dossier-builder\webui
REPO_ROOT = WEBUI_BASE.parent                 # ...\dossier-builder

if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# Ahora ya podemos importar tu script
from word.script_word import run_word
from excel.script_excel import run_excel

def save_uploaded_file(file):
    path = default_storage.save(file.name, ContentFile(file.read()))
    return default_storage.path(path)

def resolve_path(path_str: str, base: Path | None = None) -> Path:
    """
    Si la ruta es absoluta, la retorna tal cual.
    Si es relativa, la une a base (por defecto REPO_ROOT).
    """
    p = Path(path_str)
    if p.is_absolute():
        return p
    return (base or REPO_ROOT) / p

def index(request):
    plantillas_dir = REPO_ROOT / "plantillas"

    # Buscar plantillas disponibles
    word_files = sorted(plantillas_dir.glob("*.docx"))
    excel_files = sorted(plantillas_dir.glob("*.xlsx"))

    # Defaults actuales (los que antes estaban hardcodeados)
    word_base_default = REPO_ROOT / "plantillas" / "1839 - Informe parcial .docx"
    word_json_default = Path(
        r"C:\Users\ctorresb\OneDrive - Direccion de Impuestos y Aduanas Nacionales de Colombia\CASOS\ACTIVOS\ITAU COLOMBIA S A\ITAU COLOMBIA S A 2020\R - 2020\Auditoria\data.json"
    )
    word_mapping_default = REPO_ROOT / "configs" / "1839_mapping.json"
    word_out_default = Path(
        r"C:\Users\ctorresb\OneDrive - Direccion de Impuestos y Aduanas Nacionales de Colombia\CASOS\ACTIVOS\ITAU COLOMBIA S A\ITAU COLOMBIA S A 2020\R - 2020\Auditoria\F1839 - Informe parcial.docx"
    )

    excel_base_default = REPO_ROOT / "plantillas" / "1811 - VERIFICACION REQUISITOS FORMALES.xlsx"
    excel_json_default = Path(
        r"C:\Users\ctorresb\OneDrive - Direccion de Impuestos y Aduanas Nacionales de Colombia\CASOS\ACTIVOS\BANCO SERFINANZA S.A\R - 2024\Auditoria\data.json"
    )
    excel_mapping_default = REPO_ROOT / "configs" / "1811_mapping.json"
    excel_out_default = Path(
        r"C:\Users\ctorresb\OneDrive - Direccion de Impuestos y Aduanas Nacionales de Colombia\CASOS\ACTIVOS\BANCO SERFINANZA S.A\R - 2024\Auditoria\F1811.xlsx"
    )

    context = {
        # defaults como strings
        "word_base_default": str(word_base_default.relative_to(REPO_ROOT)),
        "word_json_default": str(word_json_default),
        "word_mapping_default": str(word_mapping_default),
        "word_out_default": str(word_out_default),

        "excel_base_default": str(excel_base_default.relative_to(REPO_ROOT)),
        "excel_json_default": str(excel_json_default),
        "excel_mapping_default": str(excel_mapping_default),
        "excel_out_default": str(excel_out_default),

        # listas de plantillas para los selects
        "word_templates": [
            {
                "value": str(p.relative_to(REPO_ROOT)),  # ej: "plantillas/1839 - Informe parcial .docx"
                "name": p.name,                          # solo el nombre de archivo
            }
            for p in word_files
        ],
        "excel_templates": [
            {
                "value": str(p.relative_to(REPO_ROOT)),
                "name": p.name,
            }
            for p in excel_files
        ],
    }
    return render(request, "builder/index.html", context)


def run_word_view(request):
    if request.method != "POST":
        return redirect("index")

    base_docx = resolve_path(request.POST.get("word_base"), REPO_ROOT)
    mapping_file = resolve_path(request.POST.get("word_mapping"), REPO_ROOT)
    output_docx = resolve_path(request.POST.get("word_out"), None)

    # ✅ JSON SUBIDO DESDE EXPLORADOR
    json_file = request.FILES.get("word_json")
    if not json_file:
        return HttpResponse("No se subió ningún archivo JSON.")

    data_json_path = save_uploaded_file(json_file)

    run_word(
        str(base_docx),
        data_json_path,
        str(mapping_file),
        str(output_docx),
    )

    if Path(output_docx).exists():
        return FileResponse(
            open(output_docx, "rb"),
            as_attachment=True,
            filename=Path(output_docx).name,
        )

    return HttpResponse("Se ejecutó la generación, pero no se encontró el archivo de salida.")


def run_excel_view(request):
    if request.method != "POST":
        return redirect("index")

    base_excel = resolve_path(request.POST.get("excel_base"), REPO_ROOT)
    mapping_file = resolve_path(request.POST.get("excel_mapping"), REPO_ROOT)
    output_excel = resolve_path(request.POST.get("excel_out"), None)

    # ✅ JSON SUBIDO DESDE EXPLORADOR
    json_file = request.FILES.get("excel_json")
    if not json_file:
        return HttpResponse("No se subió ningún archivo JSON.")

    data_json_path = save_uploaded_file(json_file)

    run_excel(
        str(base_excel),
        data_json_path,
        str(mapping_file),
        str(output_excel),
    )

    if Path(output_excel).exists():
        return FileResponse(
            open(output_excel, "rb"),
            as_attachment=True,
            filename=Path(output_excel).name,
        )

    return HttpResponse("Se ejecutó la generación, pero no se encontró el archivo de salida.")

