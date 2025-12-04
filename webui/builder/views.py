from pathlib import Path
import sys
import os
import io
import zipfile
import uuid
import json


from django.conf import settings
from django.http import FileResponse, HttpResponse
from django.shortcuts import render, redirect
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
from django.conf import settings
from pathlib import Path

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

def save_json_from_text(json_text: str, prefix: str) -> str:
    """
    Valida el JSON y lo guarda en un archivo temporal dentro de MEDIA_ROOT.
    Devuelve la ruta absoluta del archivo.
    """
    # Validar JSON
    try:
        parsed = json.loads(json_text)
    except json.JSONDecodeError as e:
        raise ValueError(f"JSON inválido: {e}")

    media_root = Path(settings.MEDIA_ROOT)
    media_root.mkdir(parents=True, exist_ok=True)

    filename = f"{prefix}_{uuid.uuid4().hex}.json"
    path = media_root / filename

    with open(path, "w", encoding="utf-8") as f:
        json.dump(parsed, f, ensure_ascii=False, indent=2)

    return str(path)


def index(request):
    plantillas_dir = REPO_ROOT / "Plantillas"
    configs_dir = REPO_ROOT / "configs"

    word_files = sorted(plantillas_dir.glob("*.docx"))
    excel_files = sorted(plantillas_dir.glob("*.xlsx"))
    mapping_files = sorted(configs_dir.glob("*.json"))

    word_base_default = REPO_ROOT / "Plantillas" / "1839 - Informe parcial .docx"
    word_mapping_default = REPO_ROOT / "configs" / "1839_mapping.json"

    excel_base_default = REPO_ROOT / "Plantillas" / "1811 - VERIFICACION REQUISITOS FORMALES.xlsx"
    excel_mapping_default = REPO_ROOT / "configs" / "1811_mapping.json"

    context = {
        "word_base_default": str(word_base_default.relative_to(REPO_ROOT)),
        # 👇 ahora guardamos el default como ruta relativa (para que matchee con el value del <option>)
        "word_mapping_default": str(word_mapping_default.relative_to(REPO_ROOT)),
        "word_out_name_default": "F1839 - Informe parcial.docx",

        "excel_base_default": str(excel_base_default.relative_to(REPO_ROOT)),
        "excel_mapping_default": str(excel_mapping_default.relative_to(REPO_ROOT)),
        "excel_out_name_default": "F1811.xlsx",

        "word_templates": [
            {"value": str(p.relative_to(REPO_ROOT)), "name": p.name}
            for p in word_files
        ],
        "excel_templates": [
            {"value": str(p.relative_to(REPO_ROOT)), "name": p.name}
            for p in excel_files
        ],
        # 👇 lista única de mappings para ambos (Word/Excel)
        "mapping_templates": [
            {"value": str(p.relative_to(REPO_ROOT)), "name": p.name}
            for p in mapping_files
        ],

        "word_json_default": "",
        "excel_json_default": "",
    }
    return render(request, "builder/index.html", context)


def run_word_view(request):
    if request.method != "POST":
        return redirect("index")

    base_docx = resolve_path(request.POST.get("word_base"), REPO_ROOT)
    mapping_file = resolve_path(request.POST.get("word_mapping"), REPO_ROOT)

    out_name = request.POST.get("word_out_name", "").strip()
    if not out_name:
        return HttpResponse("Debes indicar el nombre del archivo Word de salida.", status=400)

    # Carpeta interna temporal para Word
    outputs_dir = Path(settings.MEDIA_ROOT) / "word_outputs"
    outputs_dir.mkdir(parents=True, exist_ok=True)
    output_docx = outputs_dir / out_name

    json_text = request.POST.get("word_json_text", "").strip()
    if not json_text:
        return HttpResponse("No se recibió contenido JSON para Word.", status=400)

    try:
        data_json_path = save_json_from_text(json_text, "word")
    except ValueError as e:
        return HttpResponse(str(e), status=400)

    # Generar el Word
    run_word(
        str(base_docx),
        data_json_path,
        str(mapping_file),
        str(output_docx),
    )

    if not output_docx.exists():
        return HttpResponse("Se ejecutó la generación, pero no se encontró el archivo de salida.", status=500)

    # 📦 Crear ZIP en memoria con el DOCX y el JSON
    zip_buffer = io.BytesIO()
    zip_name = f"{output_docx.stem}_paquete.zip"

    try:
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            # Documento generado
            zf.write(output_docx, arcname=output_docx.name)
            # JSON usado para generarlo
            zf.write(data_json_path, arcname="data.json")

        zip_buffer.seek(0)
    finally:
        # 🧹 Limpiar temporales
        try:
            if output_docx.exists():
                os.remove(output_docx)
        except OSError:
            pass

        try:
            tmp_json = Path(data_json_path)
            if tmp_json.exists():
                os.remove(tmp_json)
        except OSError:
            pass

    # Devolver ZIP como descarga
    return FileResponse(
        zip_buffer,
        as_attachment=True,
        filename=zip_name,
    )


def run_excel_view(request):
    if request.method != "POST":
        return redirect("index")

    base_excel = resolve_path(request.POST.get("excel_base"), REPO_ROOT)
    mapping_file = resolve_path(request.POST.get("excel_mapping"), REPO_ROOT)

    out_name = request.POST.get("excel_out_name", "").strip()
    if not out_name:
        return HttpResponse("Debes indicar el nombre del archivo Excel de salida.", status=400)

    outputs_dir = Path(settings.MEDIA_ROOT) / "excel_outputs"
    outputs_dir.mkdir(parents=True, exist_ok=True)
    output_excel = outputs_dir / out_name

    json_text = request.POST.get("excel_json_text", "").strip()
    if not json_text:
        return HttpResponse("No se recibió contenido JSON para Excel.", status=400)

    try:
        data_json_path = save_json_from_text(json_text, "excel")
    except ValueError as e:
        return HttpResponse(str(e), status=400)

    # Generar el Excel
    run_excel(
        str(base_excel),
        data_json_path,
        str(mapping_file),
        str(output_excel),
    )

    if not output_excel.exists():
        return HttpResponse("Se ejecutó la generación, pero no se encontró el archivo de salida.", status=500)

    # 📦 Crear ZIP en memoria con el XLSX y el JSON
    zip_buffer = io.BytesIO()
    zip_name = f"{output_excel.stem}_paquete.zip"

    try:
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(output_excel, arcname=output_excel.name)
            zf.write(data_json_path, arcname="data.json")

        zip_buffer.seek(0)
    finally:
        # 🧹 Limpiar temporales
        try:
            if output_excel.exists():
                os.remove(output_excel)
        except OSError:
            pass

        try:
            tmp_json = Path(data_json_path)
            if tmp_json.exists():
                os.remove(tmp_json)
        except OSError:
            pass

    return FileResponse(
        zip_buffer,
        as_attachment=True,
        filename=zip_name,
    )
