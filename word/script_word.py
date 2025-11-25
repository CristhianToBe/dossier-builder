import win32com.client as win32
import os
from common.common_office import load_json, parse_args
from common.mappers import apply_mappings

def main():
    base_docx, data_json, mapping_file, output_docx = parse_args(
        5, "Uso: python script_word.py <base_docx> <data_json> <mapping_json> <output_docx>"
    )
    # Asegurar que las rutas sean absolutas
    base_docx = os.path.abspath(base_docx)
    data_json = os.path.abspath(data_json)
    mapping_file = os.path.abspath(mapping_file)
    output_docx = os.path.abspath(output_docx)

    print(f"📄 Abriendo plantilla Word: {base_docx}")

    data = load_json(data_json)
    mapping_config = load_json(mapping_file)

    try:
        word = win32.Dispatch("Word.Application")
    except AttributeError:
        # Si falla la caché COM, usar Dispatch directamente
        from win32com.client import Dispatch
        word = Dispatch("Word.Application")
    word.Visible = False
    print(f"Intentando abrir plantilla: {base_docx}")
    doc = word.Documents.Open(base_docx)
    doc.TrackRevisions = False

    # Aplica mapeo desde JSON
    apply_mappings(doc, data, mapping_config)

    if os.path.exists(output_docx):
        os.remove(output_docx)

    doc.SaveAs(output_docx)
    doc.Close(SaveChanges=True)
    word.Quit()
    print(f"✅ Documento Word generado: {output_docx}")

if __name__ == "__main__":
    main()
