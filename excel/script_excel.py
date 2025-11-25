import shutil
import win32com.client as win32
from common import load_json, parse_args, apply_mappings


def main():
    base_excel, data_json, mapping_file, output_excel = parse_args(
        5, "Uso: python script_excel.py <base_excel> <data_json> <mapping_json> <output_excel>"
    )
    data = load_json(data_json)
    mapping_config = load_json(mapping_file)

    # Copiar plantilla a salida
    shutil.copy(base_excel, output_excel)

    excel = win32.DispatchEx("Excel.Application")
    print(">>> Excel version:", excel.Version)
    #excel.Visible = False
    wb = excel.Workbooks.Open(output_excel)
    sheet_name = mapping_config.get("sheet", 1)
    sheet = wb.Sheets(sheet_name)  # configurable si quieres

    # Aplica mapeo desde JSON (incluyendo reglas especiales como __TODAY__)
    apply_mappings(sheet, data, mapping_config)

    wb.Save()
    wb.Close(SaveChanges=True)
    print(f"✅ Excel actualizado: {output_excel}")

if __name__ == "__main__":
    main()
