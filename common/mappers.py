from datetime import datetime
from win32com.client import constants
import unicodedata


def get_value_from_path(data: dict, path: str):
    """
    Navega por un dict siguiendo una ruta tipo 'A||B||C'.
    - Usa '||' como separador de niveles, para que claves con puntos no se rompan.
    - Soporta concatenaciones con '+' y literales entre comillas simples.
    """
    # Si es una concatenación de partes (ej. A + ' ' + B)
    if "+" in path:
        parts = [p.strip() for p in path.split("+")]
        values = []
        for part in parts:
            if part.startswith("'") and part.endswith("'"):
                values.append(part.strip("'"))  # literal
            else:
                values.append(str(get_value_from_path(data, part)))
        return "".join(values)

    # Ruta normal usando '||' como separador
    keys = path.split("||")
    val = data
    for k in keys:
        if isinstance(val, dict) and k in val:
            val = val[k]
        else:
            print(f"⚠️ Ruta no encontrada: {path} (faltó '{k}')")
            return ""
    return val

# Constantes mínimas para no depender de win32.constants
WD_FIND_CONTINUE = 1
WD_COLLAPSE_END = 0

def _replace_manual(rng, find_text, repl_text):
    """
    Busca find_text en el rango y lo reemplaza por repl_text (uno por uno).
    Devuelve cantidad de reemplazos.
    """
    f = rng.Find
    f.ClearFormatting()
    f.Text = find_text
    f.Forward = True
    f.Wrap = 1  # wdFindContinue
    count = 0

    while f.Execute():
        rng.Text = str(repl_text)
        count += 1
        rng.Collapse(0)  # wdCollapseEnd
        f = rng.Find
        f.Text = find_text
    return count

def apply_mappings(target, data: dict, config: dict):
    tipo = config["Tipo de documento"].lower()
    mappings = config["mapeo"]

    if tipo == "word":
        for placeholder, path in mappings.items():
            value = get_value_from_path(data, path) if not path.startswith("__") else handle_special(path)

            total = 0

            # 1) Cuerpo principal
            rng = target.Content.Duplicate
            total += _replace_manual(rng, placeholder, value)

            # 2) Tablas
            try:
                for tbl in target.Tables:
                    for row in tbl.Rows:
                        for cell in row.Cells:
                            rng_cell = cell.Range.Duplicate
                            total += _replace_manual(rng_cell, placeholder, value)
            except Exception:
                pass

            print(f"→ {placeholder}: reemplazos hechos = {total}")

    elif tipo == "excel":
        for cell, path in mappings.items():
            value = get_value_from_path(data, path) if not path.startswith("__") else handle_special(path)
            rng = target.Range(cell)

            if path == "__TODAY__":
                # Para fechas, usamos formato dd/mm/yyyy
                rng.NumberFormat = "dd/mm/yyyy"
                rng.Value = value
            else:
                # Para todo lo demás, forzamos texto
                rng.NumberFormat = "@"
                rng.Value = str(value)

            print(f"Escrito {value} en {cell}")

def handle_special(path: str):
    """Maneja valores especiales en el JSON de mapeos."""
    if path == "__TODAY__":
        return datetime.today()
    return None
