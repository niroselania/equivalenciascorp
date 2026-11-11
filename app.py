import io
import os
import unicodedata
from collections import defaultdict
from copy import copy
from datetime import datetime

import openpyxl
import pandas as pd
from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "equivalencia-app-secret")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("DATA_DIR") or os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)
EQUIV_XLS = os.path.join(DATA_DIR, "equivalencia.xls")
EQUIV_XLSX = os.path.join(DATA_DIR, "equivalencia.xlsx")
EQUIV_META = os.path.join(DATA_DIR, "equivalencia.meta")
ALLOWED_EXTENSIONS = {".xls", ".xlsx"}
_CACHE = {"path": None, "mtime": None, "lookup": None}
SIZES = {
    "XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL", "XXXXL",
    "2XS", "3XS", "2XL", "3XL", "4XL", "OS", "OSFA", "ALL",
    "UNICO", "UNICA", "U", "SM", "MD", "LG",
} | {str(number) for number in range(55)}


def _extension(filename):
    return os.path.splitext(filename)[1].lower()


def _text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    value = str(value).strip()
    return "" if value.lower() in {"nan", "none", "nat", "<na>"} else value.removesuffix(".0")


def _header(value):
    value = unicodedata.normalize("NFKD", _text(value)).encode("ascii", "ignore").decode()
    return " ".join(value.upper().split())


def _find_header(headers, *names):
    return next((headers[name] for name in names if name in headers), None)


def _split_article(value):
    parts = _text(value).split(maxsplit=1)
    if len(parts) != 2 or not parts[0] or not parts[1] or " " in parts[1]:
        return "", ""
    return parts[0], parts[1]


def _looks_like_barcode(value):
    value = _text(value).replace(" ", "")
    return value.isdigit() and len(value) >= 8


def _row_values(row):
    return [_text(value) for value in row.tolist()]


def _find_article_in_row(values):
    """Encuentra el campo con formato SKU + código de color, ej. 21170 BLK."""
    for value in values:
        sku, color = _split_article(value)
        if sku.isdigit() and color:
            return sku, color
    return "", ""


def _find_size_in_row(values):
    # El talle de los exports Crystal aparece normalmente al final de la fila.
    return next((value.upper() for value in reversed(values) if value.upper() in SIZES), "")


def _read_equivalence(path):
    engine = "xlrd" if _extension(path) == ".xls" else "openpyxl"
    excel = pd.ExcelFile(path, engine=engine)
    lookup = defaultdict(list)
    for sheet in excel.sheet_names:
        df = excel.parse(sheet, header=None, dtype=str)
        if df.empty:
            continue
        # Los exports de Crystal Reports pueden tener encabezados sólo en la primera
        # hoja y reiniciar los datos en las demás. Por eso se detectan los campos
        # en cada fila, en vez de depender de columnas fijas.
        for _, row in df.iterrows():
            values = _row_values(row)
            code = next((value for value in values if _looks_like_barcode(value)), "")
            sku, color = _find_article_in_row(values)
            size = _find_size_in_row(values)
            if sku and color and size:
                lookup[(sku.upper(), color.upper(), size)].append(code)
    return lookup


def _get_lookup(path):
    global _CACHE
    mtime = os.path.getmtime(path)
    if _CACHE["path"] == path and _CACHE["mtime"] == mtime:
        return _CACHE["lookup"]
    _CACHE = {"path": path, "mtime": mtime, "lookup": _read_equivalence(path)}
    return _CACHE["lookup"]


def _deduplicate(codes):
    seen = {}
    for code in codes:
        normalized = str(int(code)) if code.isdigit() else code
        if normalized not in seen or len(code) < len(seen[normalized]):
            seen[normalized] = code
    return list(seen.values())


def _barcode(lookup, sku, color, size):
    sku, color, size = _text(sku), _text(color), _text(size)
    if not sku or not color or not size:
        return ""
    matches = _deduplicate(lookup.get((sku.upper(), color.upper(), size.upper()), []))
    return matches[0] if len(matches) == 1 else " / ".join(matches) if matches else "SIN COINCIDENCIA"


def _process_order(equivalence_path, uploaded_file):
    lookup = _get_lookup(equivalence_path)
    workbook = openpyxl.load_workbook(uploaded_file)
    processed = 0
    for sheet in workbook.worksheets:
        headers = {_header(sheet.cell(1, col).value): col for col in range(1, sheet.max_column + 1) if _text(sheet.cell(1, col).value)}
        sku_col = _find_header(headers, "SKU", "CODIGO", "CODIGO ARTICULO")
        color_col = _find_header(headers, "COD.COLOR", "COD COLOR", "CODIGO COLOR", "COLOR")
        size_col = _find_header(headers, "TALLE", "TAL", "SIZE")
        if not all((sku_col, color_col, size_col)):
            continue
        sheet.insert_cols(1)
        processed += 1
        source_header = sheet.cell(1, 2)
        target_header = sheet.cell(1, 1, "CODIGO DE BARRAS")
        target_header._style = copy(source_header._style)
        target_header.font = copy(source_header.font)
        target_header.fill = copy(source_header.fill)
        target_header.border = copy(source_header.border)
        target_header.alignment = copy(source_header.alignment)
        target_header.number_format = "@"
        sheet.column_dimensions["A"].width = 18
        for row in range(2, sheet.max_row + 1):
            result = _barcode(lookup, sheet.cell(row, sku_col + 1).value, sheet.cell(row, color_col + 1).value, sheet.cell(row, size_col + 1).value)
            sheet.cell(row, 1, result).number_format = "@"
    if not processed:
        raise ValueError("No encontré una hoja con las columnas SKU, COLOR (o COD.COLOR) y TALLE en la fila 1.")
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def _saved_equivalence():
    return EQUIV_XLS if os.path.exists(EQUIV_XLS) else EQUIV_XLSX if os.path.exists(EQUIV_XLSX) else None


def _equivalence_info():
    if not os.path.exists(EQUIV_META):
        return None
    with open(EQUIV_META, encoding="utf-8") as file:
        return file.read().strip()


@app.get("/")
def index():
    return render_template("index.html", equivalence_info=_equivalence_info())


@app.post("/subir-equivalencia")
def upload_equivalence():
    file = request.files.get("equivalencia")
    if not file or not file.filename:
        flash("Subí la planilla de equivalencia.")
        return redirect(url_for("index"))
    if _extension(file.filename) not in ALLOWED_EXTENSIONS:
        flash("La planilla de equivalencia debe ser .xls o .xlsx.")
        return redirect(url_for("index"))
    destination = EQUIV_XLS if _extension(file.filename) == ".xls" else EQUIV_XLSX
    other = EQUIV_XLSX if destination == EQUIV_XLS else EQUIV_XLS
    if os.path.exists(other):
        os.remove(other)
    file.save(destination)
    global _CACHE
    _CACHE = {"path": None, "mtime": None, "lookup": None}
    try:
        _get_lookup(destination)
    except Exception as error:
        flash(f"La planilla se guardó, pero no pude leerla: {error}")
        return redirect(url_for("index"))
    with open(EQUIV_META, "w", encoding="utf-8") as meta:
        meta.write(f"{secure_filename(file.filename)} - cargada {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    flash("Planilla de equivalencia cargada correctamente.")
    return redirect(url_for("index"))


@app.post("/procesar-pedido")
def process_order():
    file = request.files.get("pedido")
    if not file or not file.filename:
        flash("Subí la planilla de pedido.")
        return redirect(url_for("index"))
    if _extension(file.filename) != ".xlsx":
        flash("La planilla de pedido debe ser un archivo .xlsx.")
        return redirect(url_for("index"))
    equivalence_path = _saved_equivalence()
    if not equivalence_path:
        flash("Primero cargá la planilla de equivalencia.")
        return redirect(url_for("index"))
    data = io.BytesIO(file.read())
    try:
        output = _process_order(equivalence_path, data)
    except Exception as error:
        flash(f"No pude procesar la planilla: {error}")
        return redirect(url_for("index"))
    name = os.path.splitext(secure_filename(file.filename))[0]
    return send_file(output, as_attachment=True, download_name=f"{name} - con codigos de barra.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
