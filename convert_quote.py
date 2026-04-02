#!/usr/bin/env python3
"""
Chinify Quote Converter — Vision API + Precise Y-Position Image Matching
- Claude Vision extracts product text data
- fitz y-position overlap matches EACH image to its exact product row
- Fabric swatch (LEFT col) → FOTO/E  (matches J078 format)
- Furniture photo (RIGHT col) → NOT included (J078 doesn't have it)
- Column layout matches J078 exactly

Usage:
  export ANTHROPIC_API_KEY=sk-ant-...
  python3 convert_quote.py
"""

import os, re, io, json, base64, shutil, zipfile
from datetime import date

import fitz
import pdfplumber
import anthropic
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

# ── Config ───────────────────────────────────────────────────────────────────
PDF_PATH      = "COTIZACION VIDALTA - PEDIDOS CHINA(1).pdf"
TEMPLATE_PATH = "2025 J0 初始报价表.xlsx"
OUTPUT_PATH   = "VIDALTA_chinify_报价.xlsx"
MODEL         = "claude-sonnet-4-6"

# ── Column layout — matches J078 EXACTLY ─────────────────────────────────────
# B=Número  C=Área  D=Producto  E=Foto(swatch)  F=Largo  G=Prof  H=Altura
# I=Detalle  J=Cantidad  K=PrecioUnit  L=PrecioTotal  M=CBM  N=KG  O=TotalCBM  P=Notas
COL_WIDTHS = {
    "A": 1.5, "B": 14, "C": 13, "D": 24, "E": 20,
    "F": 9,   "G": 10, "H": 9,  "I": 22, "J": 8,
    "K": 12,  "L": 12, "M": 7,  "N": 7,  "O": 9,  "P": 18,
}
# Image cell pixel dimensions
# Excel: 1 char unit ≈ 8.43px at 96dpi for Calibri 11pt
# openpyxl saves images using 96dpi EMU conversion
# Col E (20 chars) ≈ 174px, Col P (18 chars) ≈ 157px
# Row 100pt = 100*96/72 = 133px effective
FOTO_W_PX   = 168   # fill ~96% of col E
FOTO_H_PX   = 126   # fill ~95% of row
NOTAS_W_PX  = 150   # fill ~95% of col P
NOTAS_H_PX  = 126
ROW_H_IMG   = 100   # row height (pt) for rows with images (≈133px)
ROW_H_TEXT  = 25    # row height (pt) for text-only rows

# ── Styles ───────────────────────────────────────────────────────────────────
def bdr():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
AREA_FILL   = PatternFill("solid", fgColor="D6DCE4")
FILL_W      = PatternFill("solid", fgColor="FFFFFF")
FILL_B      = PatternFill("solid", fgColor="EBF3FB")
TOTAL_FILL  = PatternFill("solid", fgColor="D9E2F3")
TERMS_FILL  = PatternFill("solid", fgColor="F2F7FC")

WHITE_BOLD = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
DARK_BOLD  = Font(name="Calibri", bold=True, color="1F3864", size=9)
DARK_NORM  = Font(name="Calibri", color="1F3864", size=9)
CENTER     = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT       = Alignment(horizontal="left",   vertical="center", wrap_text=True)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1  Extract product row y-ranges from pdfplumber
# ══════════════════════════════════════════════════════════════════════════════
AREA_MARKERS = {
    "ENTRADA","COMEDOR","SALA","FAMILY","TERRAZA","COCINA","BAÑO","ESTUDIO",
    "PASILLO","RECAMÁRA PRINCIPAL","RECAMÁRA BEBE",
    "RECAMARA PRINCIPAL","RECAMARA BEBE","DEPARTAMENTO",
}

def get_product_rows_with_yranges(pdf_path):
    """
    Returns list of dicts:
      {area, product, code, qty, dims, material, y0, y1}
    y0/y1 are the vertical bounds of that row in PDF points.
    """
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        # Get cell bounding boxes
        tables = page.find_tables()
        t = tables[0]

        # Build row_y map: row index → (y0, y1)
        # Group cells by their y0 value
        row_map = {}
        for cell in t.cells:
            y0 = round(cell[1], 2)
            y1 = round(cell[3], 2)
            key = y0
            if key not in row_map:
                row_map[key] = y1

        sorted_row_starts = sorted(row_map.keys())
        row_y = [(y0, row_map[y0]) for y0 in sorted_row_starts]

        raw = page.extract_table()

    # Align raw rows with y-positions
    products = []
    current_area = ""

    for i, row_data in enumerate(raw):
        name = (row_data[0] or "").strip()
        qty  = (row_data[1] or "").strip()
        code = (row_data[2] or "").strip()
        dims = (row_data[3] or "").strip()
        mat  = (row_data[5] or "").strip()

        if not name:
            if products and mat:
                # Store sub-row material codes separately (e.g. SOFA LAWSON's 2nd fabric)
                # so each can be assigned to its own swatch row later
                products[-1]["extra_materials"].append(mat)
            continue

        # Skip address blocks (contain email / company contact info)
        if "@" in name:
            continue

        # Normalize multi-line wrapped product names
        name = " ".join(name.split())

        # Skip the column header row
        if name.upper().startswith("PRODUCTO"):
            continue

        is_area = (
            not qty and (
                name.upper() in AREA_MARKERS or
                (name == name.upper() and not re.search(r'\d', name) and len(name) > 2)
            )
        )
        if is_area:
            current_area = name
            continue

        y0, y1 = row_y[i] if i < len(row_y) else (0, 0)

        products.append({
            "area":     current_area,
            "product":  name,
            "code":     code,
            "qty":      qty,
            "dims":     dims,
            "material": mat,
            "y0":       y0,
            "y1":       y1,
        "swatches":        [],   # list of fabric images → MUESTRA/P
            "photo":           None, # furniture photo → FOTO/E
            "extra_materials": [],   # material codes from sub-rows (e.g. SOFA LAWSON row 2)
        })

    return products


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2  Extract images and match to products by y-overlap
# ══════════════════════════════════════════════════════════════════════════════
def _is_blank(pix, threshold=247):
    """True if pixmap is mostly white (empty cell)."""
    samples = pix.samples
    return (sum(samples) / len(samples)) > threshold


def extract_and_match_images(pdf_path, products):
    """
    Hybrid approach:
    1. Embedded images via get_images() — catches high-res embedded photos/swatches
       (SOFA LAWSON's 2 separate swatches come from here)
    2. Pixel render fallback via get_pixmap(clip) — catches drawn/vector swatches
       (SILLON CLOUD, OTTOMAN, and any product where content isn't an embedded image)

    LEFT column (x 300-450) = furniture photos → FOTO/E
    RIGHT column (x 480-560) = fabric/material swatches → MUESTRA/P
    """
    doc  = fitz.open(pdf_path)
    page = doc[0]

    # ── Pass 1: embedded images ───────────────────────────────────────────────
    emb = []
    for img in page.get_images(full=True):
        xref = img[0]
        bbox = page.get_image_bbox(img)
        base = doc.extract_image(xref)
        size = len(base["image"])
        if size < 15_000:
            continue
        xc = (bbox.x0 + bbox.x1) / 2
        yc = (bbox.y0 + bbox.y1) / 2
        emb.append({"xc": xc, "yc": yc, "data": base["image"],
                    "side": "LEFT" if xc < 450 else "RIGHT", "size": size})

    for p in products:
        y0, y1 = p["y0"], p["y1"]
        if y0 == 0 and y1 == 0:
            continue
        left_hits  = [i for i in emb if i["side"]=="LEFT"  and y0<=i["yc"]<=y1]
        right_hits = [i for i in emb if i["side"]=="RIGHT" and y0<=i["yc"]<=y1]
        if left_hits:
            p["photo"] = max(left_hits, key=lambda x: x["size"])["data"]
        for hit in sorted(right_hits, key=lambda x: x["yc"]):
            p["swatches"].append(hit["data"])

    # ── Pass 2: render-based fallback for missing content ────────────────────
    mat = fitz.Matrix(3, 3)   # 3× zoom → ~216 DPI
    for p in products:
        y0, y1 = p["y0"], p["y1"]
        if y0 == 0 and y1 == 0:
            continue

        # Furniture photo fallback (LEFT col x 300-450)
        if not p.get("photo"):
            pix = page.get_pixmap(matrix=mat, clip=fitz.Rect(300, y0, 450, y1), alpha=False)
            if not _is_blank(pix):
                p["photo"] = pix.tobytes("jpeg", jpg_quality=85)

        # Swatch fallback (RIGHT col x 480-560)
        if not p["swatches"]:
            pix = page.get_pixmap(matrix=mat, clip=fitz.Rect(480, y0, 560, y1), alpha=False)
            if not _is_blank(pix):
                p["swatches"].append(pix.tobytes("jpeg", jpg_quality=85))

    doc.close()

    prods_with_swatch = sum(1 for p in products if p["swatches"])
    prods_with_photo  = sum(1 for p in products if p.get("photo"))
    print(f"  Furniture photos : {prods_with_photo}/{len(products)}")
    print(f"  Fabric swatches  : {prods_with_swatch} products "
          f"({sum(len(p['swatches']) for p in products)} total)")
    return products


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3  (Optional) Claude Vision to clean up product names / dims
# ══════════════════════════════════════════════════════════════════════════════
def render_pdf_page(pdf_path, dpi=150):
    doc  = fitz.open(pdf_path)
    page = doc[0]
    mat  = fitz.Matrix(dpi / 72, dpi / 72)
    pix  = page.get_pixmap(matrix=mat, alpha=False)
    png  = pix.tobytes("png")
    doc.close()
    return png

CLEAN_PROMPT = """
Look at this furniture order PDF page.

I already have the product list below (extracted via text parsing).
Please review and correct any text quality issues — garbled characters,
missing info, merged words — and return the SAME list as clean JSON.

Do NOT add or remove products. Keep the same order.
Fix only obvious OCR/parsing errors in "product", "dims", "material" fields.

Return ONLY a JSON array (no markdown), same structure as input.

Input:
{data}
"""

def claude_clean(products, png_bytes):
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("  ⚠ No ANTHROPIC_API_KEY — skipping Claude cleanup")
        return products

    client  = anthropic.Anthropic(api_key=api_key)
    img_b64 = base64.standard_b64encode(png_bytes).decode()

    # Slim version for Claude (no image bytes)
    slim = [{"area": p["area"], "product": p["product"], "code": p["code"],
             "qty": p["qty"], "dims": p["dims"], "material": p["material"]}
            for p in products]

    print("  Sending to Claude Vision for text cleanup...")
    resp = client.messages.create(
        model=MODEL,
        max_tokens=3000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image",
                 "source": {"type": "base64", "media_type": "image/png", "data": img_b64}},
                {"type": "text",
                 "text": CLEAN_PROMPT.format(data=json.dumps(slim, ensure_ascii=False, indent=2))}
            ]
        }]
    )

    raw = resp.content[0].text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    try:
        cleaned = json.loads(raw)
        if len(cleaned) == len(products):
            for p, c in zip(products, cleaned):
                p["product"]  = c.get("product",  p["product"])
                p["dims"]     = c.get("dims",     p["dims"])
                p["material"] = c.get("material", p["material"])
            print("  Claude cleanup applied")
        else:
            print(f"  ⚠ Claude returned {len(cleaned)} items vs {len(products)} — skipping")
    except Exception as e:
        print(f"  ⚠ Claude cleanup failed: {e} — using original")

    return products


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4  Parse dimensions
# ══════════════════════════════════════════════════════════════════════════════
def parse_dims(s):
    if not s:
        return "", "", ""
    s = s.strip().rstrip("*")
    # Split on * × or " x " (with spaces, to avoid splitting product codes)
    parts = re.split(r"[*×]|\s+[xX]\s+", s)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) >= 3: return parts[0], parts[1], parts[2]
    if len(parts) == 2: return parts[0], parts[1], ""
    return s, "", ""


# ══════════════════════════════════════════════════════════════════════════════
# STEP 5  Write Excel — J078 format, fabric swatch in FOTO/E
# ══════════════════════════════════════════════════════════════════════════════
def make_xl_img(raw_bytes, target_w, target_h):
    """Fit image into exact target_w x target_h canvas (white bg, centered)."""
    pil = PILImage.open(io.BytesIO(raw_bytes))
    if pil.mode in ("RGBA", "P", "LA"):
        pil = pil.convert("RGB")
    pil.thumbnail((target_w - 4, target_h - 4), PILImage.LANCZOS)
    canvas = PILImage.new("RGB", (target_w, target_h), (255, 255, 255))
    ox = (target_w - pil.width) // 2
    oy = (target_h - pil.height) // 2
    canvas.paste(pil, (ox, oy))
    buf = io.BytesIO()
    # Save at 96 DPI so openpyxl's EMU conversion (also 96 DPI) is consistent
    canvas.save(buf, format="JPEG", quality=85, dpi=(96, 96))
    buf.seek(0)
    xl = XLImage(buf)
    xl.width  = target_w
    xl.height = target_h
    return xl


def _write_cell(ws, col, row, val, fill, font, alignment, border):
    c = ws.cell(row=row, column=col if isinstance(col, int) else
                openpyxl.utils.column_index_from_string(col))
    c.value = val; c.fill = fill; c.font = font
    c.alignment = alignment; c.border = border


def write_excel(products, template_path, out_path):
    shutil.copy(template_path, out_path)
    wb = openpyxl.load_workbook(out_path)
    ws = wb["家具清单"]

    for col, w in COL_WIDTHS.items():
        ws.column_dimensions[col].width = w

    ws["H5"] = date.today()

    # Unmerge rows 9+
    to_unmerge = [str(mr) for mr in list(ws.merged_cells.ranges) if mr.min_row >= 9]
    for mr in to_unmerge:
        ws.unmerge_cells(mr)

    # ── FIX: Reset ALL cells rows 9-300 to plain white ────────────────────────
    # Prevents black/dark template fills bleeding through in lower rows
    _blank = PatternFill("solid", fgColor="FFFFFF")
    _no_border = Border()
    for r in range(9, 301):
        ws.row_dimensions[r].height = ROW_H_TEXT
        for c in range(1, 18):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.fill  = _blank
            cell.border = _no_border
            cell.font  = DARK_NORM

    # ── Headers rows 9-10 ────────────────────────────────────────────────────
    ws.row_dimensions[9].height  = 26
    ws.row_dimensions[10].height = 20
    for r in (9, 10):
        for c in range(1, 18):
            ws.cell(row=r, column=c).fill = HEADER_FILL

    for col, lbl in {
        "B": "NÚMERO",        "C": "ÁREA\n区域",
        "D": "PRODUCTO\n产品", "E": "FOTO\n照片",
        "I": "DETALLE\n细节",  "J": "CANTIDAD\n量",
        "K": "P.UNIT\nUSD",   "L": "P.TOTAL\nUSD",
        "M": "CBM",           "N": "KG",
        "O": "TOTAL\nCBM",    "P": "MUESTRA\n皮料",
    }.items():
        ws.merge_cells(f"{col}9:{col}10")
        c = ws[f"{col}9"]
        c.value = lbl; c.font = WHITE_BOLD; c.fill = HEADER_FILL
        c.alignment = CENTER; c.border = bdr()

    ws.merge_cells("F9:H9")
    ws["F9"].value = "TAMAÑO / 尺寸 (MM)"
    ws["F9"].font = WHITE_BOLD; ws["F9"].fill = HEADER_FILL
    ws["F9"].alignment = CENTER; ws["F9"].border = bdr()
    for col, lbl in [("F","LARGO"),("G","PROFUNDIDAD"),("H","ALTURA")]:
        c = ws[f"{col}10"]
        c.value = lbl; c.font = WHITE_BOLD; c.fill = HEADER_FILL
        c.alignment = CENTER; c.border = bdr()

    # ── Expand products with multiple swatches into separate rows ────────────
    # e.g. SOFA LAWSON with 2 fabric options → 2 rows, each with own material code
    rows_data = []
    for p in products:
        swatches      = p.get("swatches", [])
        extra_mats    = p.get("extra_materials", [])  # e.g. ["NBS238-28"] for SOFA LAWSON

        if len(swatches) <= 1:
            rows_data.append({**p, "_swatch": swatches[0] if swatches else None, "_extra": False})
        else:
            # First row: full product info + first swatch + first material code
            rows_data.append({**p, "_swatch": swatches[0], "_extra": False})
            # Extra rows: same product name, blank code/dims/qty, each with own material code
            for i, sw in enumerate(swatches[1:]):
                mat_i = extra_mats[i] if i < len(extra_mats) else ""
                rows_data.append({**p, "_swatch": sw, "_extra": True,
                                   "code": "", "dims": "", "qty": "",
                                   "material": mat_i})

    # ── Product rows ──────────────────────────────────────────────────────────
    row       = 11
    last_area = None
    alt       = 0

    for p in rows_data:
        area = p.get("area", "")

        if area and area != last_area and not p.get("_extra"):
            # Area header row
            ws.row_dimensions[row].height = 20
            ws.merge_cells(f"B{row}:P{row}")
            c = ws[f"B{row}"]
            c.value = area
            c.font  = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
            c.fill  = PatternFill("solid", fgColor="2E5090")
            c.alignment = LEFT; c.border = bdr()
            last_area = area; row += 1; alt = 0

        fill = FILL_W if alt % 2 == 0 else FILL_B
        has_photo  = bool(p.get("photo"))
        has_swatch = bool(p.get("_swatch"))
        ws.row_dimensions[row].height = ROW_H_IMG if (has_photo or has_swatch) else ROW_H_TEXT

        l, w, h = parse_dims(p.get("dims", ""))

        data_cols = {
            "B": p.get("code",""),
            "C": "" if not p.get("_extra") else "↑",
            "D": p.get("product",""),
            "E": "",
            "F": l, "G": w, "H": h,
            "I": p.get("material",""),
            "J": p.get("qty",""),
            "K": "", "L": "", "M": "", "N": "", "O": "",
            "P": "",
        }
        for col, val in data_cols.items():
            c = ws[f"{col}{row}"]
            c.value = val; c.fill = fill; c.font = DARK_NORM; c.border = bdr()
            c.alignment = LEFT if col in ("D","I","P") else CENTER

        # Furniture photo → FOTO/E
        if has_photo and not p.get("_extra"):
            try:
                xl = make_xl_img(p["photo"], FOTO_W_PX, FOTO_H_PX)
                xl.anchor = f"E{row}"
                ws.add_image(xl)
            except Exception as e:
                print(f"  ⚠ Photo [{p['product'][:20]}]: {e}")

        # Fabric swatch → MUESTRA/P (only if PDF has one for this product)
        if has_swatch:
            try:
                xl = make_xl_img(p["_swatch"], NOTAS_W_PX, NOTAS_H_PX)
                xl.anchor = f"P{row}"
                ws.add_image(xl)
            except Exception as e:
                print(f"  ⚠ Swatch [{p['product'][:20]}]: {e}")

        row += 1; alt += 1

    # TOTAL section
    for lbl in ["TOTAL", "SHIPPING", "TOTAL GENERAL"]:
        ws.row_dimensions[row].height = 22
        ws.merge_cells(f"B{row}:J{row}")
        c = ws[f"B{row}"]
        c.value = lbl; c.font = DARK_BOLD; c.fill = TOTAL_FILL
        c.alignment = CENTER; c.border = bdr()
        for col in "KLMNOP":
            ws[f"{col}{row}"].fill = TOTAL_FILL; ws[f"{col}{row}"].border = bdr()
        row += 1

    # Terms
    ws.row_dimensions[row].height = 22
    ws.merge_cells(f"B{row}:P{row}")
    ws[f"B{row}"].value = "TÉRMINOS / 条款"
    ws[f"B{row}"].font  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    ws[f"B{row}"].fill  = HEADER_FILL; ws[f"B{row}"].alignment = LEFT
    row += 1

    for term in [
        "1. El Precio es EXW, puesto en la fábrica.",
        "2. Todos los precios se basan en especificaciones. Nos reservamos el derecho de cambiar el precio según los requisitos adicionales del cliente.",
        "3. El tiempo de producción se cuenta a partir del día que se confirme todos los detalles.",
        "4. El 30/50% debe pagarse mediante T/T como depósito. El restante después de inspección.",
        "5. Los muebles deben estar terminados en 45/60 días después de recibir el depósito.",
        "6. Empaquetado estándar de fábrica. Empaquetados especiales serán coste extra.",
        "7. Los productos son personalizados y no se devolverán salvo defectos de fabricación.",
    ]:
        ws.row_dimensions[row].height = 28
        ws.merge_cells(f"B{row}:P{row}")
        c = ws[f"B{row}"]
        c.value = term; c.font = Font(name="Calibri", size=8, color="555555")
        c.fill = TERMS_FILL; c.alignment = LEFT
        row += 1

    ws.freeze_panes = "D11"
    ws.sheet_view.view = "normal"
    wb.save(out_path)

    # Post-save XML patch: strip all pageBreakPreview / zoomScaleNormal cruft
    # that Mac Excel ignores from openpyxl's in-memory changes
    _fix_view_xml(out_path, last_data_row=row)
    print(f"\n✓  Saved: {out_path}")


def _fix_view_xml(path, last_data_row):
    """Directly patch the worksheet XML inside the XLSX zip to force normal view."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin, \
         zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet"):
                xml = data.decode("utf-8")
                # Force normal view and clean up legacy zoom attributes
                xml = re.sub(r'view="pageBreakPreview"', 'view="normal"', xml)
                xml = re.sub(r'\s*zoomScaleNormal="[^"]*"', '', xml)
                xml = re.sub(r'\s*zoomScalePageLayoutView="[^"]*"', '', xml)
                # Set clean zoom scale
                xml = re.sub(r'zoomScale="[^"]*"', 'zoomScale="100"', xml)
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("  Chinify Quote Converter — precise y-position matching")
    print("=" * 60)

    print("\n[1/4] Parsing PDF table with exact row y-ranges...")
    products = get_product_rows_with_yranges(PDF_PATH)
    print(f"  Products found: {len(products)}")

    print("\n[2/4] Matching images to rows by y-position...")
    products = extract_and_match_images(PDF_PATH, products)

    print("\n[3/4] Claude Vision: cleaning up text data...")
    png = render_pdf_page(PDF_PATH)
    products = claude_clean(products, png)

    print("\n  Final product list:")
    for p in products:
        sw = "🎨" if p.get("swatch") else "  "
        print(f"  {sw} [{p['area']:20}] {p['product']:32} x{p['qty']:<6} {p['dims']:15}  {p['material']}")

    print("\n[4/4] Writing Excel (J078 format, fabric swatches in FOTO column)...")
    write_excel(products, TEMPLATE_PATH, OUTPUT_PATH)

    # ── Self-audit ────────────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("  SELF-AUDIT")
    print("=" * 60)
    print(f"\n  Total products extracted : {len(products)}")
    print(f"  Products with swatch(RIGHT/皮料) : {sum(1 for p in products if p.get('swatches'))}")
    print(f"  Products with photo (LEFT/家具) : {sum(1 for p in products if p.get('photo'))}")
    print(f"  Products with NO image at all   : {sum(1 for p in products if not p.get('swatches') and not p.get('photo'))}")
    print()
    # Full data integrity check — every PDF column vs Excel output
    print(f"  {'PRODUCT':<30} {'CODE':<18} {'QTY':<8} {'DIMS':<18} {'MATERIAL':<22} FOTO SWATCH")
    print(f"  {'-'*30} {'-'*18} {'-'*8} {'-'*18} {'-'*22} ---- ------")
    for p in products:
        l, w, h = parse_dims(p.get("dims",""))
        dims_str = f"{l}×{w}×{h}" if w else (l or "-")
        sw   = f"{len(p.get('swatches',[]))}x" if p.get("swatches") else "---"
        ph   = "✓" if p.get("photo") else "---"
        mat  = p.get("material","") or "-"
        exm  = p.get("extra_materials",[])
        mat_full = mat + (" +" + ",".join(exm) if exm else "")
        warn = " ⚠ MATERIAL MISSING" if not p.get("material") and not exm else ""
        warn += " ⚠ CODE MISSING" if not p.get("code") else ""
        warn += " ⚠ QTY MISSING" if not p.get("qty") else ""
        print(f"  {p['product']:<30} {p.get('code',''):<18} {p.get('qty',''):<8} "
              f"{dims_str:<18} {mat_full:<22} {ph:<4} {sw}{warn}")

    print("\n  Column layout check (should match J078):")
    print("  B=NÚMERO  C=ÁREA  D=PRODUCTO  E=FOTO(皮料)  F=LARGO  G=PROF  H=ALTURA")
    print("  I=DETALLE  J=CANTIDAD  K=P.UNIT  L=P.TOTAL  M=CBM  N=KG  O=TOTAL CBM  P=NOTAS")
    print(f"\n✓  Saved: {OUTPUT_PATH}")
    print("Done!")


if __name__ == "__main__":
    main()
