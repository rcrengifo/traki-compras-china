"""
Genera la ORDEN DE COMPRA APROBADA en PDF (con las fotos).
A diferencia del Excel, el PDF muestra las imagenes en cualquier dispositivo
(WhatsApp, iPhone, etc.), ideal para que el jefe/esposa la revisen en el telefono.
"""
import io

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer,
)

ROJO = colors.HexColor("#E30613")
NEGRO = colors.HexColor("#000000")


def _img_flowable(data_bytes, max_w=2.6 * cm, max_h=2.3 * cm):
    """Crea una imagen reportlab escalada para caber en la celda, o '' si falla."""
    try:
        from PIL import Image as PILImage
        iw, ih = PILImage.open(io.BytesIO(bytes(data_bytes))).size
        ratio = min(max_w / iw, max_h / ih)
        return Image(io.BytesIO(bytes(data_bytes)), width=iw * ratio, height=ih * ratio)
    except Exception:  # noqa: BLE001
        return ""


def generar(cot, lineas):
    """cot: dict de la cotizacion. lineas: lista de dicts (se filtran los Aprobados).
    Devuelve los bytes del PDF."""
    aprob = [l for l in lineas if l.get("estado") == "Aprobado"]

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=1.2 * cm, rightMargin=1.2 * cm,
        topMargin=1.2 * cm, bottomMargin=1.2 * cm,
    )
    styles = getSampleStyleSheet()
    st_cell = ParagraphStyle("cell", parent=styles["Normal"], fontSize=8, leading=10)
    st_hdr = ParagraphStyle("hdr", parent=styles["Normal"], fontSize=8, leading=10,
                            textColor=colors.white, alignment=1)
    st_title = ParagraphStyle("title", parent=styles["Title"], fontSize=15,
                              textColor=colors.white, alignment=1)

    elems = []

    # --- titulo ---
    titulo = Table([[Paragraph("ORDEN DE COMPRA APROBADA&nbsp;&nbsp;/&nbsp;&nbsp;APPROVED PURCHASE ORDER", st_title)]],
                   colWidths=[26.0 * cm])
    titulo.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NEGRO),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    elems += [titulo, Spacer(1, 6)]

    # --- cabecera (bilingue) ---
    def fila_info(et, val):
        return [Paragraph(f"<b>{et}</b>", st_cell), Paragraph(str(val or ""), st_cell)]

    info = Table([
        fila_info("Ref. / Quote No:", cot.get("referencia")),
        fila_info("Proveedor / Supplier:", cot.get("proveedor")),
        fila_info("Cliente / To (Buyer):", cot.get("cliente")),
        fila_info("Fecha / Date:", cot.get("fecha_emision")),
        fila_info("Incoterm:", cot.get("incoterm")),
    ], colWidths=[4.5 * cm, 21.5 * cm])
    info.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#EEEEEE")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    elems += [info, Spacer(1, 10)]

    # --- tabla de productos ---
    headers = ["No.", "Producto / Product", "Imagen / Image", "Cant. aprob. /\nApproved Qty",
               "Unidad / Unit", "Precio unit /\nUnit Price", "Monto / Amount"]
    fila_hdr = [Paragraph(h.replace("\n", "<br/>"), st_hdr) for h in headers]
    data = [fila_hdr]

    total_amount = 0.0
    for idx, l in enumerate(aprob, start=1):
        qty = l.get("cantidad_aprob")
        if qty is None:
            qty = l.get("cantidad")
        precio = l.get("precio_unit") or 0
        monto = (qty or 0) * precio
        total_amount += monto

        img = _img_flowable(l.get("imagen")) if l.get("imagen") else ""
        data.append([
            str(idx),
            Paragraph(str(l.get("descripcion") or "").replace("\n", "<br/>"), st_cell),
            img,
            f"{qty:g}" if qty is not None else "",
            str(l.get("unidad") or ""),
            f"{precio:,.2f}",
            f"{monto:,.2f}",
        ])

    # fila total
    data.append(["", "", "", "", "", Paragraph("<b>TOTAL:</b>", st_cell),
                 Paragraph(f"<b>{total_amount:,.2f}</b>", st_cell)])

    tabla = Table(
        data,
        colWidths=[1.1 * cm, 9.0 * cm, 3.0 * cm, 3.0 * cm, 2.4 * cm, 3.2 * cm, 4.3 * cm],
        repeatRows=1,
    )
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), ROJO),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (2, 0), (2, -1), "CENTER"),
        ("ALIGN", (3, 1), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#F4F4F5")),
    ]))
    elems.append(tabla)

    doc.build(elems)
    return buf.getvalue()
