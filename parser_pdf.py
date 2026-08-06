"""
Lector de cotizaciones en PDF (tipo tabla), p.ej. formato de JINAN DIBAIER.
Extrae cabecera, lineas de producto y las fotos (emparejadas por posicion).
Devuelve la misma estructura que parser_cotizacion.leer_cotizacion:
    {"cabecera": {...}, "lineas": [ {...}, ... ]}
para que el resto de la app funcione igual con Excel o PDF.
"""
import io
import os
import re
import pdfplumber
from PIL import Image

from parser_cotizacion import _norm, _num, _buscar_referencia


# ----- cabecera --------------------------------------------------------------

def _extraer_cabecera_pdf(texto):
    cab = {
        "referencia": None,
        "proveedor": None, "contacto": None, "email": None, "whatsapp": None,
        "cliente": None, "fecha_emision": None, "incoterm": None, "moneda": None,
        "lead_time": None, "forma_pago": None, "transporte": None, "empaque": None,
    }
    for linea in texto.split("\n"):
        s = linea.strip()
        if not s:
            continue
        m = re.match(r"(?:factory|from|supplier|manufacturer)\s*[:：]\s*(.+)", s, re.I)
        if m and not cab["proveedor"]:
            cab["proveedor"] = m.group(1).strip(" .,-")
        m = re.match(r"(?:consignee|to|buyer|messrs|cliente)\s*[:：]\s*(.+)", s, re.I)
        if m and not cab["cliente"]:
            # cortar antes del RIF (J-####) o de espacios grandes
            cab["cliente"] = re.split(r"\s{2,}|\bJ-?\d", m.group(1))[0].strip(" .,-")
        m = re.search(r"([^\s:：]+@[^\s]+)", s)
        if m and not cab["email"]:
            cab["email"] = m.group(1)

    cab["referencia"] = _buscar_referencia(texto)

    low = texto.lower()
    for inc in ("exw", "fob", "cif", "cfr", "ddp", "dap"):
        if re.search(r"\b" + inc + r"\b", low):
            cab["incoterm"] = inc.upper()
            break
    if "$" in texto or "usd" in low:
        cab["moneda"] = "USD"
    elif "rmb" in low or "cny" in low or "¥" in texto:
        cab["moneda"] = "CNY"
    m = re.search(r"(\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2})", texto)
    if m:
        cab["fecha_emision"] = m.group(1)
    return cab


# ----- columnas de la tabla --------------------------------------------------

def _map_columnas_pdf(headers):
    mapa = {}
    pic = None
    for i, h in enumerate(headers):
        if not h:
            continue
        n = _norm(str(h).replace("\n", " "))
        if not n:
            continue
        if "picture" in n or "image" in n or "photo" in n:
            pic = i
        elif "amount" in n and "total" not in mapa:
            mapa["total"] = i
        elif "price" in n and "precio_unit" not in mapa:
            mapa["precio_unit"] = i
        elif "quantity" in n and "cantidad" not in mapa:
            mapa["cantidad"] = i
        elif n.startswith("unit") and "unidad" not in mapa:
            mapa["unidad"] = i
        elif "description" in n and "descripcion" not in mapa:
            mapa["descripcion"] = i
        elif n in ("item", "no", "no.", "s/n", "sn", "#") and "sn" not in mapa:
            mapa["sn"] = i
    return mapa, pic


# columnas extra que enriquecen la descripcion / nota
_EXTRA = ["materials", "thickness", "color", "length", "width", "hs code",
          "supplier item", "remarks"]


def _nota_extra(headers, row):
    partes = []
    for i, h in enumerate(headers):
        if not h or i >= len(row) or not row[i]:
            continue
        n = _norm(str(h).replace("\n", " "))
        if any(k in n for k in _EXTRA):
            val = str(row[i]).replace("\n", " ").strip()
            if val and val != "/":
                etiqueta = str(h).replace("\n", " ").strip()
                partes.append(f"{etiqueta}: {val}")
    return " · ".join(partes) or None


# ----- imagenes --------------------------------------------------------------

def _bytes_de_imagen(im):
    """Devuelve (bytes, extension) de una imagen del PDF, o (None, None)."""
    st = im["stream"]
    try:
        data = st.get_data()
    except Exception:  # noqa: BLE001
        return None, None
    filt = str(st.get("Filter"))
    if "DCTDecode" in filt or "JPXDecode" in filt:
        return data, "jpg"
    if "FlateDecode" in filt:
        w, h = im.get("srcsize", (int(im["width"]), int(im["height"])))
        cs = str(st.get("ColorSpace"))
        mode = "RGB"
        if "Gray" in cs:
            mode = "L"
        elif "CMYK" in cs:
            mode = "CMYK"
        try:
            if len(data) >= w * h * len(mode):
                img = Image.frombytes(mode, (w, h), data[: w * h * len(mode)])
                out = io.BytesIO()
                img.convert("RGB").save(out, "PNG")
                return out.getvalue(), "png"
        except Exception:  # noqa: BLE001
            return None, None
    return None, None


# ----- lectura principal -----------------------------------------------------

def leer_cotizacion_pdf(ruta_pdf, carpeta_imagenes=None):
    if carpeta_imagenes:
        os.makedirs(carpeta_imagenes, exist_ok=True)

    lineas = []
    with pdfplumber.open(ruta_pdf) as pdf:
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
        cabecera = _extraer_cabecera_pdf(texto)

        contador = 0
        for page in pdf.pages:
            imgs = list(page.images)
            for tabla in page.find_tables():
                data = tabla.extract()
                if not data or len(data) < 2:
                    continue
                mapa, _pic = _map_columnas_pdf(data[0])
                if "descripcion" not in mapa:
                    continue
                headers = data[0]
                for ri in range(1, len(data)):
                    row = data[ri]
                    desc = row[mapa["descripcion"]] if mapa["descripcion"] < len(row) else None
                    desc = "" if desc is None else str(desc).replace("\n", " ").strip()
                    prim = _norm(str(row[0])) if row and row[0] else ""
                    if prim.startswith("total") or (not desc and prim in ("", "total")):
                        continue
                    cant = _num(row[mapa["cantidad"]]) if mapa.get("cantidad") is not None and mapa["cantidad"] < len(row) else None
                    if not desc and cant is None:
                        continue
                    contador += 1

                    # emparejar imagen por posicion vertical de la fila
                    img_path = None
                    if carpeta_imagenes and ri < len(tabla.rows) and tabla.rows[ri] is not None:
                        top, bot = tabla.rows[ri].bbox[1], tabla.rows[ri].bbox[3]
                        candidatas = [im for im in imgs
                                      if top <= (im["top"] + im["bottom"]) / 2 < bot]
                        if candidatas:
                            im = max(candidatas, key=lambda x: x["width"] * x["height"])
                            b, ext = _bytes_de_imagen(im)
                            if b:
                                img_path = os.path.join(carpeta_imagenes, f"pdf_fila{contador}.{ext}")
                                with open(img_path, "wb") as f:
                                    f.write(b)

                    lineas.append({
                        "sn": (row[mapa["sn"]] if mapa.get("sn") is not None and mapa["sn"] < len(row) and row[mapa["sn"]] else contador),
                        "descripcion": desc,
                        "cantidad": cant,
                        "unidad": (str(row[mapa["unidad"]]).replace("\n", " ").strip()
                                   if mapa.get("unidad") is not None and mapa["unidad"] < len(row) and row[mapa["unidad"]] else None),
                        "precio_unit": _num(row[mapa["precio_unit"]]) if mapa.get("precio_unit") is not None and mapa["precio_unit"] < len(row) else None,
                        "total": _num(row[mapa["total"]]) if mapa.get("total") is not None and mapa["total"] < len(row) else None,
                        "nota": _nota_extra(headers, row),
                        "imagen": img_path,
                    })

    return {"cabecera": cabecera, "lineas": lineas}


if __name__ == "__main__":
    import sys
    data = leer_cotizacion_pdf(sys.argv[1], carpeta_imagenes="pdf_test")
    print("CABECERA:")
    for k, v in data["cabecera"].items():
        if v:
            print(f"  {k}: {v}")
    print(f"\nLINEAS ({len(data['lineas'])}):")
    for l in data["lineas"]:
        img = os.path.basename(l["imagen"]) if l["imagen"] else "(sin foto)"
        print(f"  #{l['sn']} {str(l['descripcion'])[:32]:32} {l['cantidad']} {l['unidad'] or ''} "
              f"x{l['precio_unit']} = {l['total']}  [{img}]")
