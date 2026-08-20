"""
Compras China - herramienta de gestion de importaciones.
App Streamlit. Ejecutar:  streamlit run app.py
"""
import os
import json
import base64
import tempfile
from functools import lru_cache
import pandas as pd
import streamlit as st

import db
import export_excel
import export_pdf
from parser_cotizacion import leer_cotizacion
from parser_pdf import leer_cotizacion_pdf
from contenedor import calcular, CONTENEDORES

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


@lru_cache(maxsize=1)
def logo_uri():
    """Devuelve el logo Traki como data URI (base64) o None si no existe el archivo."""
    for nombre in ("traki.png", "logo.png", "traki.jpg", "logo.jpg"):
        ruta = os.path.join(BASE_DIR, nombre)
        if os.path.exists(ruta):
            mime = "image/jpeg" if nombre.endswith(("jpg", "jpeg")) else "image/png"
            with open(ruta, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            return f"data:{mime};base64,{b64}"
    return None

st.set_page_config(page_title="Traki · Compras China", page_icon="🛒", layout="wide")
db.init_db()

# --- marca Traki: negro + blanco + rojo -------------------------------------
TRAKI_ROJO = "#E30613"
st.markdown(f"""
<style>
  /* sidebar negro estilo logo */
  [data-testid="stSidebar"] {{ background:#000; }}
  [data-testid="stSidebar"] * {{ color:#fff !important; }}
  [data-testid="stSidebar"] [role="radiogroup"] label {{
     padding:6px 8px; border-radius:8px; }}
  [data-testid="stSidebar"] [role="radiogroup"] label:hover {{
     background:rgba(255,255,255,.08); }}
  /* botones primarios en rojo Traki */
  .stButton>button[kind="primary"], .stButton>button[data-testid="baseButton-primary"] {{
     background:{TRAKI_ROJO}; border-color:{TRAKI_ROJO}; }}
  /* encabezado de marca */
  .traki-header {{
     background:#000; border-radius:14px; padding:16px 22px; margin-bottom:14px;
     display:flex; align-items:center; gap:14px; }}
  /* logotipo "traki" con la gota roja sobre la i */
  .tk-word {{ font-weight:800; color:#fff; letter-spacing:-.045em; line-height:1;
     font-family:'Arial Rounded MT Bold','Segoe UI',Arial,sans-serif;
     display:inline-block; white-space:nowrap; }}
  .tk-i {{ position:relative; display:inline-block; }}
  .tk-drop {{ position:absolute; left:50%; top:.05em;
     width:.34em; height:.34em; background:{TRAKI_ROJO};
     border-radius:0 50% 50% 50%;
     transform:translateX(-50%) rotate(45deg); }}
  .traki-sub {{ color:#d4d4d8; font-size:15px; border-left:2px solid {TRAKI_ROJO};
     padding-left:12px; margin-left:4px; }}
</style>
""", unsafe_allow_html=True)


# marca de palabra reutilizable: "traki" con la i sin punto + gota roja
WORDMARK = ('trak<span class="tk-i">&#305;<span class="tk-drop"></span></span>')


def header(subtitulo):
    uri = logo_uri()
    logo = (f'<img src="{uri}" style="height:52px;border-radius:8px">' if uri
            else f'<span class="tk-word" style="font-size:38px">{WORDMARK}</span>')
    st.markdown(
        f"""<div class="traki-header">
              {logo}
              <span class="traki-sub">{subtitulo}</span>
            </div>""",
        unsafe_allow_html=True,
    )


# color por estado de aprobacion
COLOR_ESTADO = {"Pendiente": "🟡", "Aprobado": "🟢", "Eliminado": "🔴"}
COLOR_ETAPA = {
    "Sin ordenar": "⚪", "Ordenado": "🟠", "En importacion": "🔵",
    "En transito": "🟣", "En aduana": "🟤", "Recibido": "🟢",
}


def mostrar_imagen(col, valor):
    """valor puede ser bytes (desde la DB) o una ruta de archivo (parser)."""
    tiene = False
    if isinstance(valor, (bytes, bytearray)) and len(valor) > 0:
        col.image(bytes(valor), width="stretch")
        tiene = True
    elif isinstance(valor, str) and valor and os.path.exists(valor):
        col.image(valor, width="stretch")
        tiene = True
    if not tiene:
        col.markdown("<div style='text-align:center;font-size:36px'>📦</div>", unsafe_allow_html=True)


# =============================================================================
# BARRA LATERAL - navegacion
# =============================================================================
_logo = logo_uri()
_logo_html = (f'<img src="{_logo}" style="height:46px;border-radius:8px">' if _logo
             else f'<span class="tk-word" style="font-size:30px">{WORDMARK}</span>')
st.sidebar.markdown(
    '<div style="padding:4px 0 10px">'
    f'{_logo_html}'
    '<div style="font-size:12px;font-weight:400;color:#a1a1aa;margin-top:4px">Compras e importaciones</div>'
    '</div>', unsafe_allow_html=True)
seccion = st.sidebar.radio(
    "Menu",
    ["➕ Nueva cotización", "📋 Tablero de compras", "🔎 Buscar", "🧮 Calculadora de contenedor"],
    label_visibility="collapsed",
)
st.sidebar.markdown("---")
cots = db.listar_cotizaciones()
st.sidebar.caption(f"{len(cots)} cotizaciones guardadas")


# =============================================================================
# 1) NUEVA COTIZACION  (subir Excel -> revisar -> aprobar -> guardar)
# =============================================================================
if seccion == "➕ Nueva cotización":
    header("Nueva cotización")
    modo = st.radio(
        "¿Cómo quieres crearla?",
        ["📤 Subir cotización de China", "✍️ Crear a mano (solicitud / compra)",
         "📦 Registrar pedido en camino"],
        horizontal=True,
    )

    archivo = None
    if modo.startswith("📤"):
        st.write("Sube la cotización que manda China (**Excel o PDF**). El sistema lee los productos y sus fotos automáticamente.")
        archivo = st.file_uploader("Archivo de cotización (.xlsx, .xls o .pdf)", type=["xlsx", "xls", "pdf"])
    elif modo.startswith("✍️"):
        # ------- MODO MANUAL: crear cotización a mano, sin archivo -------
        st.write("Crea el pedido a mano. Úsalo para **iniciar una solicitud** (solo producto y cantidad, sin precio) "
                 "o para un **proveedor de confianza** donde van directo a comprar. El precio es opcional.")
        m_nombre = st.text_input("Nombre del pedido", placeholder="Ej: Repuestos Cummins", key="m_nombre")
        c1, c2 = st.columns(2)
        m_prov = c1.text_input("Proveedor", key="m_prov")
        m_cli = c2.text_input("Cliente", value="Traki Distribuidora, C.A.", key="m_cli")
        c3, c4, c5 = st.columns(3)
        m_fecha = c3.date_input("Fecha", format="YYYY-MM-DD", key="m_fecha")
        m_inco = c4.selectbox("Incoterm", ["EXW", "FOB", "CIF", "CFR", "DDP"], key="m_inco")
        m_mon = c5.selectbox("Moneda", ["USD", "CNY"], key="m_mon")
        m_ref = st.text_input("N° de cotización / referencia (opcional)", key="m_ref")
        m_etapa = st.selectbox(
            "Etapa del pedido", db.ETAPAS_PEDIDO, index=0,
            help="Solicitud = solo producto y cantidad (aún sin precio). Elige otra si el pedido ya avanzó.",
            key="m_etapa",
        )

        st.markdown("**Productos** — escribe cada uno y agrega filas con el ➕ de la tabla:")
        UNIDADES = ["pc", "set", "pair", "kg", "g", "ton", "m", "meters", "cm",
                    "box", "roll", "bag", "liter", "m²", "m³", "unit"]
        base = pd.DataFrame([{"Producto": "", "Cantidad": None, "Unidad": "pc",
                              "Precio unit": None, "Estado": "Aprobado"}])
        editado = st.data_editor(
            base, num_rows="dynamic", width="stretch", hide_index=True,
            column_config={
                "Producto": st.column_config.TextColumn(width="large"),
                "Cantidad": st.column_config.NumberColumn(min_value=0.0, format="%.2f"),
                "Unidad": st.column_config.SelectboxColumn(options=UNIDADES, default="pc"),
                "Precio unit": st.column_config.NumberColumn(min_value=0.0, format="%.2f"),
                "Estado": st.column_config.SelectboxColumn(options=db.ESTADOS_APROB, default="Aprobado"),
            }, key="m_editor",
        )

        def _n(x):
            """convierte a numero seguro (NaN/None -> 0)."""
            return float(x) if pd.notna(x) else 0.0

        # solo filas con producto escrito (ignora la fila vacía -> evita el 'nan')
        filas = [r for _, r in editado.iterrows()
                 if pd.notna(r["Producto"]) and str(r["Producto"]).strip()]

        # fotos opcionales por producto (para que salgan en el Excel a China)
        fotos = {}
        if filas:
            with st.expander("📷 Agregar fotos a los productos (opcional)"):
                for idx, r in enumerate(filas):
                    up = st.file_uploader(
                        f"Foto de: {str(r['Producto']).strip()}",
                        type=["png", "jpg", "jpeg"], key=f"m_foto_{idx}",
                    )
                    if up is not None:
                        fotos[idx] = up.getvalue()

        total_m = sum(_n(r["Cantidad"]) * _n(r["Precio unit"])
                      for r in filas if r["Estado"] != "Eliminado")
        st.metric(f"Total ({m_mon})", f"{total_m:,.2f}")

        if st.button("💾 Guardar cotización manual", type="primary", disabled=not filas):
            cab = {"nombre": (m_nombre.strip() or None), "referencia": (m_ref.strip() or None),
                   "proveedor": m_prov or None, "cliente": m_cli or None, "etapa_pedido": m_etapa,
                   "fecha_emision": m_fecha.isoformat() if m_fecha else None,
                   "incoterm": m_inco, "moneda": m_mon}
            lineas_m = []
            for i, r in enumerate(filas, start=1):
                q = _n(r["Cantidad"])
                p = _n(r["Precio unit"])
                unidad = r["Unidad"] if pd.notna(r["Unidad"]) else None
                lineas_m.append({
                    "sn": i, "descripcion": str(r["Producto"]).strip(),
                    "cantidad": q, "cantidad_aprob": q, "unidad": unidad,
                    "precio_unit": p, "total": round(q * p, 2),
                    "estado": r["Estado"], "imagen": fotos.get(i - 1),
                })
            cid = db.guardar_cotizacion(cab, lineas_m, archivo_nombre="(manual)")
            st.session_state["_ult_manual"] = cid

        # tras guardar: mostrar confirmación + botón de descarga (persiste entre reruns)
        if st.session_state.get("_ult_manual"):
            mid = st.session_state["_ult_manual"]
            mls = db.get_lineas(mid)
            mcot = next((c for c in db.listar_cotizaciones() if c["id"] == mid), None)
            if mls and mcot:
                st.success(f"Cotización manual #{mid} guardada. Ya está en el Tablero.")
                if any(l["estado"] == "Aprobado" for l in mls):
                    bx, bp = st.columns(2)
                    bx.download_button(
                        "⬇️ Excel · China",
                        data=export_excel.generar(mcot, mls),
                        file_name=f"OC_manual_{mid}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dlmanx_{mid}", type="primary",
                    )
                    bp.download_button(
                        "⬇️ PDF · revisar/teléfono",
                        data=export_pdf.generar(mcot, mls),
                        file_name=f"OC_manual_{mid}.pdf",
                        mime="application/pdf",
                        key=f"dlmanp_{mid}",
                    )
                if st.button("➕ Crear otra cotización", key="nueva_manual"):
                    st.session_state.pop("_ult_manual", None)
                    st.rerun()

    else:
        # ------- REGISTRAR PEDIDO EN CAMINO: ya comprado, solo seguir la llegada -------
        st.write("Registra un pedido que **ya fue comprado/aprobado** en China, para seguir su **llegada a Venezuela**. "
                 "Puedes adjuntar la proforma y poner los datos del envío.")
        r_nombre = st.text_input("Nombre del pedido", placeholder="Ej: Máquinas surtidas XCMG", key="r_nombre")
        c1, c2 = st.columns(2)
        r_prov = c1.text_input("Proveedor", key="r_prov")
        r_ref = c2.text_input("N° de cotización / referencia (PO)", key="r_ref")
        c3, c4 = st.columns(2)
        r_etapa = c3.selectbox("Etapa actual", db.ETAPAS_PEDIDO,
                               index=db.ETAPAS_PEDIDO.index("Comprado"), key="r_etapa")
        r_mon = c4.selectbox("Moneda", ["USD", "CNY"], key="r_mon")

        st.markdown("**🚚 Datos del envío**")
        s1, s2 = st.columns(2)
        r_cont = s1.text_input("Contenedor", key="r_cont")
        r_guia = s2.text_input("Guía / BL", key="r_guia")
        s3, s4 = st.columns(2)
        r_nav = s3.text_input("Naviera", key="r_nav")
        r_eta = s4.text_input("ETA — llegada estimada a Venezuela", key="r_eta")

        st.markdown("**Productos** (opcional — o solo adjunta la proforma abajo):")
        UNID = ["pc", "set", "pair", "kg", "g", "ton", "m", "meters", "cm", "box", "roll", "bag", "liter", "unit"]
        base_r = pd.DataFrame([{"Producto": "", "Cantidad": None, "Unidad": "pc", "Precio unit": None}])
        edit_r = st.data_editor(
            base_r, num_rows="dynamic", width="stretch", hide_index=True,
            column_config={
                "Producto": st.column_config.TextColumn(width="large"),
                "Cantidad": st.column_config.NumberColumn(min_value=0.0, format="%.2f"),
                "Unidad": st.column_config.SelectboxColumn(options=UNID, default="pc"),
                "Precio unit": st.column_config.NumberColumn(min_value=0.0, format="%.2f"),
            }, key="r_editor",
        )
        r_doc = st.file_uploader("📎 Adjuntar la proforma / documento de China (opcional)",
                                 type=["pdf", "xlsx", "xls", "png", "jpg", "jpeg", "docx", "doc"], key="r_doc")

        if st.button("💾 Registrar pedido en camino", type="primary"):
            def _nr(x):
                return float(x) if pd.notna(x) else 0.0
            filas_r = [row for _, row in edit_r.iterrows()
                       if pd.notna(row["Producto"]) and str(row["Producto"]).strip()]
            cab = {"nombre": (r_nombre.strip() or None), "proveedor": r_prov or None,
                   "referencia": r_ref.strip() or None, "etapa_pedido": r_etapa, "moneda": r_mon,
                   "cliente": "Traki Distribuidora, C.A."}
            lineas_r = []
            for i, row in enumerate(filas_r, start=1):
                q = _nr(row["Cantidad"]); p = _nr(row["Precio unit"])
                lineas_r.append({
                    "sn": i, "descripcion": str(row["Producto"]).strip(),
                    "cantidad": q, "cantidad_aprob": q,
                    "unidad": row["Unidad"] if pd.notna(row["Unidad"]) else None,
                    "precio_unit": p, "total": round(q * p, 2), "estado": "Aprobado", "imagen": None,
                })
            cid = db.guardar_cotizacion(cab, lineas_r, "(en camino)")
            db.actualizar_cotizacion(cid, {
                "contenedor": r_cont.strip() or None, "guia_bl": r_guia.strip() or None,
                "naviera": r_nav.strip() or None, "eta": r_eta.strip() or None,
            })
            if r_doc is not None:
                db.agregar_documento(cid, r_doc.name, "Proforma", r_doc.getvalue(), r_doc.type)
            st.success(f"Pedido #{cid} registrado en etapa «{r_etapa}». Ya está en el Tablero para seguir su llegada. 📦")

    if archivo is not None:
        # parsear solo una vez por archivo
        if st.session_state.get("_archivo_cargado") != archivo.name:
            tmpdir = tempfile.mkdtemp(prefix="cot_")
            ruta_tmp = os.path.join(tmpdir, archivo.name)
            with open(ruta_tmp, "wb") as f:
                f.write(archivo.getbuffer())
            img_dir = os.path.join(tmpdir, "img")
            if archivo.name.lower().endswith(".pdf"):
                data = leer_cotizacion_pdf(ruta_tmp, carpeta_imagenes=img_dir)
            else:
                data = leer_cotizacion(ruta_tmp, carpeta_imagenes=img_dir)
            st.session_state["_data"] = data
            st.session_state["_archivo_cargado"] = archivo.name
            # estado inicial por linea
            st.session_state["_estados"] = {i: "Pendiente" for i in range(len(data["lineas"]))}

        data = st.session_state["_data"]
        cab = data["cabecera"]
        lineas = data["lineas"]

        # --- cabecera ---
        st.subheader("Datos de la cotización")
        c1, c2, c3 = st.columns(3)
        c1.metric("Proveedor", cab.get("proveedor") or "—")
        c2.metric("Cliente", cab.get("cliente") or "—")
        c3.metric("Incoterm", cab.get("incoterm") or "—")
        ref_val = st.text_input(
            "N° de cotización / referencia",
            value=cab.get("referencia") or "",
            help="Detectado del archivo (ej: PO). Puedes corregirlo o escribirlo si no vino.",
            key=f"ref_{st.session_state.get('_archivo_cargado','')}",
        )
        cab["referencia"] = ref_val.strip() or None
        _nom_def = cab.get("nombre") or ""
        if not _nom_def and lineas:
            _nom_def = lineas[0]["descripcion"].splitlines()[0][:45]
        nom_val = st.text_input(
            "Nombre del pedido", value=_nom_def,
            help="Rótulo corto para reconocerlo rápido en el tablero. Puedes editarlo.",
            key=f"nom_{st.session_state.get('_archivo_cargado','')}",
        )
        cab["nombre"] = nom_val.strip() or None
        with st.expander("Ver todos los datos de la cotización"):
            st.write({k: v for k, v in cab.items() if v and not k.startswith("_")})

        st.markdown("---")
        st.subheader(f"Productos ({len(lineas)}) — decide cuáles aprobar")
        st.caption("El jefe aprueba, elimina o deja pendiente cada producto. Si aprueba, puedes ajustar cuántas piezas.")

        # vista rápida en tabla de todo lo que se leyó del archivo
        with st.expander("📋 Ver en tabla todo lo que se cargó"):
            st.dataframe(
                pd.DataFrame([{
                    "#": l.get("sn"),
                    "Producto": (l.get("descripcion") or "").splitlines()[0],
                    "Cantidad": l.get("cantidad"),
                    "Unidad": l.get("unidad"),
                    "Precio": l.get("precio_unit"),
                    "Total": l.get("total"),
                    "Foto": "✅" if l.get("imagen") else "—",
                } for l in lineas]),
                width="stretch", hide_index=True,
            )

        # --- productos, uno por fila con foto ---
        for i, l in enumerate(lineas):
            cimg, cinfo, cest = st.columns([1, 3, 1.3])
            mostrar_imagen(cimg, l.get("imagen"))
            with cinfo:
                st.markdown(f"**#{l['sn']} — {l['descripcion'].splitlines()[0]}**")
                resto = " ".join(l["descripcion"].splitlines()[1:]).strip()
                if resto:
                    st.caption(resto[:180])
                precio = l.get("precio_unit")
                total = l.get("total")
                st.write(
                    f"Cantidad: **{l.get('cantidad') or '—'} {l.get('unidad') or ''}**  ·  "
                    f"P.Unit: **{precio if precio is not None else '—'}**  ·  "
                    f"Total: **{total if total is not None else '—'}**"
                )
            with cest:
                st.session_state["_estados"][i] = st.radio(
                    "Estado", db.ESTADOS_APROB,
                    index=db.ESTADOS_APROB.index(st.session_state["_estados"].get(i, "Pendiente")),
                    key=f"est_{i}", horizontal=False,
                )
                cant_cot = l.get("cantidad") or 0
                if st.session_state["_estados"][i] == "Aprobado":
                    st.number_input(
                        f"Cantidad aprobada (de {cant_cot:g})", min_value=0.0,
                        value=float(cant_cot), step=1.0, key=f"cant_{i}",
                    )
            st.markdown("<hr style='margin:4px 0;border:0;border-top:1px solid #eee'>", unsafe_allow_html=True)

        # --- fotos: agregar o cambiar (útil si el archivo vino sin fotos, p.ej. .xls) ---
        fotos_sub = {}
        with st.expander("📷 Agregar o cambiar fotos de los productos (opcional)"):
            st.caption("Si la cotización vino sin fotos (p.ej. archivos .xls), súbelas aquí; saldrán en el tablero y en el Excel para China.")
            for i, l in enumerate(lineas):
                estado_foto = "✅ tiene foto" if l.get("imagen") else "— sin foto"
                up = st.file_uploader(
                    f"#{l['sn']} {l['descripcion'].splitlines()[0][:45]}  ({estado_foto})",
                    type=["png", "jpg", "jpeg"],
                    key=f"upfoto_{st.session_state.get('_archivo_cargado','')}_{i}",
                )
                if up is not None:
                    fotos_sub[i] = up.getvalue()

        # --- resumen y guardar ---
        def _cant_aprob(i):
            """cantidad aprobada elegida (o la cotizada si no se tocó)."""
            if st.session_state["_estados"][i] == "Aprobado":
                return st.session_state.get(f"cant_{i}", lineas[i].get("cantidad") or 0)
            return lineas[i].get("cantidad")

        aprob = sum(1 for i in range(len(lineas)) if st.session_state["_estados"][i] == "Aprobado")
        elim = sum(1 for i in range(len(lineas)) if st.session_state["_estados"][i] == "Eliminado")
        total_aprob = sum(
            (_cant_aprob(i) or 0) * (lineas[i].get("precio_unit") or 0)
            for i in range(len(lineas)) if st.session_state["_estados"][i] == "Aprobado"
        )
        cA, cB, cC = st.columns(3)
        cA.metric("Aprobados", aprob)
        cB.metric("Eliminados", elim)
        cC.metric(f"Total aprobado ({cab.get('moneda') or ''})", f"{total_aprob:,.2f}")

        if st.button("💾 Guardar cotización", type="primary"):
            for i in range(len(lineas)):
                lineas[i]["estado"] = st.session_state["_estados"][i]
                lineas[i]["cantidad_aprob"] = _cant_aprob(i)
                if i in fotos_sub:
                    lineas[i]["imagen"] = fotos_sub[i]
            cot_id = db.guardar_cotizacion(cab, lineas, archivo_nombre=archivo.name)
            # guardar el archivo original pegado al pedido, para poder verlo/descargarlo luego
            try:
                db.agregar_documento(cot_id, archivo.name, "Cotización original",
                                     archivo.getvalue(), archivo.type)
            except Exception:  # noqa: BLE001
                pass
            st.success(f"Cotización guardada (#{cot_id}). Los productos aprobados ya están en el tablero, "
                       "y el archivo original quedó en «📎 Documentos».")
            for k in ("_data", "_archivo_cargado", "_estados"):
                st.session_state.pop(k, None)


# =============================================================================
# 2) TABLERO DE COMPRAS
# =============================================================================
elif seccion == "📋 Tablero de compras":
    header("Tablero de compras")
    EMOJI_ETAPA = {"Solicitud": "📝", "Cotizado": "💰", "Aprobado": "✅", "Comprado": "🛒", "Recibido": "📦"}
    if not cots:
        st.info("Aún no hay pedidos. Crea uno en «Nueva cotización».")
    cbusca, cfiltro = st.columns([2, 1])
    q = cbusca.text_input("🔎 Buscar pedido", placeholder="Nombre, proveedor, PO o producto…", key="tb_q").strip().lower()
    filtro_etapa = cfiltro.selectbox("Etapa", ["Todas"] + db.ETAPAS_PEDIDO, key="tb_filtro")

    def _match(c):
        if not q:
            return True
        blob = " ".join(str(c.get(k) or "") for k in ("nombre", "proveedor", "referencia", "primer_producto")).lower()
        return q in blob

    cots_f = [c for c in cots
              if (filtro_etapa == "Todas" or (c.get("etapa_pedido") or "Cotizado") == filtro_etapa) and _match(c)]
    if cots and not cots_f:
        st.caption("Ningún pedido coincide con la búsqueda/filtro.")
    st.caption(f"Mostrando {len(cots_f)} de {len(cots)} pedidos.")
    for co in cots_f:
        et = co.get("etapa_pedido") or "Cotizado"
        _prim = (co.get("primer_producto") or "").splitlines()[0] if co.get("primer_producto") else ""
        nombre = co.get("nombre") or _prim or co.get("referencia") or co.get("proveedor") or f"Pedido #{co['id']}"
        try:
            _fe_t = json.loads(co.get("fechas_etapa") or "{}")
        except (ValueError, TypeError):
            _fe_t = {}

        def _dm(iso):
            p = str(iso).split("-")
            return f"{p[2]}/{p[1]}" if len(p) == 3 else str(iso)

        # --- resumen SIEMPRE visible (liviano: no carga fotos ni genera archivos) ---
        st.divider()
        _fdate = _dm(_fe_t.get(et)) if _fe_t.get(et) else ""
        _abierto = st.session_state.get("pedido_abierto") == co["id"]
        _cinfo, _cbtn = st.columns([6, 1])
        with _cinfo:
            st.markdown(f"#### {EMOJI_ETAPA.get(et, '')} {nombre}  ·  #{co['id']}")
            _chips = [f"**{et}**" + (f" · {_fdate}" if _fdate else "")]
            if co.get("proveedor"):
                _chips.append(co["proveedor"])
            if co.get("referencia"):
                _chips.append(f"Ref {co['referencia']}")
            if co.get("n_lineas"):
                _chips.append(f"🧾 {co['n_lineas']} productos")
            if co.get("eta"):
                _chips.append(f"🗓️ Llega {co['eta']}")
            if co.get("n_docs"):
                _chips.append(f"📎 {co['n_docs']} documento(s)")
            st.caption("  ·  ".join(_chips))
        with _cbtn:
            if _abierto:
                if st.button("Cerrar ✕", key=f"cerrar_{co['id']}"):
                    st.session_state["pedido_abierto"] = None
                    st.rerun()
            elif st.button("🔍 Abrir", key=f"abrir_{co['id']}"):
                st.session_state["pedido_abierto"] = co["id"]
                st.rerun()

        # --- detalle PESADO: SOLO del pedido abierto (clave para escalar) ---
        if _abierto:
            m = f"Incoterm {co.get('incoterm') or '—'} · Total {co.get('total_exw') or 0:,.2f} {co.get('moneda') or ''} · {co.get('fecha_emision') or ''}"
            st.caption(m)
            _nom = st.text_input("Nombre del pedido", value=co.get("nombre") or "",
                                 placeholder="Ej: Máquinas surtidas XCMG",
                                 key=f"tb_nom_{co['id']}")
            if _nom.strip() != (co.get("nombre") or ""):
                db.actualizar_cotizacion(co["id"], {"nombre": _nom.strip() or None})
                st.rerun()
            _obs = st.text_area("📝 Observaciones / comentarios", value=co.get("observaciones") or "",
                                placeholder="Cualquier nota o comentario sobre este pedido…",
                                height=68, key=f"tb_obs_{co['id']}")
            if _obs.strip() != (co.get("observaciones") or ""):
                db.actualizar_cotizacion(co["id"], {"observaciones": _obs.strip() or None})
                st.rerun()

            # --- documentos adjuntos (lo primero: ver el archivo de lo comprado) ---
            st.markdown("**📎 Documentos del pedido**")
            docs = db.listar_documentos(co["id"])
            if not docs:
                st.caption("Aún no hay documentos. Adjunta la proforma, factura o guía que manda China.")
            for d in docs:
                dc1, dc2, dc3 = st.columns([4, 1, 1])
                dc1.write(f"📄 **{d['tipo']}** · {d['nombre']}  ·  _{(d.get('subido_en') or '')[:10]}_")
                _full = db.get_documento(d["id"])
                if _full and _full.get("archivo"):
                    dc2.download_button("⬇️", data=_full["archivo"], file_name=d["nombre"],
                                        mime=d.get("mime") or "application/octet-stream",
                                        key=f"docdl_{d['id']}")
                if dc3.button("🗑️", key=f"docdel_{d['id']}"):
                    db.eliminar_documento(d["id"])
                    st.rerun()
            with st.form(f"docup_{co['id']}", clear_on_submit=True):
                u1, u2 = st.columns([1, 2])
                _tipo = u1.selectbox("Tipo", db.TIPOS_DOC, key=f"doctipo_{co['id']}")
                _arch = u2.file_uploader("Archivo (PDF, Excel, imagen…)",
                                         type=["pdf", "xlsx", "xls", "png", "jpg", "jpeg", "docx", "doc"],
                                         key=f"docfile_{co['id']}")
                if st.form_submit_button("📎 Adjuntar documento"):
                    if _arch is not None:
                        db.agregar_documento(co["id"], _arch.name, _tipo, _arch.getvalue(), _arch.type)
                        st.rerun()
                    else:
                        st.warning("Elige un archivo primero.")

            # --- barra visual del flujo (etapas + fecha de cada paso) ---
            _idx = db.ETAPAS_PEDIDO.index(et) if et in db.ETAPAS_PEDIDO else 0
            try:
                _fechas = json.loads(co.get("fechas_etapa") or "{}")
            except (ValueError, TypeError):
                _fechas = {}

            def _fmt(iso):
                p = str(iso).split("-")
                return f"{p[2]}/{p[1]}" if len(p) == 3 else ""

            _items = []
            for _i, _e in enumerate(db.ETAPAS_PEDIDO):
                if _i < _idx:
                    _sty = "background:rgba(22,163,74,.15);color:#16a34a;"        # completada
                elif _i == _idx:
                    _sty = "background:#E30613;color:#fff;font-weight:700;"        # actual
                else:
                    _sty = "background:rgba(120,120,120,.14);color:#9ca3af;"       # pendiente
                _fecha = _fmt(_fechas.get(_e, ""))
                _items.append(
                    '<div style="display:flex;flex-direction:column;align-items:center;gap:3px">'
                    f'<span style="padding:5px 12px;border-radius:999px;font-size:12px;white-space:nowrap;{_sty}">'
                    f'{EMOJI_ETAPA.get(_e, "")} {_e}</span>'
                    f'<span style="font-size:10px;color:#9ca3af;min-height:12px">{_fecha}</span></div>'
                )
            st.markdown('<div style="display:flex;gap:7px;align-items:flex-start;flex-wrap:wrap;margin:4px 0 12px">'
                        + '<span style="color:#9ca3af;font-weight:700;padding-top:6px">›</span>'.join(_items)
                        + '</div>', unsafe_allow_html=True)

            # botones para avanzar / retroceder + selector para saltar a cualquier etapa
            bprev, bnext, bsel = st.columns([1, 1, 2])
            if bprev.button("◀ Atrás", key=f"et_prev_{co['id']}", disabled=_idx == 0):
                db.actualizar_cotizacion(co["id"], {"etapa_pedido": db.ETAPAS_PEDIDO[_idx - 1]})
                st.rerun()
            if bnext.button("Avanzar ▶", key=f"et_next_{co['id']}", disabled=_idx >= len(db.ETAPAS_PEDIDO) - 1, type="primary"):
                db.actualizar_cotizacion(co["id"], {"etapa_pedido": db.ETAPAS_PEDIDO[_idx + 1]})
                st.rerun()
            nueva_et = bsel.selectbox(
                "Cambiar etapa", db.ETAPAS_PEDIDO,
                index=_idx, key=f"tb_etapa_{co['id']}", label_visibility="collapsed",
            )
            if nueva_et != et:
                db.actualizar_cotizacion(co["id"], {"etapa_pedido": nueva_et})
                st.rerun()
            lineas = db.get_lineas(co["id"])
            for l in lineas:
                cimg, cinfo, cacc = st.columns([1, 3, 2])
                mostrar_imagen(cimg, l.get("imagen"))
                cant = l.get("cantidad")
                ca = l.get("cantidad_aprob")
                with cinfo:
                    st.markdown(f"{COLOR_ESTADO.get(l['estado'],'')} **{l['descripcion'].splitlines()[0]}**")
                    linea_txt = f"{cant or '—'} {l.get('unidad') or ''}"
                    if l["estado"] == "Aprobado" and ca is not None and ca != cant:
                        linea_txt += f"  ·  ✅ aprobadas: **{ca:g}** de {cant:g}"
                    precio = l.get("precio_unit") or 0
                    qeff = ca if (l["estado"] == "Aprobado" and ca is not None) else (cant or 0)
                    if precio:
                        linea_txt += f"  ·  Total {qeff * precio:,.2f}"
                    elif l.get("total") is not None:
                        linea_txt += f"  ·  Total {l.get('total')}"
                    st.caption(linea_txt)
                with cacc:
                    nuevo_estado = st.selectbox(
                        "Aprobación", db.ESTADOS_APROB,
                        index=db.ESTADOS_APROB.index(l["estado"]) if l["estado"] in db.ESTADOS_APROB else 0,
                        key=f"tb_est_{l['id']}",
                    )
                    nueva_cant = ca
                    if nuevo_estado == "Aprobado":
                        nueva_cant = st.number_input(
                            "Cant. aprobada", min_value=0.0,
                            value=float(ca if ca is not None else (cant or 0)),
                            step=1.0, key=f"tb_cant_{l['id']}",
                        )
                    cambios = {}
                    if nuevo_estado != l["estado"]:
                        cambios["estado"] = nuevo_estado
                    if nuevo_estado == "Aprobado" and nueva_cant != ca:
                        cambios["cantidad_aprob"] = nueva_cant
                    if cambios:
                        db.actualizar_linea(l["id"], cambios)
                        st.rerun()

            # --- seguimiento del envío (nivel pedido) ---
            st.markdown("**🚚 Seguimiento del envío**")
            with st.form(f"envio_{co['id']}"):
                e1, e2 = st.columns(2)
                v_cont = e1.text_input("Contenedor", value=co.get("contenedor") or "")
                v_guia = e2.text_input("Guía / BL", value=co.get("guia_bl") or "")
                e3, e4 = st.columns(2)
                v_nav = e3.text_input("Naviera", value=co.get("naviera") or "")
                v_eta = e4.text_input("ETA — llegada estimada a Venezuela", value=co.get("eta") or "")
                if st.form_submit_button("💾 Guardar seguimiento", type="primary"):
                    db.actualizar_cotizacion(co["id"], {
                        "contenedor": v_cont.strip() or None, "guia_bl": v_guia.strip() or None,
                        "naviera": v_nav.strip() or None, "eta": v_eta.strip() or None,
                    })
                    st.rerun()

            # --- acciones de la cotización: descargar aprobados / eliminar ---
            st.markdown("---")
            n_aprob = sum(1 for l in lineas if l["estado"] == "Aprobado")
            cbaja, cborra = st.columns([2, 1])
            with cbaja:
                if n_aprob:
                    nombre = (co.get("proveedor") or "proveedor")[:20]
                    bx, bp = st.columns(2)
                    bx.download_button(
                        f"⬇️ Excel · China ({n_aprob})",
                        data=export_excel.generar(co, lineas),
                        file_name=f"OC_{co['id']}_{nombre}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dlx_{co['id']}", type="primary",
                        help="Editable. Para el proveedor en China.",
                    )
                    bp.download_button(
                        "⬇️ PDF · revisar/teléfono",
                        data=export_pdf.generar(co, lineas),
                        file_name=f"OC_{co['id']}_{nombre}.pdf",
                        mime="application/pdf",
                        key=f"dlp_{co['id']}",
                        help="Las fotos se ven en WhatsApp/iPhone. Para que el jefe o tu esposa revisen.",
                    )
                else:
                    st.caption("Aprueba productos para poder descargar la orden.")
            with cborra:
                if not st.session_state.get(f"del_{co['id']}"):
                    if st.button("🗑️ Eliminar", key=f"delbtn_{co['id']}"):
                        st.session_state[f"del_{co['id']}"] = True
                        st.rerun()
                else:
                    st.warning("¿Eliminar esta cotización completa?")
                    if st.button("Sí, eliminar", key=f"delyes_{co['id']}", type="primary"):
                        db.eliminar_cotizacion(co["id"])
                        st.session_state.pop(f"del_{co['id']}", None)
                        st.rerun()
                    if st.button("Cancelar", key=f"delno_{co['id']}"):
                        st.session_state.pop(f"del_{co['id']}", None)
                        st.rerun()


# =============================================================================
# 3) BUSCAR
# =============================================================================
elif seccion == "🔎 Buscar":
    header("Buscar en el histórico")
    st.caption("Ejemplo: escribe «lámpara» y filtra por fecha para ver cuántas se compraron.")
    c1, c2, c3 = st.columns([2, 1, 1])
    texto = c1.text_input("Producto contiene…", "")
    fdesde = c2.date_input("Desde", value=None, format="YYYY-MM-DD")
    fhasta = c3.date_input("Hasta", value=None, format="YYYY-MM-DD")
    estado = st.selectbox("Estado", ["Todos"] + db.ESTADOS_APROB)

    filas, total_cant = db.buscar(
        texto_q=texto,
        fecha_desde=fdesde.isoformat() if fdesde else None,
        fecha_hasta=fhasta.isoformat() if fhasta else None,
        estado=estado,
    )
    st.metric("Total de unidades", f"{total_cant:,.0f}",
              help="Suma de las cantidades (usa la cantidad aprobada cuando el producto está aprobado)")
    if filas:
        df = pd.DataFrame([{
            "Producto": f["descripcion"].splitlines()[0],
            "Cotizada": f.get("cantidad"),
            "Aprobada": f.get("cant_efectiva") if f.get("estado") == "Aprobado" else None,
            "Unidad": f.get("unidad"),
            "Estado": f.get("estado"),
            "Etapa": f.get("etapa"),
            "Proveedor": f.get("proveedor"),
            "Fecha": f.get("fecha_pedido") or f.get("fecha_emision"),
        } for f in filas])
        st.dataframe(df, width="stretch", hide_index=True)
    else:
        st.info("Sin resultados con esos filtros.")


# =============================================================================
# 4) CALCULADORA DE CONTENEDOR
# =============================================================================
elif seccion == "🧮 Calculadora de contenedor":
    header("Calculadora de contenedor")
    st.caption("¿Cuántas unidades caben? Para carga densa (acero) casi siempre manda el peso, no el espacio.")

    c1, c2 = st.columns(2)
    with c1:
        cont = st.selectbox("Tipo de contenedor", list(CONTENEDORES.keys()))
        largo = st.number_input("Largo por unidad (cm)", min_value=0.0, value=200.0)
        ancho = st.number_input("Ancho por unidad (cm)", min_value=0.0, value=100.0)
        alto = st.number_input("Alto/espesor por unidad (cm)", min_value=0.0, value=1.0)
    with c2:
        peso = st.number_input("Peso por unidad (kg)", min_value=0.0, value=157.0,
                               help="Ej: una lámina de acero 2×1 m × 10 mm pesa ~157 kg")
        cantidad = st.number_input("Cantidad que quieres enviar (opcional)", min_value=0, value=0, step=1)

    if st.button("Calcular", type="primary"):
        r = calcular(cont, largo, ancho, alto, peso, cantidad or None)
        if "error" in r:
            st.error(r["error"])
        else:
            st.markdown(f"### Caben **{r['max_por_contenedor']}** unidades por contenedor")
            st.markdown(f"👉 Lo que manda es el **{r['limita']}**")
            a, b, c = st.columns(3)
            a.metric("Límite por peso", f"{r['por_peso']} u", help=f"Carga máx {r['peso_max']:,} kg")
            b.metric("Límite por volumen", f"{r['por_volumen']} u", help=f"Volumen útil {r['vol_util']:.1f} m³ (85%)")
            c.metric("Peso al tope", f"{r['peso_total_lleno']:,.0f} kg")
            if r.get("cantidad"):
                st.markdown("---")
                if r["cabe_en_uno"]:
                    st.success(f"✅ Las {r['cantidad']} unidades caben en 1 contenedor {cont} "
                               f"(peso {r['peso_carga']:,.0f} kg, volumen {r['vol_carga']:.1f} m³).")
                else:
                    st.warning(f"⚠️ Necesitas **{r['contenedores_necesarios']} contenedores** {cont} "
                               f"para {r['cantidad']} unidades (peso total {r['peso_carga']:,.0f} kg).")
