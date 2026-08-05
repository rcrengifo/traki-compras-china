"""
Compras China - herramienta de gestion de importaciones.
App Streamlit. Ejecutar:  streamlit run app.py
"""
import os
import base64
import tempfile
from functools import lru_cache
import pandas as pd
import streamlit as st

import db
import export_excel
from parser_cotizacion import leer_cotizacion
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
        col.image(bytes(valor), use_container_width=True)
        tiene = True
    elif isinstance(valor, str) and valor and os.path.exists(valor):
        col.image(valor, use_container_width=True)
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
        ["📤 Subir Excel de China", "✍️ Crear manual (proveedor de confianza)"],
        horizontal=True,
    )

    archivo = None
    if modo.startswith("📤"):
        st.write("Sube el Excel que manda China. El sistema lee los productos y sus fotos automáticamente.")
        archivo = st.file_uploader("Archivo de cotización (.xlsx)", type=["xlsx"])
    else:
        # ------- MODO MANUAL: crear cotización a mano, sin archivo -------
        st.write("Crea la cotización a mano. Útil cuando ya tienen un proveedor de confianza y van directo a comprar.")
        c1, c2 = st.columns(2)
        m_prov = c1.text_input("Proveedor", key="m_prov")
        m_cli = c2.text_input("Cliente", value="Traki Distribuidora, C.A.", key="m_cli")
        c3, c4, c5 = st.columns(3)
        m_fecha = c3.date_input("Fecha", format="YYYY-MM-DD", key="m_fecha")
        m_inco = c4.selectbox("Incoterm", ["EXW", "FOB", "CIF", "CFR", "DDP"], key="m_inco")
        m_mon = c5.selectbox("Moneda", ["USD", "CNY"], key="m_mon")

        st.markdown("**Productos** — escribe cada uno y agrega filas con el ➕ de la tabla:")
        UNIDADES = ["pc", "set", "pair", "kg", "g", "ton", "m", "meters", "cm",
                    "box", "roll", "bag", "liter", "m²", "m³", "unit"]
        base = pd.DataFrame([{"Producto": "", "Cantidad": None, "Unidad": "pc",
                              "Precio unit": None, "Estado": "Aprobado"}])
        editado = st.data_editor(
            base, num_rows="dynamic", use_container_width=True, hide_index=True,
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
            cab = {"proveedor": m_prov or None, "cliente": m_cli or None,
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
            st.success(f"Cotización manual guardada (#{cid}) con {len(lineas_m)} productos. Ya está en el tablero.")

    if archivo is not None:
        # parsear solo una vez por archivo
        if st.session_state.get("_archivo_cargado") != archivo.name:
            tmpdir = tempfile.mkdtemp(prefix="cot_")
            ruta_tmp = os.path.join(tmpdir, archivo.name)
            with open(ruta_tmp, "wb") as f:
                f.write(archivo.getbuffer())
            img_dir = os.path.join(tmpdir, "img")
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
        with st.expander("Ver todos los datos de la cotización"):
            st.write({k: v for k, v in cab.items() if v and not k.startswith("_")})

        st.markdown("---")
        st.subheader(f"Productos ({len(lineas)}) — decide cuáles aprobar")
        st.caption("El jefe aprueba, elimina o deja pendiente cada producto. Si aprueba, puedes ajustar cuántas piezas.")

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
            cot_id = db.guardar_cotizacion(cab, lineas, archivo_nombre=archivo.name)
            st.success(f"Cotización guardada (#{cot_id}). Los productos aprobados ya están en el tablero.")
            for k in ("_data", "_archivo_cargado", "_estados"):
                st.session_state.pop(k, None)


# =============================================================================
# 2) TABLERO DE COMPRAS
# =============================================================================
elif seccion == "📋 Tablero de compras":
    header("Tablero de compras")
    if not cots:
        st.info("Aún no hay cotizaciones. Sube una en «Nueva cotización».")
    for co in cots:
        titulo = f"#{co['id']} · {co.get('proveedor') or 'Proveedor?'} · {co.get('fecha_emision') or ''} · {co['n_aprob']}/{co['n_lineas']} aprobados"
        with st.expander(titulo):
            m = f"Incoterm {co.get('incoterm') or '—'} · Total EXW {co.get('total_exw') or 0:,.2f} {co.get('moneda') or ''}"
            st.caption(m)
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
                    nueva_etapa = st.selectbox(
                        "Etapa importación", db.ETAPAS,
                        index=db.ETAPAS.index(l["etapa"]) if l.get("etapa") in db.ETAPAS else 0,
                        key=f"tb_eta_{l['id']}",
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
                    if nueva_etapa != l.get("etapa"):
                        cambios["etapa"] = nueva_etapa
                    if nuevo_estado == "Aprobado" and nueva_cant != ca:
                        cambios["cantidad_aprob"] = nueva_cant
                    if cambios:
                        db.actualizar_linea(l["id"], cambios)
                        st.rerun()

            # --- acciones de la cotización: descargar aprobados / eliminar ---
            st.markdown("---")
            n_aprob = sum(1 for l in lineas if l["estado"] == "Aprobado")
            cbaja, cborra = st.columns([2, 1])
            with cbaja:
                if n_aprob:
                    xlsx = export_excel.generar(co, lineas)
                    st.download_button(
                        f"⬇️ Descargar aprobados para China ({n_aprob})",
                        data=xlsx,
                        file_name=f"OC_aprobada_{co['id']}_{(co.get('proveedor') or 'proveedor')[:20]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{co['id']}", type="primary",
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
        st.dataframe(df, use_container_width=True, hide_index=True)
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
