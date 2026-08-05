"""
Compras China - herramienta de gestion de importaciones.
App Streamlit. Ejecutar:  streamlit run app.py
"""
import os
import tempfile
import pandas as pd
import streamlit as st

import db
from parser_cotizacion import leer_cotizacion
from contenedor import calcular, CONTENEDORES

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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
  .traki-word {{ font-size:34px; font-weight:800; color:#fff; letter-spacing:-1px;
     font-family:'Segoe UI',sans-serif; line-height:1; }}
  .traki-dot {{ color:{TRAKI_ROJO}; }}
  .traki-sub {{ color:#d4d4d8; font-size:15px; border-left:2px solid {TRAKI_ROJO};
     padding-left:12px; margin-left:4px; }}
</style>
""", unsafe_allow_html=True)


def header(subtitulo):
    st.markdown(
        f"""<div class="traki-header">
              <span class="traki-word">trak<span class="traki-dot">i</span></span>
              <span class="traki-sub">{subtitulo}</span>
            </div>""",
        unsafe_allow_html=True,
    )


# color por estado de aprobacion
COLOR_ESTADO = {"Pendiente": "🟡", "Aprobado": "🟢", "Quitado": "🔴"}
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
st.sidebar.markdown(
    '<div style="font-size:26px;font-weight:800;letter-spacing:-1px;padding:4px 0 10px">'
    'trak<span style="color:#E30613">i</span>'
    '<div style="font-size:12px;font-weight:400;color:#a1a1aa">Compras e importaciones</div>'
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
    st.write("Sube el Excel que manda China. El sistema lee los productos y sus fotos automáticamente.")

    archivo = st.file_uploader("Archivo de cotización (.xlsx)", type=["xlsx"])

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
        st.caption("El jefe aprueba, quita o deja pendiente cada producto.")

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
            st.markdown("<hr style='margin:4px 0;border:0;border-top:1px solid #eee'>", unsafe_allow_html=True)

        # --- resumen y guardar ---
        aprob = sum(1 for i in range(len(lineas)) if st.session_state["_estados"][i] == "Aprobado")
        quit_ = sum(1 for i in range(len(lineas)) if st.session_state["_estados"][i] == "Quitado")
        total_aprob = sum(
            (lineas[i].get("total") or 0)
            for i in range(len(lineas)) if st.session_state["_estados"][i] == "Aprobado"
        )
        cA, cB, cC = st.columns(3)
        cA.metric("Aprobados", aprob)
        cB.metric("Quitados", quit_)
        cC.metric(f"Total aprobado ({cab.get('moneda') or ''})", f"{total_aprob:,.2f}")

        if st.button("💾 Guardar cotización", type="primary"):
            for i in range(len(lineas)):
                lineas[i]["estado"] = st.session_state["_estados"][i]
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
                with cinfo:
                    st.markdown(f"{COLOR_ESTADO.get(l['estado'],'')} **{l['descripcion'].splitlines()[0]}**")
                    st.caption(f"{l.get('cantidad') or '—'} {l.get('unidad') or ''} · Total {l.get('total') or '—'}")
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
                    if nuevo_estado != l["estado"] or nueva_etapa != l.get("etapa"):
                        db.actualizar_linea(l["id"], {"estado": nuevo_estado, "etapa": nueva_etapa})
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
    st.metric("Total de unidades encontradas", f"{total_cant:,.0f}", help="Suma de las cantidades de todas las líneas encontradas")
    if filas:
        df = pd.DataFrame([{
            "Producto": f["descripcion"].splitlines()[0],
            "Cantidad": f.get("cantidad"),
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
