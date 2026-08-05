"""
Capa de datos con SQLAlchemy.

- En local (sin configuracion) usa SQLite: archivo compras.db.
- En la nube usa Postgres (Supabase) si existe DATABASE_URL en el entorno
  o en los secrets de Streamlit.

Las imagenes de los productos se guardan COMO BYTES dentro de la base, para
no depender de archivos (Streamlit Cloud borra el sistema de archivos al reiniciar).
"""
import os
from datetime import datetime

from sqlalchemy import (
    create_engine, MetaData, Table, Column, Integer, Text, Float,
    LargeBinary, ForeignKey, select, text, func,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ESTADOS_APROB = ["Pendiente", "Aprobado", "Eliminado"]
ETAPAS = [
    "Sin ordenar", "Ordenado", "En importacion",
    "En transito", "En aduana", "Recibido",
]


def _database_url():
    """Prioridad: env var DATABASE_URL -> st.secrets -> SQLite local."""
    url = os.environ.get("DATABASE_URL")
    if not url:
        try:
            import streamlit as st  # solo disponible dentro de la app
            url = st.secrets.get("DATABASE_URL")
        except Exception:  # noqa: BLE001
            url = None
    if not url:
        url = "sqlite:///" + os.path.join(BASE_DIR, "compras.db").replace("\\", "/")
    # Supabase a veces entrega 'postgres://'; SQLAlchemy quiere 'postgresql://'
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql://", 1)
    return url


_engine = None
_metadata = MetaData()

cotizaciones = Table(
    "cotizaciones", _metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("proveedor", Text), Column("contacto", Text), Column("email", Text),
    Column("whatsapp", Text), Column("cliente", Text), Column("fecha_emision", Text),
    Column("incoterm", Text), Column("moneda", Text), Column("lead_time", Text),
    Column("forma_pago", Text), Column("transporte", Text), Column("empaque", Text),
    Column("total_exw", Float), Column("archivo", Text), Column("creado_en", Text),
)

lineas = Table(
    "lineas", _metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("cotizacion_id", Integer, ForeignKey("cotizaciones.id", ondelete="CASCADE")),
    Column("sn", Text), Column("descripcion", Text), Column("categoria", Text),
    Column("cantidad", Float), Column("cantidad_aprob", Float), Column("unidad", Text),
    Column("precio_unit", Float), Column("total", Float),
    Column("imagen", LargeBinary),        # bytes de la foto (o None)
    Column("estado", Text, default="Pendiente"),
    Column("nota", Text),
    Column("etapa", Text, default="Sin ordenar"),
    Column("guia", Text), Column("contenedor", Text), Column("tracking_num", Text),
    Column("eta", Text), Column("fecha_pedido", Text), Column("fecha_recibido", Text),
    Column("creado_en", Text),
)


def engine():
    global _engine
    if _engine is None:
        url = _database_url()
        kw = {"pool_pre_ping": True}
        _engine = create_engine(url, **kw)
    return _engine


def init_db():
    _metadata.create_all(engine())
    _asegurar_columnas()


def _asegurar_columnas():
    """Agrega columnas nuevas a tablas ya existentes (migracion simple e idempotente).
    Necesario porque create_all no altera tablas que ya existen (p.ej. en Supabase)."""
    from sqlalchemy import inspect
    nuevas = {"lineas": {"cantidad_aprob": "FLOAT"}}
    insp = inspect(engine())
    with engine().begin() as cx:
        for tabla, cols in nuevas.items():
            existentes = {c["name"] for c in insp.get_columns(tabla)}
            for col, tipo in cols.items():
                if col not in existentes:
                    cx.execute(text(f"ALTER TABLE {tabla} ADD COLUMN {col} {tipo}"))


def _row_imagen_a_bytes(src):
    """src puede ser una ruta de archivo (del parser) o ya bytes. Devuelve bytes o None."""
    if not src:
        return None
    if isinstance(src, (bytes, bytearray)):
        return bytes(src)
    if isinstance(src, str) and os.path.exists(src):
        with open(src, "rb") as f:
            return f.read()
    return None


def guardar_cotizacion(cabecera, filas, archivo_nombre=""):
    ahora = datetime.now().isoformat(timespec="seconds")
    total = sum((l.get("total") or 0) for l in filas if l.get("estado", "Pendiente") != "Eliminado")
    with engine().begin() as cx:
        res = cx.execute(cotizaciones.insert().values(
            proveedor=cabecera.get("proveedor"), contacto=cabecera.get("contacto"),
            email=cabecera.get("email"), whatsapp=cabecera.get("whatsapp"),
            cliente=cabecera.get("cliente"), fecha_emision=cabecera.get("fecha_emision"),
            incoterm=cabecera.get("incoterm"), moneda=cabecera.get("moneda"),
            lead_time=cabecera.get("lead_time"), forma_pago=cabecera.get("forma_pago"),
            transporte=cabecera.get("transporte"), empaque=cabecera.get("empaque"),
            total_exw=total, archivo=archivo_nombre, creado_en=ahora,
        ))
        cot_id = res.inserted_primary_key[0]
        for l in filas:
            cant = l.get("cantidad")
            cant_aprob = l.get("cantidad_aprob")
            if cant_aprob is None:
                cant_aprob = cant
            cx.execute(lineas.insert().values(
                cotizacion_id=cot_id, sn=str(l.get("sn", "")),
                descripcion=l.get("descripcion"), categoria=l.get("categoria"),
                cantidad=cant, cantidad_aprob=cant_aprob, unidad=l.get("unidad"),
                precio_unit=l.get("precio_unit"), total=l.get("total"),
                imagen=_row_imagen_a_bytes(l.get("imagen")),
                estado=l.get("estado", "Pendiente"), nota=l.get("nota"),
                etapa="Sin ordenar", creado_en=ahora,
            ))
        return cot_id


def eliminar_cotizacion(cotizacion_id):
    """Borra la cotizacion y todas sus lineas (incluye sus fotos)."""
    with engine().begin() as cx:
        cx.execute(lineas.delete().where(lineas.c.cotizacion_id == cotizacion_id))
        cx.execute(cotizaciones.delete().where(cotizaciones.c.id == cotizacion_id))


def listar_cotizaciones():
    with engine().connect() as cx:
        rows = cx.execute(text("""
            SELECT co.*,
                   (SELECT COUNT(*) FROM lineas l WHERE l.cotizacion_id=co.id) AS n_lineas,
                   (SELECT COUNT(*) FROM lineas l WHERE l.cotizacion_id=co.id AND l.estado='Aprobado') AS n_aprob
            FROM cotizaciones co ORDER BY co.id DESC
        """)).mappings().all()
        return [dict(r) for r in rows]


def get_lineas(cotizacion_id):
    with engine().connect() as cx:
        rows = cx.execute(
            select(lineas).where(lineas.c.cotizacion_id == cotizacion_id).order_by(lineas.c.id)
        ).mappings().all()
        return [dict(r) for r in rows]


def actualizar_linea(linea_id, campos):
    permitidas = {
        "estado", "nota", "categoria", "cantidad", "cantidad_aprob", "precio_unit", "total",
        "etapa", "guia", "contenedor", "tracking_num", "eta",
        "fecha_pedido", "fecha_recibido", "descripcion", "unidad",
    }
    campos = {k: v for k, v in campos.items() if k in permitidas}
    if not campos:
        return
    with engine().begin() as cx:
        cx.execute(lineas.update().where(lineas.c.id == linea_id).values(**campos))


def buscar(texto_q="", fecha_desde=None, fecha_hasta=None, estado=None, etapa=None):
    q = """
        SELECT l.*, co.proveedor, co.fecha_emision, co.cliente
        FROM lineas l JOIN cotizaciones co ON co.id = l.cotizacion_id
        WHERE 1=1
    """
    params = {}
    if texto_q:
        q += " AND (LOWER(l.descripcion) LIKE :like OR LOWER(COALESCE(l.categoria,'')) LIKE :like)"
        params["like"] = f"%{texto_q.lower()}%"
    if fecha_desde:
        q += " AND COALESCE(l.fecha_pedido, co.fecha_emision) >= :fd"
        params["fd"] = fecha_desde
    if fecha_hasta:
        q += " AND COALESCE(l.fecha_pedido, co.fecha_emision) <= :fh"
        params["fh"] = fecha_hasta
    if estado and estado != "Todos":
        q += " AND l.estado = :est"
        params["est"] = estado
    if etapa and etapa != "Todas":
        q += " AND l.etapa = :eta"
        params["eta"] = etapa
    q += " ORDER BY l.id DESC"
    with engine().connect() as cx:
        rows = [dict(r) for r in cx.execute(text(q), params).mappings().all()]
    # cantidad efectiva: si esta aprobado usa la cantidad aprobada; si no, la cotizada
    def _efectiva(r):
        if r.get("estado") == "Aprobado" and r.get("cantidad_aprob") is not None:
            return r["cantidad_aprob"]
        return r.get("cantidad") or 0
    for r in rows:
        r["cant_efectiva"] = _efectiva(r)
    total_cant = sum((_efectiva(r) or 0) for r in rows)
    return rows, total_cant


if __name__ == "__main__":
    init_db()
    print("Base lista. Motor:", engine().url.get_backend_name())
