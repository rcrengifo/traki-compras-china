"""
Calculadora de contenedor: cuantas unidades caben por VOLUMEN y por PESO.
Para carga densa (acero) casi siempre manda el peso.
"""

# dimensiones internas aprox (m) y carga maxima util (kg)
CONTENEDORES = {
    "20' Standard":  {"largo": 5.90, "ancho": 2.35, "alto": 2.39, "vol": 33.0, "peso_max": 28000},
    "40' Standard":  {"largo": 12.03, "ancho": 2.35, "alto": 2.39, "vol": 67.0, "peso_max": 26500},
    "40' High Cube": {"largo": 12.03, "ancho": 2.35, "alto": 2.69, "vol": 76.0, "peso_max": 28700},
}

# factor de aprovechamiento real (no se llena al 100% por huecos/estiba)
FACTOR_UTIL = 0.85


def calcular(contenedor, largo_cm, ancho_cm, alto_cm, peso_kg, cantidad=None):
    """
    largo/ancho/alto en cm, peso por unidad en kg.
    Devuelve dict con capacidad por volumen, por peso, el limite real,
    y (si se pasa cantidad) si esa cantidad cabe y en cuantos contenedores.
    """
    cont = CONTENEDORES[contenedor]
    vol_unidad = (largo_cm / 100) * (ancho_cm / 100) * (alto_cm / 100)  # m3
    if vol_unidad <= 0 or peso_kg <= 0:
        return {"error": "Las dimensiones y el peso deben ser mayores a cero."}

    vol_util = cont["vol"] * FACTOR_UTIL
    por_volumen = int(vol_util // vol_unidad)
    por_peso = int(cont["peso_max"] // peso_kg)

    limite = min(por_volumen, por_peso)
    manda = "PESO" if por_peso <= por_volumen else "VOLUMEN"

    res = {
        "contenedor": contenedor,
        "vol_unidad": vol_unidad,
        "vol_util": vol_util,
        "peso_max": cont["peso_max"],
        "por_volumen": por_volumen,
        "por_peso": por_peso,
        "max_por_contenedor": limite,
        "limita": manda,
        "peso_total_lleno": round(limite * peso_kg, 1),
        "vol_total_lleno": round(limite * vol_unidad, 2),
    }

    if cantidad:
        cantidad = int(cantidad)
        res["cantidad"] = cantidad
        res["cabe_en_uno"] = cantidad <= limite
        res["contenedores_necesarios"] = (cantidad + limite - 1) // limite if limite else None
        res["peso_carga"] = round(cantidad * peso_kg, 1)
        res["vol_carga"] = round(cantidad * vol_unidad, 2)
    return res
