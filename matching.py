"""
Motor de cruce cliente-convocatoria, adaptado de motor_cruce.py (prototipo Sprint 1+2)
para funcionar contra la base de datos SQLite en vez de ficheros Excel.

Misma logica de puntuacion: CNAE 40 / CCAA 25 / forma juridica 15 / antiguedad 15 /
plazo vigente 5. >=80 con CNAE exacto = alta, 50-79 = media, <50 = baja.
"""

import datetime as dt
import re
import unicodedata

# ---------------------------------------------------------- normalizacion de texto
#
# Los Excel reales llegan con variantes de formato (mayusculas/minusculas,
# puntos, acentos, sinonimos): "S.L." vs "SL", "Castilla Leon" vs
# "Castilla y Leon", "Cataluna" vs "Catalunya", "Pais Vasco" vs "Euskadi"...
# Antes de esta normalizacion, esas variantes hacian que un cliente perdiera
# el match SIN NINGUN AVISO (no da error, simplemente no aparece en el
# cruce) -- el fallo mas peligroso posible porque es invisible. Se corrige
# aqui normalizando texto (mayusculas, acentos, puntos) y con un mapa de
# sinonimos conocidos para CCAA y forma juridica.

def _normalizar_texto(valor):
    valor = (valor or "").strip().lower()
    valor = valor.replace(".", "")
    valor = "".join(c for c in unicodedata.normalize("NFKD", valor) if not unicodedata.combining(c))
    return " ".join(valor.split())


# canonico -> lista de variantes normalizadas que debe reconocer como el mismo valor
_CCAA_ALIASES = {
    "andalucia": ["andalucia"],
    "aragon": ["aragon"],
    "asturias": ["asturias", "principado de asturias"],
    "baleares": ["baleares", "islas baleares", "illes balears"],
    "canarias": ["canarias", "islas canarias"],
    "cantabria": ["cantabria"],
    "castilla y leon": ["castilla y leon", "castilla leon"],
    "castilla-la mancha": ["castilla-la mancha", "castilla la mancha"],
    "cataluna": ["cataluna", "catalunya"],
    "comunidad valenciana": ["comunidad valenciana", "comunitat valenciana", "valencia", "comunidad valenciana (cv)"],
    "extremadura": ["extremadura"],
    "galicia": ["galicia"],
    "madrid": ["madrid", "comunidad de madrid"],
    "murcia": ["murcia", "region de murcia"],
    "navarra": ["navarra", "comunidad foral de navarra"],
    "pais vasco": ["pais vasco", "euskadi", "comunidad autonoma del pais vasco"],
    "rioja": ["la rioja", "rioja"],
    "ceuta": ["ceuta"],
    "melilla": ["melilla"],
}
_CCAA_LOOKUP = {alias: canon for canon, alias_list in _CCAA_ALIASES.items() for alias in alias_list}

_FORMA_JURIDICA_ALIASES = {
    "sl": ["sl", "s l", "sociedad limitada", "srl"],
    "sa": ["sa", "s a", "sociedad anonima"],
    "autonomo": ["autonomo", "autonoma", "trabajador autonomo", "empresario individual", "profesional autonomo"],
    "pyme": ["pyme", "pymes"],
    "otra": ["otra", "otros"],
}
_FORMA_JURIDICA_LOOKUP = {alias: canon for canon, alias_list in _FORMA_JURIDICA_ALIASES.items() for alias in alias_list}


def _canonico_ccaa(valor):
    norm = _normalizar_texto(valor)
    return _CCAA_LOOKUP.get(norm, norm)


# ---------------------------------------------------------- clasificador de forma juridica
#
# El campo "forma_juridica_elegible" de una convocatoria real casi nunca llega
# en un formato limpio de 2-3 opciones: llega como texto libre copiado de la
# convocatoria oficial -- "autonomo (RETA)", "microempresa, pequena empresa,
# mediana empresa", "cualquier forma juridica, incl. CB y sociedad civil",
# "SL, SA, otra (empresa constituida)", "pyme (10-249 empleados)", etc. Un
# lookup exacto por alias (como el de arriba, pensado para el dato LIMPIO del
# cliente) deja sin reconocer la mayoria de estas variantes -- y si ese
# requisito pasa a ser EXCLUYENTE, cualquier variante no reconocida borraria
# el match para todo el mundo. Por eso aqui se clasifica por palabras clave
# en vez de por igualdad exacta, y se ignoran los incisos entre parentesis.
_PATRON_PARENTESIS = re.compile(r"\([^)]*\)")


def _limpiar_incisos(texto):
    return _PATRON_PARENTESIS.sub("", texto).strip()


def _clasificar_forma_juridica(texto):
    t = _limpiar_incisos(_normalizar_texto(texto))
    if not t:
        return ""
    if "cualquier forma juridica" in t or t in ("cualquier", "cualquiera", "cualquier forma"):
        return "cualquiera"
    if re.search(r"\bsrl\b", t) or "sociedad limitada" in t:
        return "sl"
    if re.search(r"\bsl\b", t):
        return "sl"
    if "sociedad anonima" in t:
        return "sa"
    if re.search(r"\bsa\b", t):
        return "sa"
    if (
        "autonomo" in t or "autonoma" in t or "persona fisica" in t or "reta" in t
        or "empresario individual" in t or "trabajador por cuenta propia" in t
    ):
        return "autonomo"
    if (
        "pyme" in t or "pymes" in t or "microempresa" in t or "micropyme" in t
        or "pequena empresa" in t or "mediana empresa" in t or "startup" in t
        or "start up" in t or "nueva creacion" in t
    ):
        return "pyme"
    if (
        t in ("otra", "otros") or "entidad local" in t or "comunidad de bienes" in t
        or "sociedad civil" in t or "asociacion" in t or "persona juridica" in t
        or "empresa constituida" in t
    ):
        return "otra"
    return t


def _canonico_forma_juridica(valor):
    return _clasificar_forma_juridica(valor)


_INDUSTRIA_10_33 = {str(n).zfill(2) for n in range(10, 34)}  # seccion C, industria manufacturera completa

SECTOR_MAP = {
    "CONV-003": {"62", "63", "86", "87", "88", "55", "56", "10", "11", "51"},
    "CONV-005": {"41", "42", "43", "10", "11", "45"},
    "CONV-007": {"47", "10", "13", "14"},
    "CONV-011": None,
    "CONV-018": {"10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20",
                 "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31",
                 "32", "33"},
    # Anadidos al ampliar la exclusion dura a CNAE: estas convocatorias tienen
    # texto libre en cnae_elegibles que el parser de "todos"/"todos excepto"
    # no reconoce. Sin esta entrada, caian en "no_evaluable" -> cnae_ok=False
    # SIEMPRE -> con exclusion dura se habrian quedado sin ningun match para
    # ningun cliente, aunque la convocatoria fuera legitima para su sector.
    "CONV-008": _INDUSTRIA_10_33,  # "industrial" (reindustrializacion de pymes industriales)
    "CONV-013": {"62", "63", "72"},  # "62, 63, 72 (base tecnologica / I+D)"
    "CONV-016": {"55", "56"},  # "55, 56 (alojamiento y actividades turisticas)"
    "CONV-020": None,  # sector "no detallado en fuente oficial, pendiente de confirmar" -> universal (ver aviso ya existente en verificado_manualmente)
    "CONV-021": {str(n).zfill(2) for n in range(10, 33)},  # "industria (CNAE seccion C, 10-32)"
    "CONV-025": None,  # "todos (proyecto validado como innovador)"
    "CONV-029": None,  # "todos (proyectos alineados con la estrategia S3 regional)"
    "CONV-037": None,  # "todos (excluidos sectores del articulo 7 del Rgto FEDER)" -- exclusion no detallada por CNAE
    "CONV-038": {"58", "62", "63", "72"},  # "tecnologico/digital"
    "CONV-039": {"58", "62", "63", "72"},  # "tecnologico/digital"
    "CONV-048": {"58", "62", "63", "72"},  # "tecnologico/digital" (2a ventana de CONV-039, mismos requisitos)
}
# CONV-004, CONV-015 y CONV-017 se retiraron de aqui: sus convocatorias
# (2025, 2025 y 2024 respectivamente) se eliminaron de la base en la revision
# del 2026-07-31 por estar cerradas y fuera de vigencia.


def _hoy():
    return dt.date.today()


def parse_cnae_rule(conv):
    raw = (conv["cnae_elegibles"] or "").strip().lower()
    if raw == "todos":
        return ("todos", None)
    if raw.startswith("todos excepto"):
        codigos = [c.strip() for c in raw.replace("todos excepto", "").split(",")]
        return ("excepto", codigos)
    if conv["id_convocatoria"] in SECTOR_MAP:
        mapa = SECTOR_MAP[conv["id_convocatoria"]]
        if mapa is None:
            return ("universal", None)
        return ("mapa", mapa)
    return ("no_evaluable", None)


def cnae_coincide(cliente_cnae, conv):
    tipo, datos = parse_cnae_rule(conv)
    division = (cliente_cnae or "")[:2]
    if tipo == "todos":
        return True, "coincidencia total (convocatoria abierta a todos los sectores)"
    if tipo == "excepto":
        if division in datos:
            return False, None
        return True, "CNAE no esta en la lista de exclusiones"
    if tipo == "mapa":
        if division in datos:
            return True, "division CNAE coincide con el mapeo sectorial"
        return False, None
    if tipo == "universal":
        return True, "criterio no depende de CNAE (revisar sector manualmente)"
    return False, None


def ccaa_coincide(cliente_ccaa, conv):
    conv_ccaa_norm = _normalizar_texto(conv["comunidad_autonoma"])
    if conv_ccaa_norm in ("todas", ""):
        return True
    return _canonico_ccaa(conv["comunidad_autonoma"]) == _canonico_ccaa(cliente_ccaa)


def forma_juridica_coincide(cliente_fj, conv):
    raw = (conv["forma_juridica_elegible"] or "").strip()
    if raw == "":
        return True
    opciones = [_clasificar_forma_juridica(o) for o in raw.split(",")]
    if "cualquiera" in opciones:
        return True
    cliente_fj_canon = _canonico_forma_juridica(cliente_fj)
    if cliente_fj_canon in opciones:
        return True
    # "pyme" en una convocatoria es una categoria de TAMANO (pequena y mediana
    # empresa), no de forma juridica -- por definicion (UE y española) incluye
    # a autonomos, SL, SA y cualquier otra forma que este por debajo del
    # umbral de tamano. No hay forma de comprobar el tamano real del cliente
    # con los datos que maneja Subvantia, asi que "pyme" no restringe por
    # forma juridica: se admite cualquiera. (Antes esta regla excluia a los
    # autonomos, lo cual era incorrecto -- un autonomo SI es pyme.)
    if "pyme" in opciones:
        return True
    return False


def antiguedad_meses(fecha_alta_str, hoy):
    try:
        d = dt.datetime.strptime(fecha_alta_str[:10], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return 0
    return (hoy.year - d.year) * 12 + (hoy.month - d.month)


def antiguedad_coincide(meses, conv):
    amin = conv["antiguedad_min_meses"]
    amax = conv["antiguedad_max_meses"]
    if amin is not None and meses < amin:
        return False
    if amax is not None and meses > amax:
        return False
    return True


def plazo_vigente(conv, hoy):
    estado = (conv["estado_convocatoria"] or "").strip().lower()
    if estado not in ("abierta", "verificar_urgente", "prevista", "verificar"):
        return False
    fc = conv["fecha_cierre"]
    if fc:
        try:
            fecha_cierre = dt.datetime.strptime(str(fc)[:10], "%Y-%m-%d").date()
            if fecha_cierre < hoy:
                return False
        except ValueError:
            pass
    return True


def calcular_match(cliente, conv, hoy=None):
    """Los 4 requisitos objetivos de una convocatoria son EXCLUYENTES: si el
       cliente no cumple alguno, no hay match en absoluto (no aparece en el
       cruce, ni siquiera en banda baja).
       - Comunidad autonoma: si la convocatoria esta restringida a una CCAA
         concreta y el cliente es de otra, se descarta. Si la convocatoria es
         "todas" (estatal/BOE o fondo europeo), no aplica restriccion.
       - CNAE: si la convocatoria limita a ciertos sectores y el cliente no
         esta en ninguno, se descarta.
       - Forma juridica: si la convocatoria limita a ciertas formas juridicas
         y el cliente no encaja en ninguna, se descarta. "pyme" no cuenta
         como restriccion real porque es una categoria de tamano, no de forma
         juridica: un autonomo tambien es pyme.
       - Antiguedad: si la convocatoria fija un minimo y/o maximo de meses de
         antiguedad y el cliente no esta dentro de ese rango, se descarta.
       - Plazo: si la convocatoria ya cerro o no esta en un estado activo, se
         descarta (logica ya existente en plazo_vigente).
       Como los 4 requisitos son ahora excluyentes, cualquier match que
       sobrevive cumple el 100% de los criterios objetivos -> siempre cae en
       banda "alta". Puntuacion y banda se mantienen por compatibilidad con
       el resto de la app (dashboard, informes), pero ya no hay gradacion
       media/baja: eso pasa a ser una decision de "es o no es elegible",
       no una estimacion de probabilidad."""
    hoy = hoy or _hoy()
    if not plazo_vigente(conv, hoy):
        return None

    if not ccaa_coincide(cliente["comunidad_autonoma"], conv):
        return None

    cnae_ok, _ = cnae_coincide(cliente["cnae"], conv)
    if not cnae_ok:
        return None

    if not forma_juridica_coincide(cliente["forma_juridica"], conv):
        return None

    meses = antiguedad_meses(cliente["fecha_alta"], hoy)
    if not antiguedad_coincide(meses, conv):
        return None

    motivos = ["CNAE coincide", f"comunidad {cliente['comunidad_autonoma']}", "forma juridica elegible"]
    if conv["antiguedad_min_meses"] is not None or conv["antiguedad_max_meses"] is not None:
        motivos.append("cumple el requisito de antiguedad de la convocatoria")

    puntos = 100  # los 4 requisitos objetivos se han verificado arriba: cumplimiento total
    banda = "alta"

    return {
        "puntuacion": puntos,
        "banda": banda,
        "motivo": ", ".join(motivos),
        "cuantia_estimada": conv["cuantia_max_eur"],
        "alerta": "PLAZO SIN CONFIRMAR - verificar antes de usar" if conv["estado_convocatoria"] == "verificar_urgente" else "",
    }


def recalcular_matches_cliente(conn, cliente_id, usuario_id):
    """Borra los matches existentes de un cliente y los vuelve a calcular contra
    todas las convocatorias actuales. Se llama tras importar/editar un cliente."""
    cliente = conn.execute("SELECT * FROM clientes WHERE id = ?", (cliente_id,)).fetchone()
    convocatorias = conn.execute("SELECT * FROM convocatorias").fetchall()
    hoy = _hoy()

    conn.execute("DELETE FROM matches WHERE cliente_id = ?", (cliente_id,))
    nuevos = 0
    for conv in convocatorias:
        m = calcular_match(cliente, conv, hoy)
        if m:
            conn.execute(
                """INSERT INTO matches
                   (usuario_id, cliente_id, id_convocatoria, puntuacion, banda, motivo,
                    cuantia_estimada, alerta, calculado_en)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (usuario_id, cliente_id, conv["id_convocatoria"], m["puntuacion"], m["banda"],
                 m["motivo"], m["cuantia_estimada"], m["alerta"], dt.datetime.now().isoformat()),
            )
            nuevos += 1
    conn.commit()
    return nuevos


def recalcular_todo(conn, usuario_id):
    clientes = conn.execute("SELECT id FROM clientes WHERE usuario_id = ?", (usuario_id,)).fetchall()
    total = 0
    for c in clientes:
        total += recalcular_matches_cliente(conn, c["id"], usuario_id)
    return total, len(clientes)


def recalcular_todas_las_gestorias(conn):
    """Recalcula el cruce para TODAS las gestorias (no solo la del admin que
    esta editando la convocatoria). Imprescindible: al crear/editar/corregir
    una convocatoria, cada gestoria cliente tiene que ver el cambio reflejado
    en su propio cruce -- si solo se recalculara para quien hace el cambio
    (normalmente el admin, que no tiene cartera propia), el resto se quedaria
    con datos de cruce desactualizados sin ningun aviso."""
    gestorias = conn.execute("SELECT id FROM usuarios WHERE rol = 'gestoria'").fetchall()
    total_matches = 0
    total_clientes = 0
    for g in gestorias:
        matches, n_clientes = recalcular_todo(conn, g["id"])
        total_matches += matches
        total_clientes += n_clientes
    return total_matches, total_clientes, len(gestorias)


def generar_email(conn, match_id):
    row = conn.execute(
        """SELECT m.*, c.nombre AS cliente_nombre, conv.titulo AS convocatoria_titulo
           FROM matches m
           JOIN clientes c ON c.id = m.cliente_id
           JOIN convocatorias conv ON conv.id_convocatoria = m.id_convocatoria
           WHERE m.id = ?""",
        (match_id,),
    ).fetchone()
    if not row:
        return None
    cuantia_txt = (
        f"hasta {int(row['cuantia_estimada']):,} EUR".replace(",", ".")
        if row["cuantia_estimada"] else "una cuantia a determinar segun tu proyecto"
    )
    return (
        f"Asunto: Hay una ayuda para {row['cliente_nombre']} que no querras dejar pasar\n\n"
        f"Hola,\n\n"
        f"Hemos revisado tu perfil y hemos detectado que {row['cliente_nombre']} puede ser "
        f"elegible para:\n\n"
        f"  {row['convocatoria_titulo']}  (ref. {row['id_convocatoria']})\n\n"
        f"Motivo del match: {row['motivo']}.\n"
        f"Cuantia estimada: {cuantia_txt}.\n\n"
        f"Si quieres, preparamos ya la documentacion necesaria para presentar la solicitud.\n"
        f"Contesta a este email y lo revisamos juntos esta semana.\n\n"
        f"Un saludo,\n"
        f"Tu gestoria\n"
    )
