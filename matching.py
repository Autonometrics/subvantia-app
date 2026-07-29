"""
Motor de cruce cliente-convocatoria, adaptado de motor_cruce.py (prototipo Sprint 1+2)
para funcionar contra la base de datos SQLite en vez de ficheros Excel.

Misma logica de puntuacion: CNAE 40 / CCAA 25 / forma juridica 15 / antiguedad 15 /
plazo vigente 5. >=80 con CNAE exacto = alta, 50-79 = media, <50 = baja.
"""

import datetime as dt
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


def _canonico_forma_juridica(valor):
    norm = _normalizar_texto(valor)
    return _FORMA_JURIDICA_LOOKUP.get(norm, norm)


SECTOR_MAP = {
    "CONV-003": {"62", "63", "86", "87", "88", "55", "56", "10", "11", "51"},
    "CONV-005": {"41", "42", "43", "10", "11", "45"},
    "CONV-007": {"47", "10", "13", "14"},
    "CONV-011": None,
    "CONV-018": {"10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20",
                 "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31",
                 "32", "33"},
}


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
    raw = _normalizar_texto(conv["forma_juridica_elegible"])
    if raw == "":
        return True
    opciones = [_FORMA_JURIDICA_LOOKUP.get(o.strip(), o.strip()) for o in raw.split(",")]
    cliente_fj_canon = _canonico_forma_juridica(cliente_fj)
    if cliente_fj_canon in opciones:
        return True
    if "pyme" in opciones and cliente_fj_canon in ("sl", "sa", "otra"):
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
    hoy = hoy or _hoy()
    if not plazo_vigente(conv, hoy):
        return None

    puntos = 0
    motivos = []

    cnae_ok, _ = cnae_coincide(cliente["cnae"], conv)
    if cnae_ok:
        puntos += 40
        motivos.append("CNAE coincide")

    ccaa_ok = ccaa_coincide(cliente["comunidad_autonoma"], conv)
    if ccaa_ok:
        puntos += 25
        motivos.append(f"comunidad {cliente['comunidad_autonoma']}")

    if forma_juridica_coincide(cliente["forma_juridica"], conv):
        puntos += 15

    meses = antiguedad_meses(cliente["fecha_alta"], hoy)
    if antiguedad_coincide(meses, conv):
        puntos += 15
        if conv["antiguedad_max_meses"] is not None:
            motivos.append(f"antiguedad menor a {conv['antiguedad_max_meses']//12} anyos")

    puntos += 5  # plazo vigente

    banda = "baja"
    if puntos >= 80 and cnae_ok:
        banda = "alta"
    elif puntos >= 50:
        banda = "media"

    motivo_final = ", ".join(motivos) or "coincidencia parcial de criterios secundarios"

    return {
        "puntuacion": puntos,
        "banda": banda,
        "motivo": motivo_final,
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
