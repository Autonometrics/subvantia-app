"""
Motor de cruce cliente-convocatoria, adaptado de motor_cruce.py (prototipo Sprint 1+2)
para funcionar contra la base de datos SQLite en vez de ficheros Excel.

Misma logica de puntuacion: CNAE 40 / CCAA 25 / forma juridica 15 / antiguedad 15 /
plazo vigente 5. >=80 con CNAE exacto = alta, 50-79 = media, <50 = baja.
"""

import datetime as dt

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
    conv_ccaa = (conv["comunidad_autonoma"] or "").strip()
    if conv_ccaa.lower() in ("todas", ""):
        return True
    return conv_ccaa.strip().lower() == (cliente_ccaa or "").strip().lower()


def forma_juridica_coincide(cliente_fj, conv):
    raw = (conv["forma_juridica_elegible"] or "").strip().lower()
    if raw == "":
        return True
    opciones = [o.strip() for o in raw.split(",")]
    cliente_fj_norm = (cliente_fj or "").strip().lower()
    if cliente_fj_norm in opciones:
        return True
    if "pyme" in opciones and cliente_fj_norm in ("sl", "sa", "otra"):
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
