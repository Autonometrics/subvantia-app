"""
Capa de base de datos de Subvantia.

Funciona con dos backends segun el entorno:

- SQLite (por defecto, fichero subvantia_data.db en este mismo directorio):
  para desarrollo local, cero configuracion.
- PostgreSQL (si existe la variable de entorno DATABASE_URL, p. ej. apuntando
  a un proyecto gratuito de Neon): para produccion, porque el disco de los
  servicios gratuitos de Render es efimero (se borra en cada reinicio, y el
  servicio se reinicia solo con que este inactivo un rato) y con SQLite ahi
  los datos reales no sobrevivirian.

El resto de la aplicacion (app.py, matching.py) no necesita saber cual de
los dos backends esta activo: sigue escribiendo SQL con placeholders "?" y
llamando conn.execute(...)/.fetchone()/.fetchall()/.commit()/.close() exactamente
igual que si fuera sqlite3 puro. Este modulo traduce por debajo al dialecto
de Postgres cuando hace falta (placeholders, AUTOINCREMENT vs SERIAL,
INSERT OR REPLACE vs ON CONFLICT).
"""

import sqlite3
import os
import datetime as dt
from werkzeug.security import generate_password_hash

try:
    import fcntl  # solo existe en Unix (Render); en Windows se usa como no-op mas abajo
except ImportError:
    fcntl = None

DATABASE_URL = os.environ.get("DATABASE_URL")
USANDO_POSTGRES = bool(DATABASE_URL)

DB_PATH = os.path.join(os.path.dirname(__file__), "subvantia_data.db")
# Los ficheros de siembra viven DENTRO de subvantia_app/seed_data/ (no en la
# carpeta padre) para que el arranque en un despliegue nuevo (p. ej. Render,
# donde solo se sube el contenido de subvantia_app/) cargue igualmente las
# convocatorias y clientes de ejemplo.
CONVOCATORIAS_SEED_XLSX = os.path.join(
    os.path.dirname(__file__), "seed_data", "Subvantia_Sprint0_Convocatorias_Piloto.xlsx"
)
CLIENTES_SEED_XLSX = os.path.join(
    os.path.dirname(__file__), "seed_data", "Subvantia_Cartera_Cliente_Ejemplo.xlsx"
)

SCHEMA_SQLITE = """
CREATE TABLE IF NOT EXISTS usuarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    nombre TEXT,
    gestoria TEXT,
    rol TEXT DEFAULT 'gestoria',
    estado_suscripcion TEXT DEFAULT 'activa',
    creado_en TEXT
);

CREATE TABLE IF NOT EXISTS convocatorias (
    id_convocatoria TEXT PRIMARY KEY,
    titulo TEXT,
    origen TEXT,
    comunidad_autonoma TEXT,
    cnae_elegibles TEXT,
    forma_juridica_elegible TEXT,
    antiguedad_min_meses INTEGER,
    antiguedad_max_meses INTEGER,
    cuantia_min_eur REAL,
    cuantia_max_eur REAL,
    fecha_apertura TEXT,
    fecha_cierre TEXT,
    requisitos_texto TEXT,
    url_fuente TEXT,
    estado_convocatoria TEXT,
    verificado_manualmente TEXT
);

CREATE TABLE IF NOT EXISTS clientes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    usuario_id INTEGER NOT NULL,
    cif_nif TEXT,
    nombre TEXT NOT NULL,
    cnae TEXT,
    comunidad_autonoma TEXT,
    provincia TEXT,
    forma_juridica TEXT,
    fecha_alta TEXT,
    gestor_asignado TEXT,
    activo INTEGER DEFAULT 1,
    creado_en TEXT,
    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
);

CREATE TABLE IF NOT EXISTS matches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    usuario_id INTEGER NOT NULL,
    cliente_id INTEGER NOT NULL,
    id_convocatoria TEXT NOT NULL,
    puntuacion INTEGER,
    banda TEXT,
    motivo TEXT,
    cuantia_estimada REAL,
    alerta TEXT,
    calculado_en TEXT,
    FOREIGN KEY (cliente_id) REFERENCES clientes(id),
    FOREIGN KEY (id_convocatoria) REFERENCES convocatorias(id_convocatoria)
);

CREATE TABLE IF NOT EXISTS expedientes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    match_id INTEGER NOT NULL,
    estado TEXT DEFAULT 'detectada',
    creado_en TEXT,
    actualizado_en TEXT,
    notas TEXT,
    FOREIGN KEY (match_id) REFERENCES matches(id)
);

CREATE TABLE IF NOT EXISTS accesos_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    usuario_id INTEGER NOT NULL,
    fecha_hora TEXT,
    ip TEXT,
    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
);
"""

# Identica a la de arriba salvo id INTEGER PRIMARY KEY AUTOINCREMENT -> SERIAL
# PRIMARY KEY, que es como Postgres genera identificadores autoincrementales.
SCHEMA_POSTGRES = """
CREATE TABLE IF NOT EXISTS usuarios (
    id SERIAL PRIMARY KEY,
    email TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    nombre TEXT,
    gestoria TEXT,
    rol TEXT DEFAULT 'gestoria',
    estado_suscripcion TEXT DEFAULT 'activa',
    creado_en TEXT
);

CREATE TABLE IF NOT EXISTS convocatorias (
    id_convocatoria TEXT PRIMARY KEY,
    titulo TEXT,
    origen TEXT,
    comunidad_autonoma TEXT,
    cnae_elegibles TEXT,
    forma_juridica_elegible TEXT,
    antiguedad_min_meses INTEGER,
    antiguedad_max_meses INTEGER,
    cuantia_min_eur REAL,
    cuantia_max_eur REAL,
    fecha_apertura TEXT,
    fecha_cierre TEXT,
    requisitos_texto TEXT,
    url_fuente TEXT,
    estado_convocatoria TEXT,
    verificado_manualmente TEXT
);

CREATE TABLE IF NOT EXISTS clientes (
    id SERIAL PRIMARY KEY,
    usuario_id INTEGER NOT NULL,
    cif_nif TEXT,
    nombre TEXT NOT NULL,
    cnae TEXT,
    comunidad_autonoma TEXT,
    provincia TEXT,
    forma_juridica TEXT,
    fecha_alta TEXT,
    gestor_asignado TEXT,
    activo INTEGER DEFAULT 1,
    creado_en TEXT,
    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
);

CREATE TABLE IF NOT EXISTS matches (
    id SERIAL PRIMARY KEY,
    usuario_id INTEGER NOT NULL,
    cliente_id INTEGER NOT NULL,
    id_convocatoria TEXT NOT NULL,
    puntuacion INTEGER,
    banda TEXT,
    motivo TEXT,
    cuantia_estimada REAL,
    alerta TEXT,
    calculado_en TEXT,
    FOREIGN KEY (cliente_id) REFERENCES clientes(id),
    FOREIGN KEY (id_convocatoria) REFERENCES convocatorias(id_convocatoria)
);

CREATE TABLE IF NOT EXISTS expedientes (
    id SERIAL PRIMARY KEY,
    match_id INTEGER NOT NULL,
    estado TEXT DEFAULT 'detectada',
    creado_en TEXT,
    actualizado_en TEXT,
    notas TEXT,
    FOREIGN KEY (match_id) REFERENCES matches(id)
);

CREATE TABLE IF NOT EXISTS accesos_log (
    id SERIAL PRIMARY KEY,
    usuario_id INTEGER NOT NULL,
    fecha_hora TEXT,
    ip TEXT,
    FOREIGN KEY (usuario_id) REFERENCES usuarios(id)
);
"""

ESTADOS_EXPEDIENTE = ["detectada", "informada", "en_preparacion", "presentada", "resuelta"]
ESTADOS_EXPEDIENTE_LABEL = {
    "detectada": "Detectada",
    "informada": "Informada",
    "en_preparacion": "En preparación",
    "presentada": "Presentada",
    "resuelta": "Resuelta",
}


class _CursorResult:
    """Envuelve el resultado de una consulta Postgres para que fetchone()/
    fetchall() se comporten igual que con sqlite3.Row (acceso por nombre de
    columna), y expone .lastrowid cuando la consulta llevaba "RETURNING"
    (Postgres no tiene lastrowid nativo como sqlite3)."""

    def __init__(self, rows, lastrowid=None):
        self._rows = rows
        self._pos = 0
        self.lastrowid = lastrowid

    def fetchone(self):
        if self._pos >= len(self._rows):
            return None
        fila = self._rows[self._pos]
        self._pos += 1
        return fila

    def fetchall(self):
        resto = self._rows[self._pos:]
        self._pos = len(self._rows)
        return resto

    def __iter__(self):
        return iter(self.fetchall())


class _PostgresConn:
    """Envoltorio minimo sobre una conexion psycopg2 para que el resto del
    codigo (escrito pensando en sqlite3.Connection) funcione sin cambios."""

    def __init__(self, raw_conn):
        self._conn = raw_conn

    @staticmethod
    def _traducir(sql):
        # Unico placeholder usado en todo el proyecto es "?" (estilo sqlite);
        # psycopg2 espera "%s".
        return sql.replace("?", "%s")

    def execute(self, sql, params=()):
        import psycopg2.extras
        cur = self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(self._traducir(sql), params)
        rows = []
        lastrowid = None
        if cur.description is not None:
            rows = cur.fetchall()
            if "returning" in sql.lower() and rows:
                lastrowid = list(rows[0].values())[0]
        cur.close()
        return _CursorResult(rows, lastrowid=lastrowid)

    def executemany(self, sql, seq_of_params):
        cur = self._conn.cursor()
        cur.executemany(self._traducir(sql), list(seq_of_params))
        cur.close()

    def executescript(self, script):
        cur = self._conn.cursor()
        cur.execute(script)
        cur.close()

    def commit(self):
        self._conn.commit()

    def close(self):
        self._conn.close()


def get_db():
    if USANDO_POSTGRES:
        import psycopg2
        raw = psycopg2.connect(DATABASE_URL)
        return _PostgresConn(raw)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    # El directorio de salida de este entorno no soporta el modo de journal por
    # defecto de SQLite (fallaria con "disk I/O error" al crear el fichero
    # -journal). MEMORY evita escribir ese fichero auxiliar; para un uso real
    # en un servidor propio esto no hace falta y puede quitarse.
    conn.execute("PRAGMA journal_mode = MEMORY")
    return conn


import contextlib


@contextlib.contextmanager
def _bloqueo_arranque(conn):
    """Con --workers > 1, gunicorn arranca varios procesos a la vez y CADA
    UNO importa app.py y llama a init_db() de forma independiente. Sin este
    bloqueo, dos procesos pueden ejecutar a la vez el "si la tabla esta vacia,
    siembra" de _seed_usuarios/_seed_demo_clientes: ambos ven la tabla vacia,
    ambos intentan insertar las mismas filas, y el segundo revienta con un
    error de clave duplicada (UNIQUE/PRIMARY KEY) -- ese worker no arranca.
    Esto se reprodujo de verdad al probar el servidor con 2 workers en local.

    La solucion estandar es un lock de arranque: el primer proceso que llega
    se queda con el bloqueo y siembra tranquilo; los demas esperan a que
    termine y luego arrancan ya sobre datos consistentes (sin volver a
    sembrar nada, porque para entonces las tablas ya no estan vacias).

    En Postgres se usa un advisory lock (bloqueo a nivel de sesion, liberado
    automaticamente si la conexion se cae). En SQLite (solo desarrollo local)
    se usa un fichero de bloqueo con flock; si fcntl no esta disponible
    (Windows) simplemente no se bloquea -- en local casi siempre hay un solo
    proceso, así que el riesgo real es minimo."""
    if USANDO_POSTGRES:
        cur = conn._conn.cursor()
        cur.execute("SELECT pg_advisory_lock(727270001)")
        cur.close()
        try:
            yield
        finally:
            cur = conn._conn.cursor()
            cur.execute("SELECT pg_advisory_unlock(727270001)")
            cur.close()
    elif fcntl is not None:
        lock_path = DB_PATH + ".initlock"
        fh = open(lock_path, "w")
        try:
            fcntl.flock(fh, fcntl.LOCK_EX)
            yield
        finally:
            fcntl.flock(fh, fcntl.LOCK_UN)
            fh.close()
    else:
        yield


def init_db(seed_demo_data=True):
    fresh = (not USANDO_POSTGRES) and (not os.path.exists(DB_PATH))
    conn = get_db()
    with _bloqueo_arranque(conn):
        _init_db_con_bloqueo(conn, seed_demo_data)
    return fresh


def _init_db_con_bloqueo(conn, seed_demo_data):
    conn.executescript(SCHEMA_POSTGRES if USANDO_POSTGRES else SCHEMA_SQLITE)
    conn.commit()
    if not USANDO_POSTGRES:
        # Migraciones solo relevantes para bases sqlite creadas antes de estos
        # cambios; una base Postgres nueva ya incluye estas columnas desde el
        # primer CREATE TABLE.
        _migrar_columna_rol(conn)
        _migrar_columna_estado_suscripcion(conn)

    # A diferencia de usuarios/clientes (que solo se siembran la primera vez),
    # las convocatorias se resiembran en CADA arranque: _seed_convocatorias usa
    # INSERT OR REPLACE / ON CONFLICT DO UPDATE por id_convocatoria, asi que
    # reejecutarlo es seguro (nunca duplica) y permite que el Excel maestro
    # (subvantia_app/seed_data/Subvantia_Sprint0_Convocatorias_Piloto.xlsx) sea
    # la fuente de verdad: para anyadir o corregir convocatorias basta con
    # actualizar ese Excel y redesplegar en Render, sin tocar el panel admin.
    #
    # Contrapartida importante: si en el futuro se edita una convocatoria a
    # mano desde el panel de admin en produccion SIN reflejar ese cambio en el
    # Excel maestro, el siguiente despliegue la sobreescribira con los valores
    # del Excel. Mientras el Excel se mantenga sincronizado (como se ha hecho
    # hasta ahora), esto no supone perdida de datos.
    _seed_convocatorias(conn)

    cur = conn.execute("SELECT COUNT(*) AS n FROM usuarios")
    if cur.fetchone()["n"] == 0:
        _seed_usuarios(conn)
    else:
        _asegurar_usuario_admin(conn)

    if seed_demo_data:
        cur = conn.execute("SELECT COUNT(*) AS n FROM clientes")
        if cur.fetchone()["n"] == 0:
            _seed_demo_clientes(conn)
            _calcular_matches_iniciales(conn)

    conn.close()


def _migrar_columna_rol(conn):
    """Bases de datos sqlite creadas antes de introducir roles no tienen la
    columna 'rol' en usuarios. La anyadimos si falta, sin perder los datos
    existentes. (Solo aplica a sqlite: PRAGMA table_info es sintaxis sqlite.)"""
    columnas = [c["name"] for c in conn.execute("PRAGMA table_info(usuarios)").fetchall()]
    if "rol" not in columnas:
        conn.execute("ALTER TABLE usuarios ADD COLUMN rol TEXT DEFAULT 'gestoria'")
        conn.execute("UPDATE usuarios SET rol = 'admin' WHERE email = 'admin@subvantia.test'")
        conn.commit()
        print("Migracion aplicada: columna 'rol' anyadida a usuarios")


def _migrar_columna_estado_suscripcion(conn):
    """Bases sqlite creadas antes de trackear suscripciones no tienen esta
    columna. La anyadimos con todas las cuentas existentes como 'activa' (no
    se corta el acceso a nadie retroactivamente por una migracion)."""
    columnas = [c["name"] for c in conn.execute("PRAGMA table_info(usuarios)").fetchall()]
    if "estado_suscripcion" not in columnas:
        conn.execute("ALTER TABLE usuarios ADD COLUMN estado_suscripcion TEXT DEFAULT 'activa'")
        conn.commit()
        print("Migracion aplicada: columna 'estado_suscripcion' anyadida a usuarios")


def _calcular_matches_iniciales(conn):
    """Los clientes de ejemplo se insertan directamente en la tabla (no pasan por
    la ruta /clientes/importar), asi que hay que lanzar el cruce a mano una vez."""
    import matching
    admin = conn.execute("SELECT id FROM usuarios WHERE email = ?", ("demo@subvantia.test",)).fetchone()
    if not admin:
        return
    total, n_clientes = matching.recalcular_todo(conn, admin["id"])
    print(f"Cruce inicial calculado: {total} matches sobre {n_clientes} clientes")


def _seed_convocatorias(conn):
    import openpyxl
    if not os.path.exists(CONVOCATORIAS_SEED_XLSX):
        print(f"AVISO: no se encontro {CONVOCATORIAS_SEED_XLSX}, no se sembraron convocatorias")
        return
    wb = openpyxl.load_workbook(CONVOCATORIAS_SEED_XLSX, data_only=True)
    ws = wb["Convocatorias"]
    headers = [c.value for c in ws[1]]
    cols = ", ".join(headers)
    placeholders = ", ".join("?" for _ in headers)
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(headers, row))
        # normalizar fechas a texto ISO (openpyxl a veces las da como datetime)
        for campo in ("fecha_apertura", "fecha_cierre"):
            v = d.get(campo)
            if hasattr(v, "strftime"):
                d[campo] = v.strftime("%Y-%m-%d")
        rows.append(tuple(d.get(h) for h in headers))

    # El fichero de siembra es la fuente de verdad COMPLETA de convocatorias:
    # asi es como se ha gestionado este proyecto desde el principio (anadir,
    # corregir o retirar una convocatoria = editar este Excel y redesplegar).
    # El upsert de mas abajo da de alta o actualiza filas por id_convocatoria,
    # pero NUNCA borra -- si una convocatoria se quita de este fichero (p.ej.
    # por estar cerrada y obsoleta), sin este paso seguiria viva para siempre
    # en la base de datos real. Por eso, antes de sembrar, se eliminan de la
    # base las convocatorias que ya no aparecen en el fichero (y en cascada
    # sus matches/expedientes asociados, para no violar las claves foraneas).
    id_col_idx = headers.index("id_convocatoria")
    ids_seed = {r[id_col_idx] for r in rows}
    ids_actuales = {
        f["id_convocatoria"] for f in conn.execute("SELECT id_convocatoria FROM convocatorias").fetchall()
    }
    ids_a_borrar = ids_actuales - ids_seed
    for id_conv in ids_a_borrar:
        conn.execute(
            "DELETE FROM expedientes WHERE match_id IN "
            "(SELECT id FROM matches WHERE id_convocatoria = ?)",
            (id_conv,),
        )
        conn.execute("DELETE FROM matches WHERE id_convocatoria = ?", (id_conv,))
        conn.execute("DELETE FROM convocatorias WHERE id_convocatoria = ?", (id_conv,))
    if ids_a_borrar:
        conn.commit()

    if USANDO_POSTGRES:
        # Postgres no entiende "INSERT OR REPLACE"; el equivalente es
        # ON CONFLICT ... DO UPDATE sobre la clave primaria id_convocatoria.
        actualizaciones = ", ".join(f"{h} = EXCLUDED.{h}" for h in headers if h != "id_convocatoria")
        sql = (
            f"INSERT INTO convocatorias ({cols}) VALUES ({placeholders}) "
            f"ON CONFLICT (id_convocatoria) DO UPDATE SET {actualizaciones}"
        )
    else:
        sql = f"INSERT OR REPLACE INTO convocatorias ({cols}) VALUES ({placeholders})"

    conn.executemany(sql, rows)
    conn.commit()
    if ids_a_borrar:
        print(f"Eliminadas {len(ids_a_borrar)} convocatorias que ya no estan en el fichero de siembra: "
              f"{', '.join(sorted(ids_a_borrar))}")
    print(f"Sembradas {len(rows)} convocatorias desde {os.path.basename(CONVOCATORIAS_SEED_XLSX)}")


def _seed_usuarios(conn):
    # INSERT ... "ignora si ya existe" (por email) en vez de un INSERT liso:
    # es una segunda red de seguridad ademas del advisory lock de mas arriba
    # -- aunque dos procesos entraran aqui a la vez, ninguno revienta con un
    # error de clave duplicada, simplemente el segundo no hace nada.
    insertar = (
        "INSERT INTO usuarios (email, password_hash, nombre, gestoria, rol, creado_en) "
        "VALUES (?, ?, ?, ?, ?, ?) ON CONFLICT (email) DO NOTHING"
        if USANDO_POSTGRES else
        "INSERT OR IGNORE INTO usuarios (email, password_hash, nombre, gestoria, rol, creado_en) "
        "VALUES (?, ?, ?, ?, ?, ?)"
    )
    ahora = dt.datetime.now().isoformat()
    conn.execute(
        insertar,
        (
            "demo@subvantia.test",
            generate_password_hash("subvantia2026"),
            "Usuario demo",
            "Gestoria Demo",
            "gestoria",
            ahora,
        ),
    )
    conn.execute(
        insertar,
        (
            "admin@subvantia.test",
            generate_password_hash("subvantia2026admin"),
            "Administrador Subvantia",
            "Subvantia",
            "admin",
            ahora,
        ),
    )
    conn.commit()
    print("Usuarios creados: demo@subvantia.test / subvantia2026 (gestoria)")
    print("                  admin@subvantia.test / subvantia2026admin (admin)")


def _asegurar_usuario_admin(conn):
    """Si la base de datos ya existia de antes de separar roles, puede tener
    usuarios pero no la cuenta de administrador. La creamos si falta."""
    existe = conn.execute("SELECT 1 FROM usuarios WHERE email = ?", ("admin@subvantia.test",)).fetchone()
    if not existe:
        conn.execute(
            "INSERT INTO usuarios (email, password_hash, nombre, gestoria, rol, creado_en) VALUES (?, ?, ?, ?, ?, ?)",
            (
                "admin@subvantia.test",
                generate_password_hash("subvantia2026admin"),
                "Administrador Subvantia",
                "Subvantia",
                "admin",
                dt.datetime.now().isoformat(),
            ),
        )
        conn.commit()
        print("Cuenta admin creada retroactivamente: admin@subvantia.test / subvantia2026admin")


def _seed_demo_clientes(conn):
    import openpyxl
    if not os.path.exists(CLIENTES_SEED_XLSX):
        print(f"AVISO: no se encontro {CLIENTES_SEED_XLSX}, no se sembraron clientes demo")
        return
    admin = conn.execute("SELECT id FROM usuarios WHERE email = ?", ("demo@subvantia.test",)).fetchone()
    usuario_id = admin["id"]

    wb = openpyxl.load_workbook(CLIENTES_SEED_XLSX, data_only=True)
    ws = wb["Cartera"]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        fecha_alta = row[idx["Fecha de alta"]]
        if hasattr(fecha_alta, "strftime"):
            fecha_alta = fecha_alta.strftime("%Y-%m-%d")
        else:
            fecha_alta = str(fecha_alta)[:10]
        rows.append((
            usuario_id,
            row[idx["CIF/NIF"]],
            row[idx["Nombre"]],
            row[idx["CNAE"]],
            row[idx["Comunidad Autonoma"]],
            row[idx["Provincia"]],
            row[idx["Forma juridica"]],
            fecha_alta,
            row[idx["Gestor asignado"]],
            1,
            dt.datetime.now().isoformat(),
        ))
    conn.executemany(
        """INSERT INTO clientes
           (usuario_id, cif_nif, nombre, cnae, comunidad_autonoma, provincia,
            forma_juridica, fecha_alta, gestor_asignado, activo, creado_en)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        rows,
    )
    conn.commit()
    print(f"Sembrados {len(rows)} clientes demo desde {os.path.basename(CLIENTES_SEED_XLSX)}")


if __name__ == "__main__":
    fresh = init_db()
    backend = "Postgres" if USANDO_POSTGRES else "SQLite"
    print(f"Base de datos inicializada ({backend})" + (" (nueva)" if fresh else ""))
