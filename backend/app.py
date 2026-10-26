import time
import json
import os
import io
import re

# Cargar .env local si existe (para desarrollo local)
_env_path = os.path.join(os.path.dirname(__file__), ".env")
if os.path.exists(_env_path):
    with open(_env_path) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())
import base64
import zipfile
import threading
from datetime import datetime, timedelta
from flask import Flask, jsonify, request, render_template, send_file
from flask_cors import CORS
from flask_sqlalchemy import SQLAlchemy
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
# Solo el dominio de la app puede llamar a la API desde un navegador.
# Se puede ampliar en Render con CORS_ORIGINS="https://a.com,https://b.com"
_origenes = [o.strip() for o in os.environ.get("CORS_ORIGINS", "https://displays-app.onrender.com").split(",") if o.strip()]
CORS(app, origins=_origenes)

# Tope por petición: evita que un archivo enorme tumbe el servidor (512 MB de RAM)
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "30")) * 1024 * 1024

@app.errorhandler(413)
def _archivo_muy_grande(_e):
    lim = app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024)
    return jsonify({"ok": False, "error": f"El archivo es demasiado grande (máximo {lim} MB)"}), 413

DATABASE_URL = os.environ.get("DATABASE_URL", "")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+pg8000://", 1)
elif DATABASE_URL.startswith("postgresql://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+pg8000://", 1)

if DATABASE_URL:
    app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URL
else:
    DB_PATH = os.path.join(os.path.dirname(__file__), "reportes.db")
    app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"

app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
# Render cierra conexiones inactivas: probar cada conexión antes de usarla y renovarlas cada ~5 min
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"pool_pre_ping": True, "pool_recycle": 280}
db = SQLAlchemy(app)

class Reporte(db.Model):
    id         = db.Column(db.Integer, primary_key=True)
    negocio    = db.Column(db.String(50), default="mundo_vegano", index=True)
    tienda     = db.Column(db.String(200))
    producto   = db.Column(db.String(200))
    item_nbr   = db.Column(db.String(50))
    comentario = db.Column(db.Text)
    foto       = db.Column(db.String(300))
    foto_b64   = db.Column(db.Text)
    usuario    = db.Column(db.String(100))
    fecha      = db.Column(db.String(20))
    semana         = db.Column(db.String(20), default="")
    modificaciones = db.Column(db.Integer, default=0)
    modificado_en  = db.Column(db.String(30))

class Inventario(db.Model):
    id              = db.Column(db.Integer, primary_key=True)
    negocio         = db.Column(db.String(50), default="mundo_vegano", index=True)
    tienda          = db.Column(db.String(200), index=True)
    producto        = db.Column(db.String(200))
    item_nbr        = db.Column(db.String(50))
    cantidad        = db.Column(db.Integer, default=0)
    nombre_empaque  = db.Column(db.String(300))

class Diferencia(db.Model):
    id         = db.Column(db.Integer, primary_key=True)
    negocio    = db.Column(db.String(50), default="mundo_vegano", index=True)
    semana     = db.Column(db.String(20), index=True)
    tienda     = db.Column(db.String(200))
    producto   = db.Column(db.String(200))
    cantidad   = db.Column(db.Integer, default=0)
    estado     = db.Column(db.String(50))   # SIN_FOTO / CON_JUSTIFICACION / OK
    comentario = db.Column(db.Text)

class Usuario(db.Model):
    id       = db.Column(db.Integer, primary_key=True)
    nombre   = db.Column(db.String(150))
    usuario  = db.Column(db.String(100), unique=True)
    password = db.Column(db.String(100))
    rol      = db.Column(db.String(20), default="display")
    # "mundo_vegano", "rooster", "mundo_vegano,rooster" (display en varios) o "ambos" (super-admin / admin general)
    negocio  = db.Column(db.Text, default="mundo_vegano")

class Config(db.Model):
    clave = db.Column(db.String(120), primary_key=True)
    valor = db.Column(db.Text)   # JSON largos: productos desactivados, registro de fotos, etc.

class Negocio(db.Model):
    """Cliente de la app. Se crea/edita desde Sistema → Negocios (superadmin)."""
    __tablename__ = "negocio"
    slug       = db.Column(db.String(50), primary_key=True)   # ej: mundo_vegano
    nombre     = db.Column(db.String(100), nullable=False)
    cadenas    = db.Column(db.String(200), default="")        # códigos separados por coma, ej: WM,MXM,AM
    logo_b64   = db.Column(db.Text)                           # data URL (image/png|jpeg|webp)
    rurales    = db.Column(db.Text, default="")               # una tienda por línea (visita mensual)
    excluidas  = db.Column(db.Text, default="")               # tiendas quitadas (no se visitan), una por línea
    prompt     = db.Column(db.Text)                           # prompt de la IA ({producto} = producto a validar)
    activo     = db.Column(db.Boolean, default=True)
    orden      = db.Column(db.Integer, default=0)
    creado_en  = db.Column(db.String(30))

class SemanaConfig(db.Model):
    __tablename__ = "semana_config"
    id         = db.Column(db.Integer, primary_key=True)
    codigo     = db.Column(db.String(20), unique=True, nullable=False)  # ej: 2026-S37
    nombre     = db.Column(db.String(100), nullable=False)              # ej: Semana 2 de septiembre
    fecha_ini  = db.Column(db.String(10))   # YYYY-MM-DD
    fecha_fin  = db.Column(db.String(10))   # YYYY-MM-DD
    activa     = db.Column(db.Boolean, default=False)
    creada_en  = db.Column(db.String(30))

class Liquidacion(db.Model):
    __tablename__ = "liquidacion"
    id         = db.Column(db.Integer, primary_key=True)
    negocio    = db.Column(db.String(50), default="mundo_vegano", index=True)
    tienda     = db.Column(db.String(200), index=True)
    semana     = db.Column(db.String(20), index=True)
    tipo       = db.Column(db.String(50), default="Liquidación")
    usuario    = db.Column(db.String(100))
    comentario = db.Column(db.Text)
    foto1_b64  = db.Column(db.Text)
    foto2_b64  = db.Column(db.Text)
    foto3_b64  = db.Column(db.Text)
    fecha      = db.Column(db.String(30))

class ValidacionIA(db.Model):
    __tablename__ = "validacion_ia"
    id           = db.Column(db.Integer, primary_key=True)
    negocio      = db.Column(db.String(50), default="mundo_vegano", index=True)
    reporte_id   = db.Column(db.Integer, db.ForeignKey("reporte.id"), nullable=True)
    semana       = db.Column(db.String(20), index=True)
    tienda       = db.Column(db.String(200))
    producto        = db.Column(db.String(200))
    nombre_empaque  = db.Column(db.String(300))
    marca           = db.Column(db.String(200))
    # PENDIENTE / PROCESANDO / APROBADO / RECHAZADO / REVISAR / ERROR
    estado       = db.Column(db.String(20), default="PENDIENTE", index=True)
    confianza    = db.Column(db.String(10))   # alta / media / baja
    motivo       = db.Column(db.Text)         # explicación corta de Claude
    tokens_input = db.Column(db.Integer, default=0)
    tokens_output= db.Column(db.Integer, default=0)
    costo_usd    = db.Column(db.Float, default=0.0)
    intentos     = db.Column(db.Integer, default=0)
    creado_en    = db.Column(db.String(30))
    procesado_en = db.Column(db.String(30))

USUARIOS_INICIALES = [
    {"nombre": "Display 1",    "usuario": "display1",  "password": "1234",       "rol": "display",     "negocio": "mundo_vegano"},
    {"nombre": "Supervisor MV","usuario": "admin",     "password": "admin123",   "rol": "supervisor",  "negocio": "mundo_vegano"},
    {"nombre": "Super Admin",  "usuario": "superadmin","password": "super2024",  "rol": "superadmin",  "negocio": "ambos"},
    {"nombre": "Display R1",   "usuario": "rdisplay1", "password": "rooster1234", "rol": "display",    "negocio": "rooster"},
    {"nombre": "Admin Rooster","usuario": "radmin",    "password": "radmin123",   "rol": "supervisor", "negocio": "rooster"},
]

def _migrar_columnas():
    # db.create_all() no altera tablas existentes; agrega columnas nuevas a mano.
    with db.engine.connect() as conn:
        try:
            if db.engine.dialect.name == "postgresql":   # displays con varios negocios
                conn.execute(db.text("ALTER TABLE usuario ALTER COLUMN negocio TYPE TEXT"))
                conn.commit()
        except Exception:
            conn.rollback()

        try:
            if db.engine.dialect.name == "postgresql":
                conn.execute(db.text("ALTER TABLE reporte ADD COLUMN IF NOT EXISTS foto_b64 TEXT"))
            else:
                conn.execute(db.text("ALTER TABLE reporte ADD COLUMN foto_b64 TEXT"))
            conn.commit()
        except Exception:
            conn.rollback()

        try:
            if db.engine.dialect.name == "postgresql":
                conn.execute(db.text("ALTER TABLE inventario ADD COLUMN IF NOT EXISTS nombre_empaque VARCHAR(300)"))
            else:
                conn.execute(db.text("ALTER TABLE inventario ADD COLUMN nombre_empaque VARCHAR(300)"))
            conn.commit()
        except Exception:
            conn.rollback()

        try:
            if db.engine.dialect.name == "postgresql":
                conn.execute(db.text("ALTER TABLE validacion_ia ADD COLUMN IF NOT EXISTS nombre_empaque VARCHAR(300)"))
            else:
                conn.execute(db.text("ALTER TABLE validacion_ia ADD COLUMN nombre_empaque VARCHAR(300)"))
            conn.commit()
        except Exception:
            conn.rollback()

        try:
            if db.engine.dialect.name == "postgresql":
                conn.execute(db.text("ALTER TABLE reporte ADD COLUMN IF NOT EXISTS modificaciones INTEGER DEFAULT 0"))
                conn.execute(db.text("ALTER TABLE reporte ADD COLUMN IF NOT EXISTS modificado_en VARCHAR(30)"))
            else:
                conn.execute(db.text("ALTER TABLE reporte ADD COLUMN modificaciones INTEGER DEFAULT 0"))
                conn.commit()
                conn.execute(db.text("ALTER TABLE reporte ADD COLUMN modificado_en VARCHAR(30)"))
            conn.commit()
        except Exception:
            conn.rollback()

        try:
            if db.engine.dialect.name == "postgresql":
                conn.execute(db.text("ALTER TABLE liquidacion ADD COLUMN IF NOT EXISTS tipo VARCHAR(50) DEFAULT 'Liquidación'"))
            else:
                conn.execute(db.text("ALTER TABLE liquidacion ADD COLUMN tipo VARCHAR(50) DEFAULT 'Liquidación'"))
            conn.commit()
        except Exception:
            conn.rollback()

        try:
            if db.engine.dialect.name == "postgresql":
                conn.execute(db.text("ALTER TABLE reporte ADD COLUMN IF NOT EXISTS item_nbr VARCHAR(50)"))
                conn.execute(db.text("ALTER TABLE inventario ADD COLUMN IF NOT EXISTS item_nbr VARCHAR(50)"))
            else:
                conn.execute(db.text("ALTER TABLE reporte ADD COLUMN item_nbr VARCHAR(50)"))
                conn.execute(db.text("ALTER TABLE inventario ADD COLUMN item_nbr VARCHAR(50)"))
            conn.commit()
        except Exception:
            conn.rollback()

        # Migrar tabla validacion_ia si no existe (db.create_all la crea, pero por si acaso)
        try:
            if db.engine.dialect.name == "postgresql":
                conn.execute(db.text("""
                    CREATE TABLE IF NOT EXISTS validacion_ia (
                        id SERIAL PRIMARY KEY,
                        reporte_id INTEGER REFERENCES reporte(id) ON DELETE SET NULL,
                        semana VARCHAR(20),
                        tienda VARCHAR(200),
                        producto VARCHAR(200),
                        marca VARCHAR(200),
                        estado VARCHAR(20) DEFAULT 'PENDIENTE',
                        confianza VARCHAR(10),
                        motivo TEXT,
                        tokens_input INTEGER DEFAULT 0,
                        tokens_output INTEGER DEFAULT 0,
                        costo_usd FLOAT DEFAULT 0,
                        intentos INTEGER DEFAULT 0,
                        creado_en VARCHAR(30),
                        procesado_en VARCHAR(30)
                    )
                """))
            conn.commit()
        except Exception:
            conn.rollback()

        # --- Multi-tenant: columna negocio ---
        for tabla in ("reporte", "inventario", "diferencia", "liquidacion", "validacion_ia"):
            try:
                if db.engine.dialect.name == "postgresql":
                    conn.execute(db.text(
                        f"ALTER TABLE {tabla} ADD COLUMN IF NOT EXISTS negocio VARCHAR(50) DEFAULT 'mundo_vegano'"
                    ))
                else:
                    conn.execute(db.text(
                        f"ALTER TABLE {tabla} ADD COLUMN negocio VARCHAR(50) DEFAULT 'mundo_vegano'"
                    ))
                conn.commit()
            except Exception:
                conn.rollback()

        try:
            if db.engine.dialect.name == "postgresql":
                conn.execute(db.text(
                    "ALTER TABLE usuario ADD COLUMN IF NOT EXISTS negocio TEXT DEFAULT 'mundo_vegano'"
                ))
            else:
                conn.execute(db.text(
                    "ALTER TABLE usuario ADD COLUMN negocio TEXT DEFAULT 'mundo_vegano'"
                ))
            conn.commit()
        except Exception:
            conn.rollback()

def _migrar_negocio():
    with db.engine.connect() as conn:
        # claves de config por negocio pueden pasar de 50 caracteres
        if db.engine.dialect.name == "postgresql":
            try:
                conn.execute(db.text("ALTER TABLE config ALTER COLUMN clave TYPE VARCHAR(120)"))
                conn.execute(db.text("ALTER TABLE config ALTER COLUMN valor TYPE TEXT"))
                conn.commit()
            except Exception:
                conn.rollback()
        try:
            if db.engine.dialect.name == "postgresql":
                conn.execute(db.text("ALTER TABLE negocio ADD COLUMN IF NOT EXISTS excluidas TEXT DEFAULT ''"))
            else:
                conn.execute(db.text("ALTER TABLE negocio ADD COLUMN excluidas TEXT DEFAULT ''"))
            conn.commit()
        except Exception:
            conn.rollback()
        try:
            if db.engine.dialect.name == "postgresql":
                conn.execute(db.text("ALTER TABLE negocio ADD COLUMN IF NOT EXISTS prompt TEXT"))
            else:
                conn.execute(db.text("ALTER TABLE negocio ADD COLUMN prompt TEXT"))
            conn.commit()
        except Exception:
            conn.rollback()

def _seed_usuarios():
    try:
        if Usuario.query.count() == 0:
            for u in USUARIOS_INICIALES:
                db.session.add(Usuario(**u))
            db.session.commit()
    except Exception:
        db.session.rollback()
    _asegurar_superadmin()

def _asegurar_superadmin():
    """Si no hay ningún superadmin y está definida SUPERADMIN_PASSWORD, lo crea. Nunca modifica uno existente."""
    pw = os.environ.get("SUPERADMIN_PASSWORD", "").strip()
    if not pw or Usuario.query.filter_by(rol="superadmin").first():
        return
    login = os.environ.get("SUPERADMIN_USUARIO", "superadmin").strip() or "superadmin"
    if Usuario.query.filter_by(usuario=login).first():
        return   # ese nombre de usuario ya lo usa otra persona: no se toca
    try:
        db.session.add(Usuario(nombre="Super Admin", usuario=login, password=pw, rol="superadmin", negocio="ambos"))
        db.session.commit()
    except Exception:
        db.session.rollback()   # otro worker lo creó al mismo tiempo

with app.app_context():
    db.create_all()
    _migrar_columnas()
    _migrar_negocio()
    _seed_usuarios()

SYNC_KEY = os.environ.get("SYNC_KEY", "cpfr2024")

FOTOS_DIR = os.path.join(os.path.dirname(__file__), '..', 'fotos')
os.makedirs(FOTOS_DIR, exist_ok=True)
# ZIP del último cierre (disco del servidor; se pierde al reiniciar/publicar, sirve para re-descargar)
CIERRES_DIR = os.environ.get("CIERRES_DIR", os.path.join(os.path.dirname(__file__), '..', 'cierres'))

# Mundo Vegano — visita mensual (rurales + occidente ya separados en JS)
TIENDAS_RURALES = {
    "WM LIBERIA", "WM CIUDAD QUESADA", "WM PEREZ ZELEDON",
    "MXM LIBERIA", "MXM GUAPILES", "MXM LIMON", "MXM JACO",
    "MXM SAN RAMON", "MXM GRECIA",
    "A10 AUTO MERCADO HERRADURA", "A11 AUTO MERCADO TAMARINDO", "A12 AUTO MERCADO COCO",
}

# Rooster — visita mensual (rurales WM/MXM/MP/AM + Occidente MP)
TIENDAS_RURALES_ROOSTER = {
    # WM rurales
    "WM LA LIMA", "WM LIBERIA", "WM CIUDAD QUESADA", "WM PEREZ ZELEDON",
    # MXM rurales
    "MXM GUAPILES", "MXM LIMON", "MXM JACO", "MXM LIBERIA",
    # MP rurales (Más x Menos)
    "MP CIUDAD QUESADA", "MP LIMON", "MP LIBERIA", "MP JACO", "MP SAN ISIDRO",
    "MP PUNTARENAS", "MP TURRIALBA", "MP NICOYA", "MP GUAPILES", "MP CAÑAS",
    "MP QUEPOS", "MP RIO CLARO", "MP SANTA CRUZ", "MP OROTINA", "MP ESPARZA",
    "MP SIQUIRRES", "MP PURISCAL", "MP AGUAS ZARCAS", "MP PARRITA", "MP TAMARINDO",
    "MP TURRIALBA NORTE", "MP LA FORTUNA", "MP UPALA", "MP PLAYAS DEL COCO",
    "MP CARIARI", "MP UVITA",
    # MP Occidente (también mensual)
    "MP PALMARES", "MP GRECIA", "MP SAN RAMON", "MP ATENAS", "MP NARANJO",
    # MXM Occidente
    "MXM SAN RAMON", "MXM GRECIA",
    # AM rurales
    "A10 AUTO MERCADO HERRADURA", "A11 AUTO MERCADO TAMARINDO", "A12 AUTO MERCADO COCO",
}

# Cadenas conocidas. "patron" detecta la cadena por el nombre de la tienda (como viene en el
# inventario); "grupo" es la tarjeta del panel principal donde aparecen sus tiendas.
CADENAS_CATALOGO = [
    {"codigo": "WM",    "nombre": "Walmart",      "grupo": "walmart",     "patron": r"^WM\b|WALMART"},
    {"codigo": "MXM",   "nombre": "Más x Menos",  "grupo": "walmart",     "patron": r"^MXM\b"},
    {"codigo": "MP",    "nombre": "Maxi Palí",    "grupo": "walmart",     "patron": r"^MP\b|MAXI ?PALI"},
    {"codigo": "PALI",  "nombre": "Palí",         "grupo": "walmart",     "patron": r"^PALI\b"},
    {"codigo": "AM",    "nombre": "Auto Mercado", "grupo": "automercado", "patron": r"AUTO ?MERCADO|^A\d|^AM\b"},
    {"codigo": "VINDI", "nombre": "Vindi",        "grupo": "vindi",       "patron": r"VINDI|^C\d"},
]
_CAT_POR_NOMBRE = {c["nombre"].upper(): c["codigo"] for c in CADENAS_CATALOGO}

def _patron_cadena(codigo):
    c = next((c for c in CADENAS_CATALOGO if c["codigo"] == codigo), None)
    # cadena personalizada: la tienda empieza con el código
    return c["patron"] if c else r"^" + re.escape(codigo) + r"\b"

def cadena_de_tienda(tienda, extra=()):
    """Código de la cadena de una tienda, o '' si no se reconoce."""
    t = (tienda or "").strip().upper()
    for c in CADENAS_CATALOGO:
        if re.search(c["patron"], t):
            return c["codigo"]
    for cod in extra:
        if re.search(_patron_cadena(cod), t):
            return cod
    return ""

def parse_cadenas(valor):
    """Acepta lista o texto ('WM,MXM' o el formato viejo 'WM · MXM · Vindi') → lista de códigos."""
    if isinstance(valor, (list, tuple)):
        partes = valor
    else:
        partes = re.split(r"[,·]", valor or "")
    out = []
    for p in partes:
        p = str(p).strip().upper()
        if not p:
            continue
        p = _CAT_POR_NOMBRE.get(p, p)
        p = re.sub(r"[^A-Z0-9 ]", "", p).strip()[:20]
        if p and p not in out:
            out.append(p)
    return out

def nombre_cadena(codigo):
    c = next((c for c in CADENAS_CATALOGO if c["codigo"] == codigo), None)
    return c["nombre"] if c else codigo.title()

def _lineas(txt):
    return {t.strip().upper() for t in (txt or "").splitlines() if t.strip()}

def filtro_tiendas(negocio):
    """Devuelve f(tienda)->bool: False si la tienda fue quitada o su cadena está desactivada."""
    n = Negocio.query.get(negocio)
    if not n:
        return lambda t: True
    excl = _lineas(n.excluidas)
    activas = set(parse_cadenas(n.cadenas))
    extra = [c for c in activas if c not in {x["codigo"] for x in CADENAS_CATALOGO}]
    def ok(t):
        if (t or "").strip().upper() in excl:
            return False
        cad = cadena_de_tienda(t, extra)
        return not activas or not cad or cad in activas
    return ok

def rurales_maestro():
    """'Memoria' de tiendas rurales: las de fábrica + las marcadas en cualquier negocio."""
    todas = set(TIENDAS_RURALES) | set(TIENDAS_RURALES_ROOSTER)
    for n in Negocio.query.all():
        todas |= _lineas(n.rurales)
    return {t.upper() for t in todas}

# Datos iniciales: solo se usan para sembrar la tabla negocio la primera vez
NEGOCIOS_INICIALES = [
    {"slug": "mundo_vegano", "nombre": "Mundo Vegano", "cadenas": "WM,MXM,PALI,AM",
     "rurales": TIENDAS_RURALES, "logo": "mv_logo.png", "orden": 1},
    {"slug": "rooster", "nombre": "Rooster", "cadenas": "WM,MXM,MP,AM",
     "rurales": TIENDAS_RURALES_ROOSTER, "logo": "rooster_logo.jpg", "orden": 2},
]

def _seed_negocios():
    try:
        _seed_negocios_()
    except Exception:
        db.session.rollback()   # otro worker lo sembró al mismo tiempo

def _seed_negocios_():
    if Negocio.query.count():
        return
    static_dir = os.path.join(os.path.dirname(__file__), "static")
    for n in NEGOCIOS_INICIALES:
        logo = None
        ruta = os.path.join(static_dir, n["logo"])
        if os.path.exists(ruta):
            with open(ruta, "rb") as f:
                mime = "jpeg" if ruta.lower().endswith((".jpg", ".jpeg")) else "png"
                logo = f"data:image/{mime};base64," + base64.b64encode(f.read()).decode()
        db.session.add(Negocio(slug=n["slug"], nombre=n["nombre"], cadenas=n["cadenas"],
                               rurales="\n".join(sorted(n["rurales"])), logo_b64=logo,
                               activo=True, orden=n["orden"],
                               creado_en=datetime.now().strftime("%Y-%m-%d %H:%M")))
    db.session.commit()

with app.app_context():
    _seed_negocios()

def negocios_de_usuario(u):
    """Lista de negocios a los que entra el usuario ("ambos" = todos)."""
    v = (getattr(u, "negocio", None) or "mundo_vegano").strip()
    if v == "ambos":
        return list(slugs_negocios(solo_activos=True))
    return [x for x in v.split(",") if x]

def _usuario_en_negocio(u, slug):
    return (u.negocio or "") == "ambos" or slug in (u.negocio or "").split(",")

def slugs_negocios(solo_activos=False):
    q = Negocio.query
    if solo_activos:
        q = q.filter_by(activo=True)
    return [n.slug for n in q.all()]

def rurales_de(negocio):
    n = Negocio.query.get(negocio)
    if not n or not n.rurales:
        return set()
    return {t.strip().upper() for t in n.rurales.splitlines() if t.strip()}

def semana_captura(negocio="mundo_vegano"):
    """Semana en la que guardan los displays: la del calendario en hora de Costa Rica (UTC-6).
    Cambia sola el lunes a las 00:00; no depende de cierres ni de lo que se elija en el panel."""
    return (datetime.utcnow() - timedelta(hours=6)).strftime("%G-S%V")

def semana_actual(negocio="mundo_vegano"):
    """Semana del PANEL (la que se revisa, valida y cierra): la elegida con 'Cambiar semana',
    o la del calendario si no se eligió ninguna."""
    clave = f"{negocio}:semana_override"
    cfg = Config.query.get(clave)
    if cfg and cfg.valor:
        return cfg.valor
    # fallback a la clave legacy (solo mundo_vegano)
    if negocio == "mundo_vegano":
        cfg_leg = Config.query.get("semana_override")
        if cfg_leg and cfg_leg.valor:
            return cfg_leg.valor
    return semana_captura(negocio)

MESES_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]

def rango_de_codigo(semana_code):
    """Convierte '2026-S30' al rango de fechas legible y para nombre de archivo."""
    try:
        parts = semana_code.split("-S")
        year, week = int(parts[0]), int(parts[1])
        lunes  = datetime.fromisocalendar(year, week, 1)
        domingo = lunes + timedelta(days=6)
    except Exception:
        now = datetime.now()
        lunes = now - timedelta(days=now.weekday())
        domingo = lunes + timedelta(days=6)
    mes_lunes   = MESES_ES[lunes.month - 1]
    mes_domingo = MESES_ES[domingo.month - 1]
    if lunes.month == domingo.month:
        archivo = f"{lunes.day:02d}-{domingo.day:02d}_{mes_domingo}_{domingo.year}"
        legible = f"{lunes.day} – {domingo.day} {mes_domingo} {domingo.year}"
    else:
        archivo = f"{lunes.day:02d}_{mes_lunes}-{domingo.day:02d}_{mes_domingo}_{domingo.year}"
        legible = f"{lunes.day} {mes_lunes} – {domingo.day} {mes_domingo} {domingo.year}"
    return archivo, legible

def rango_semana_actual():
    return rango_de_codigo(semana_actual())

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/ping")
def ping():
    return jsonify({"ok": True})

@app.route("/api/login", methods=["POST"])
def login():
    data     = request.json
    usuario  = data.get("usuario", "").strip()
    password = data.get("password", "").strip()
    u = Usuario.query.filter_by(usuario=usuario, password=password).first()
    if u:
        negocio = getattr(u, "negocio", None) or "mundo_vegano"
        negs = negocios_de_usuario(u)
        return jsonify({
            "ok": True, "id": u.id, "nombre": u.nombre, "rol": u.rol,
            "negocio": negocio if negocio == "ambos" else (negs[0] if negs else "mundo_vegano"),
            "negocios": negs,
            # super-admin, admin general y displays de varios negocios eligen; los demás entran directo
            "selector_negocio": negocio == "ambos" or len(negs) > 1,
        })
    return jsonify({"ok": False, "error": "Usuario o contraseña incorrectos"}), 401

# ── Usuarios en línea (latido cada 60 s desde la app; en memoria: gunicorn corre con 1 worker) ──
_PRESENCIA = {}            # usuario_id -> {nombre, rol, negocio, pantalla, visto}
_PRESENCIA_LOCK = threading.Lock()
PRESENCIA_SEG = 150        # sin latido en 2.5 min = desconectado

@app.route("/api/presencia", methods=["POST"])
def presencia_latido():
    d = request.get_json(silent=True) or {}
    try:
        uid = int(d.get("id"))
    except (TypeError, ValueError):
        return jsonify({"ok": False}), 400
    u = db.session.get(Usuario, uid)
    if not u:
        return jsonify({"ok": False}), 404
    with _PRESENCIA_LOCK:
        if d.get("salir"):
            _PRESENCIA.pop(uid, None)
        else:
            _PRESENCIA[uid] = {"nombre": u.nombre, "usuario": u.usuario, "rol": u.rol,
                               "negocio": (d.get("negocio") or "")[:50], "pantalla": (d.get("pantalla") or "")[:30],
                               "visto": time.time()}
    return jsonify({"ok": True})

@app.route("/api/presencia")
def presencia_lista():
    if request.args.get("solicitante_rol") != "superadmin":
        return jsonify({"error": "No autorizado"}), 403
    ahora = time.time()
    with _PRESENCIA_LOCK:
        for k in [k for k, v in _PRESENCIA.items() if ahora - v["visto"] > PRESENCIA_SEG]:
            _PRESENCIA.pop(k)
        lista = [dict(v, id=k, hace_seg=int(ahora - v["visto"])) for k, v in _PRESENCIA.items()]
    lista.sort(key=lambda x: (x["rol"] != "display", x["nombre"].lower()))
    return jsonify({"en_linea": len(lista), "usuarios": lista})

# ── FOTOS QUE EL CELULAR NO PUEDE ABRIR: el servidor las convierte (HEIC/HEIF, muy pesadas, etc.) ──
try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except Exception:   # sin la librería igual convierte JPG/PNG/WebP pesados
    pass

_LOG_FOTOS_CLAVE = "log_fotos"

def _log_foto(d):
    """Guarda los últimos 200 eventos de fotos con problema (para diagnosticar sin pedir capturas)."""
    app.logger.warning("FOTO %s", json.dumps(d, ensure_ascii=False)[:500])
    try:
        cfg = db.session.get(Config, _LOG_FOTOS_CLAVE)
        lista = json.loads(cfg.valor) if cfg and cfg.valor else []
        lista = (lista + [d])[-200:]
        if not cfg:
            cfg = Config(clave=_LOG_FOTOS_CLAVE, valor="")
            db.session.add(cfg)
        cfg.valor = json.dumps(lista, ensure_ascii=False)
        db.session.commit()
    except Exception:
        db.session.rollback()

@app.route("/api/foto/convertir", methods=["POST"])
def foto_convertir():
    f = request.files.get("foto")
    info = {k: (request.form.get(k) or "")[:200] for k in ("usuario", "negocio", "tienda", "ua", "motivo")}
    if not f:
        return jsonify({"ok": False, "error": "No llegó la foto"}), 400
    raw = f.read()
    info.update({"fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "etapa": "servidor",
                 "tipo": f.mimetype or "", "nombre": (f.filename or "")[-40:], "mb": round(len(raw) / 1048576, 2)})
    try:
        from PIL import Image, ImageOps
        Image.MAX_IMAGE_PIXELS = 250_000_000          # hasta ~250 MP (celulares de 200 MP)
        img = Image.open(io.BytesIO(raw))
        info["formato"] = img.format or ""
        info["px"] = f"{img.width}x{img.height}"
        img.draft("RGB", (2400, 2400))                 # JPG grande: decodifica ya reducido (menos memoria)
        img = ImageOps.exif_transpose(img).convert("RGB")
        lado = max(400, min(int(request.form.get("max") or 1200), 2000))
        img.thumbnail((lado, lado))
        out = io.BytesIO()
        img.save(out, "JPEG", quality=82, optimize=True)
        del raw, img
        info["resultado"] = "convertida"
        _log_foto(info)
        return jsonify({"ok": True, "foto_b64": base64.b64encode(out.getvalue()).decode()})
    except Exception as e:
        info["resultado"] = f"error: {type(e).__name__}: {str(e)[:120]}"
        _log_foto(info)
        return jsonify({"ok": False, "error": "El servidor tampoco pudo leer esa foto", "detalle": info["resultado"]}), 422

@app.route("/api/foto/log", methods=["POST"])
def foto_log():
    d = request.get_json(silent=True) or {}
    d = {k: str(v)[:200] for k, v in d.items() if k in ("usuario", "negocio", "tienda", "ua", "motivo", "tipo", "mb", "nombre", "etapa", "resultado")}
    d["fecha"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    _log_foto(d)
    return jsonify({"ok": True})

@app.route("/api/foto/log")
def foto_log_ver():
    if request.args.get("solicitante_rol") != "superadmin":
        return jsonify({"error": "No autorizado"}), 403
    cfg = db.session.get(Config, _LOG_FOTOS_CLAVE)
    return jsonify(json.loads(cfg.valor) if cfg and cfg.valor else [])

@app.route("/api/semana-activa")
def get_semana_activa():
    negocio = request.args.get("negocio", "mundo_vegano")
    codigo = semana_actual(negocio)
    _, legible = rango_de_codigo(codigo)
    ya_cerrada = Diferencia.query.filter_by(semana=codigo, negocio=negocio).first() is not None
    cap = semana_captura(negocio)
    return jsonify({"semana_activa": codigo, "semana_legible": legible, "ya_cerrada": ya_cerrada,
                    "semana_captura": cap, "semana_captura_legible": rango_de_codigo(cap)[1]})

@app.route("/api/set-semana", methods=["POST"])
def set_semana():
    data = request.json or {}
    if data.get("rol") not in ROLES_ADMIN:
        return jsonify({"ok": False, "error": "Solo un administrador puede hacer esto"}), 403
    semana  = (data.get("semana") or "").strip()
    negocio = (data.get("negocio") or "mundo_vegano").strip()
    clave   = f"{negocio}:semana_override"
    cfg = Config.query.get(clave)
    if semana:
        if not cfg:
            cfg = Config(clave=clave, valor=semana)
            db.session.add(cfg)
        else:
            cfg.valor = semana
    else:
        if cfg:
            cfg.valor = ""
    db.session.commit()
    return jsonify({"ok": True, "semana_activa": semana_actual(negocio)})

# ─────────────────────────────────────────────
# NEGOCIOS (clientes) — alta/edición desde la app
# ─────────────────────────────────────────────
def _es_superadmin(data):
    return (data or {}).get("solicitante_rol", "") == "superadmin"

def _negocio_json(n, admin=False):
    cods = parse_cadenas(n.cadenas)
    d = {"slug": n.slug, "nombre": n.nombre, "cadenas": cods,
         "cadenas_txt": " · ".join(cods),
         "logo": f"/api/negocios/{n.slug}/logo" if n.logo_b64 else "",
         "activo": bool(n.activo), "orden": n.orden or 0}
    if admin:
        d["rurales"] = n.rurales or ""
        d["excluidas"] = n.excluidas or ""
        d["prompt"] = prompt_de(n.slug)
        d["prompt_original"] = _prompt_por_defecto(n.slug, n.nombre)
        d["creado_en"] = n.creado_en or ""
    return d

def _texto_tiendas(valor):
    """Lista o texto → texto normalizado (una tienda por línea, mayúsculas, sin duplicados)."""
    items = valor if isinstance(valor, (list, tuple)) else (valor or "").splitlines()
    vistos = []
    for t in items:
        t = str(t).strip().upper()
        if t and t not in vistos:
            vistos.append(t)
    return "\n".join(vistos)

def _resumen_tiendas(tiendas, extra=()):
    """Agrupa tiendas por cadena y sugiere rurales según la memoria de la app."""
    maestro = rurales_maestro()
    por_cadena = {}
    for t in tiendas:
        cod = cadena_de_tienda(t, extra) or "OTRA"
        por_cadena[cod] = por_cadena.get(cod, 0) + 1
    detectadas = [{"codigo": c, "nombre": nombre_cadena(c) if c != "OTRA" else "Sin cadena", "tiendas": k}
                  for c, k in sorted(por_cadena.items(), key=lambda x: -x[1])]
    return {
        "tiendas": sorted(tiendas),
        "cadenas_detectadas": detectadas,
        "rurales_sugeridas": sorted(t for t in tiendas if t.upper() in maestro),
    }

@app.route("/api/cadenas")
def catalogo_cadenas():
    return jsonify(CADENAS_CATALOGO)

@app.route("/api/negocios/analizar-inventario", methods=["POST"])
def analizar_inventario_negocio():
    """Lee el/los Excel de inventario SIN guardar nada: devuelve tiendas, cadenas y rurales sugeridas."""
    if request.form.get("solicitante_rol") != "superadmin":
        return jsonify({"error": "No autorizado"}), 403
    archivo_wm = request.files.get("archivo_wm")
    archivo_am = request.files.get("archivo_am")
    if not archivo_wm and not archivo_am:
        return jsonify({"error": "Subí al menos un Excel de inventario"}), 400
    try:
        registros = []
        if archivo_wm:
            registros += _leer_excel_walmart(archivo_wm)
        if archivo_am:
            registros += _leer_excel_automercado(archivo_am)
    except Exception as e:
        return jsonify({"error": f"No se pudo leer el Excel: {e}"}), 400
    tiendas = {r["tienda"].strip().upper() for r in registros if r.get("tienda")}
    res = _resumen_tiendas(tiendas)
    res["productos"] = len({r["producto"] for r in registros})
    res["registros"] = len(registros)
    return jsonify(res)

def _formato_excel(archivo):
    """'wm' (Retail Link: Store Name + Signing Desc), 'am' (Auto Mercado: NombrePDX) o None."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(archivo.read()), read_only=True, data_only=True)
        if _buscar_encabezados(wb, ("Store Name", "Signing Desc"))[0]:
            return "wm"
        if _buscar_encabezados(wb, ("NombrePDX", "NombreProducto"))[0]:
            return "am"
        return None
    except Exception:
        return None
    finally:
        archivo.seek(0)

@app.route("/api/negocios/detectar-excel", methods=["POST"])
def detectar_excel():
    """Recibe varios Excel (archivos sueltos o una carpeta) y dice cuál es de Walmart y cuál de Auto Mercado."""
    if request.form.get("solicitante_rol") != "superadmin":
        return jsonify({"error": "No autorizado"}), 403
    res = []
    for i, f in enumerate(request.files.getlist("archivos")):
        res.append({"indice": i, "nombre": f.filename, "formato": _formato_excel(f)})
    return jsonify(res)

@app.route("/api/negocios/<slug>/tiendas")
def tiendas_config_negocio(slug):
    """Tiendas del inventario actual del negocio, para editar semanales/rurales/quitadas."""
    if request.args.get("solicitante_rol") != "superadmin":
        return jsonify({"error": "No autorizado"}), 403
    n = Negocio.query.get_or_404(slug)
    rows = db.session.query(Inventario.tienda).filter(Inventario.negocio == slug).distinct().all()
    tiendas = {r.tienda.strip().upper() for r in rows if r.tienda}
    extra = [c for c in parse_cadenas(n.cadenas) if c not in {x["codigo"] for x in CADENAS_CATALOGO}]
    res = _resumen_tiendas(tiendas, extra)
    res["rurales"] = sorted(_lineas(n.rurales))
    res["excluidas"] = sorted(_lineas(n.excluidas))
    return jsonify(res)

def _slugify(texto):
    import unicodedata, re as _re
    t = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode().lower()
    return _re.sub(r"[^a-z0-9]+", "_", t).strip("_")[:30].strip("_")

def _validar_prompt(texto):
    texto = (texto or "").strip()
    if not texto:
        return None   # vacío = usar el prompt por defecto
    if "{producto}" not in texto:
        raise ValueError("El prompt debe incluir {producto} donde va el nombre del producto a validar")
    if "ESTADO:" not in texto:
        raise ValueError("El prompt debe terminar pidiendo el formato ESTADO / CONFIANZA / MOTIVO")
    if len(texto) > 20000:
        raise ValueError("El prompt es demasiado largo (máximo 20.000 caracteres)")
    return texto

def _validar_logo(logo):
    """Acepta data URL png/jpeg/webp de hasta ~1.5 MB."""
    if not logo:
        return None
    if not logo.startswith(("data:image/png;base64,", "data:image/jpeg;base64,", "data:image/webp;base64,")):
        raise ValueError("El logo debe ser una imagen PNG, JPG o WEBP")
    if len(logo) > 2_000_000:
        raise ValueError("El logo es muy pesado (máximo ~1.5 MB)")
    return logo

@app.route("/api/negocios")
def listar_negocios():
    """Lista para el selector de negocio: solo activos."""
    rows = Negocio.query.filter_by(activo=True).order_by(Negocio.orden, Negocio.nombre).all()
    return jsonify([_negocio_json(n) for n in rows])

@app.route("/api/negocios/admin")
def listar_negocios_admin():
    if request.args.get("solicitante_rol") != "superadmin":
        return jsonify({"error": "No autorizado"}), 403
    rows = Negocio.query.order_by(Negocio.orden, Negocio.nombre).all()
    return jsonify([_negocio_json(n, admin=True) for n in rows])

@app.route("/api/negocios/<slug>/logo")
def logo_negocio(slug):
    n = Negocio.query.get_or_404(slug)
    if not n.logo_b64:
        return "", 404
    cabecera, datos = n.logo_b64.split(",", 1)
    mime = cabecera[5:].split(";")[0]
    resp = send_file(io.BytesIO(base64.b64decode(datos)), mimetype=mime)
    resp.headers["Cache-Control"] = "no-cache"
    return resp

@app.route("/api/negocios/prompt-base")
def prompt_base_negocio():
    """Prompt inicial sugerido para un negocio nuevo."""
    if request.args.get("solicitante_rol") != "superadmin":
        return jsonify({"error": "No autorizado"}), 403
    return jsonify({"prompt": _prompt_por_defecto("", request.args.get("nombre", ""))})

@app.route("/api/negocios/<slug>/rurales")
def rurales_negocio(slug):
    return jsonify(sorted(rurales_de(slug)))

@app.route("/api/negocios", methods=["POST"])
def crear_negocio():
    data = request.get_json(force=True) or {}
    if not _es_superadmin(data):
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    nombre = (data.get("nombre") or "").strip()
    if not nombre:
        return jsonify({"ok": False, "error": "El nombre es requerido"}), 400
    if len(nombre) > 100:
        return jsonify({"ok": False, "error": "El nombre es demasiado largo (máximo 100 caracteres)"}), 400
    slug = _slugify(nombre)
    if not slug or slug == "ambos":
        return jsonify({"ok": False, "error": "Nombre inválido"}), 400
    if Negocio.query.get(slug):
        return jsonify({"ok": False, "error": "Ya existe un negocio con ese nombre"}), 409
    try:
        logo = _validar_logo(data.get("logo"))
        prompt = _validar_prompt(data.get("prompt"))
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    orden = (db.session.query(db.func.max(Negocio.orden)).scalar() or 0) + 1
    n = Negocio(slug=slug, nombre=nombre, cadenas=",".join(parse_cadenas(data.get("cadenas")))[:200],
                rurales=_texto_tiendas(data.get("rurales")), excluidas=_texto_tiendas(data.get("excluidas")),
                logo_b64=logo, prompt=prompt, activo=True,
                orden=orden, creado_en=datetime.now().strftime("%Y-%m-%d %H:%M"))
    db.session.add(n)
    db.session.commit()
    return jsonify({"ok": True, "negocio": _negocio_json(n, admin=True)})

@app.route("/api/negocios/<slug>", methods=["PUT"])
def editar_negocio(slug):
    data = request.get_json(force=True) or {}
    if not _es_superadmin(data):
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    n = Negocio.query.get_or_404(slug)
    if "nombre" in data:
        nombre = (data.get("nombre") or "").strip()
        if not nombre:
            return jsonify({"ok": False, "error": "El nombre es requerido"}), 400
        if len(nombre) > 100:
            return jsonify({"ok": False, "error": "El nombre es demasiado largo (máximo 100 caracteres)"}), 400
        n.nombre = nombre
    if "cadenas" in data:
        n.cadenas = ",".join(parse_cadenas(data.get("cadenas")))[:200]
    if "rurales" in data:
        n.rurales = _texto_tiendas(data.get("rurales"))
    if "excluidas" in data:
        n.excluidas = _texto_tiendas(data.get("excluidas"))
    if "prompt" in data:
        try:
            p = _validar_prompt(data.get("prompt"))
        except ValueError as e:
            return jsonify({"ok": False, "error": str(e)}), 400
        # si es igual al de fábrica no se guarda copia: así hereda mejoras futuras
        n.prompt = None if p and p == _prompt_por_defecto(n.slug, n.nombre).strip() else p
    if "activo" in data:
        n.activo = bool(data.get("activo"))
    if data.get("logo"):
        try:
            n.logo_b64 = _validar_logo(data.get("logo"))
        except ValueError as e:
            return jsonify({"ok": False, "error": str(e)}), 400
    db.session.commit()
    return jsonify({"ok": True, "negocio": _negocio_json(n, admin=True)})

@app.route("/api/negocios/<slug>", methods=["DELETE"])
def eliminar_negocio(slug):
    """Borra el negocio y TODOS sus datos (reportes, inventario, historial, novedades, usuarios, config)."""
    data = request.get_json(force=True, silent=True) or {}
    if not _es_superadmin(data):
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    n = Negocio.query.get_or_404(slug)
    if (data.get("confirmar") or "").strip().lower() != n.nombre.strip().lower():
        return jsonify({"ok": False, "error": "Para eliminar, escribí el nombre exacto del negocio"}), 400
    if Negocio.query.count() <= 1:
        return jsonify({"ok": False, "error": "No se puede eliminar el único negocio"}), 400
    ValidacionIA.query.filter_by(negocio=slug).delete()
    for modelo in (Reporte, Inventario, Diferencia, Liquidacion, Usuario):
        modelo.query.filter_by(negocio=slug).delete()
    for u in Usuario.query.filter(Usuario.negocio.like(f"%{slug}%")).all():   # displays de varios negocios
        quedan = [x for x in u.negocio.split(",") if x and x != slug]
        if quedan: u.negocio = ",".join(quedan)
        elif u.negocio != "ambos": db.session.delete(u)
    Config.query.filter(Config.clave.like(f"{slug}:%")).delete(synchronize_session=False)
    db.session.delete(n)
    db.session.commit()
    try:
        for f in os.listdir(CIERRES_DIR):
            if f.startswith(f"{slug}__"):
                os.remove(os.path.join(CIERRES_DIR, f))
    except FileNotFoundError:
        pass
    return jsonify({"ok": True})

@app.route("/api/semanas")
def listar_semanas():
    # El calendario de semanas es compartido; la semana activa es por negocio
    negocio = request.args.get("negocio", "mundo_vegano")
    activa  = semana_actual(negocio)
    rows = SemanaConfig.query.order_by(SemanaConfig.id.desc()).all()
    return jsonify([{
        "id": s.id, "codigo": s.codigo, "nombre": s.nombre,
        "fecha_ini": s.fecha_ini, "fecha_fin": s.fecha_fin,
        "activa": s.codigo == activa, "creada_en": s.creada_en
    } for s in rows])

@app.route("/api/semanas", methods=["POST"])
def crear_semana():
    data = request.json or {}
    codigo   = (data.get("codigo") or "").strip()
    nombre   = (data.get("nombre") or "").strip()
    fecha_ini = (data.get("fecha_ini") or "").strip()
    fecha_fin = (data.get("fecha_fin") or "").strip()
    if not codigo or not nombre:
        return jsonify({"error": "Código y nombre son requeridos"}), 400
    if SemanaConfig.query.filter_by(codigo=codigo).first():
        return jsonify({"error": "Ya existe una semana con ese código"}), 409
    from datetime import datetime
    s = SemanaConfig(codigo=codigo, nombre=nombre, fecha_ini=fecha_ini,
                     fecha_fin=fecha_fin, activa=False,
                     creada_en=datetime.now().strftime("%Y-%m-%d %H:%M"))
    db.session.add(s)
    db.session.commit()
    return jsonify({"ok": True, "id": s.id})

@app.route("/api/semanas/<int:sid>/activar", methods=["POST"])
def activar_semana(sid):
    if (request.get_json(silent=True) or {}).get("rol") not in ROLES_ADMIN:
        return jsonify({"ok": False, "error": "Solo un administrador puede hacer esto"}), 403
    s = SemanaConfig.query.get_or_404(sid)
    negocio = (request.get_json(silent=True) or {}).get("negocio", "mundo_vegano")
    # Actualizar {negocio}:semana_override para que todo el sistema de ese negocio lo use
    clave = f"{negocio}:semana_override"
    cfg = Config.query.get(clave)
    if not cfg:
        cfg = Config(clave=clave, valor=s.codigo)
        db.session.add(cfg)
    else:
        cfg.valor = s.codigo
    db.session.commit()
    return jsonify({"ok": True, "semana_activa": s.codigo, "nombre": s.nombre})

@app.route("/api/semanas/<int:sid>", methods=["DELETE"])
def eliminar_semana(sid):
    s = SemanaConfig.query.get_or_404(sid)
    if any(s.codigo == semana_actual(n) for n in slugs_negocios()):
        return jsonify({"error": "No se puede eliminar una semana activa"}), 400
    db.session.delete(s)
    db.session.commit()
    return jsonify({"ok": True})

# Roles con permisos de administración dentro de un negocio.
#   supervisor = admin de UN negocio · admin = admin general (todos los negocios, sin Sistema global)
#   superadmin = todo, incluido Sistema (negocios, prompts, consumo)
ROLES_ADMIN = ("supervisor", "admin", "superadmin")

def _es_supervisor(data):
    rol = (data or {}).get("solicitante_rol", "")
    return rol in ROLES_ADMIN

@app.route("/api/usuarios")
def listar_usuarios():
    negocio = request.args.get("negocio", "mundo_vegano")
    rows = (Usuario.query.filter(db.or_(Usuario.negocio.like(f"%{negocio}%"), Usuario.rol == "admin"))
            .order_by(Usuario.rol.desc(), Usuario.nombre).all())
    rows = [u for u in rows if u.rol == "admin" or _usuario_en_negocio(u, negocio)]
    return jsonify([{"id": u.id, "nombre": u.nombre, "usuario": u.usuario, "password": u.password,
                     "rol": u.rol, "negocio": u.negocio,
                     "negocios": negocios_de_usuario(u) if u.rol != "admin" else []} for u in rows])

@app.route("/api/usuarios/<int:usuario_id>/negocios", methods=["POST"])
def negocios_usuario(usuario_id):
    """Superadmin / admin general: a qué negocios entra un display (una sola cuenta para varios)."""
    data = request.json or {}
    if data.get("solicitante_rol") not in ("superadmin", "admin"):
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    u = Usuario.query.get_or_404(usuario_id)
    if u.rol in ("superadmin", "admin"):
        return jsonify({"ok": False, "error": "Ese usuario ya entra a todos los negocios"}), 400
    if u.rol != "display" and data.get("solicitante_rol") != "superadmin":
        return jsonify({"ok": False, "error": "Solo el superadmin cambia los negocios de administradores"}), 403
    validos = set(slugs_negocios())
    negs = [x for x in dict.fromkeys(data.get("negocios") or []) if x in validos]
    if not negs:
        return jsonify({"ok": False, "error": "Elegí al menos un negocio"}), 400
    u.negocio = ",".join(negs)
    db.session.commit()
    return jsonify({"ok": True, "negocios": negs})

@app.route("/api/usuarios", methods=["POST"])
def crear_usuario():
    data = request.json
    if not _es_supervisor(data):
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    nombre   = data.get("nombre", "").strip()
    usuario  = data.get("usuario", "").strip()
    password = data.get("password", "").strip()
    if not (nombre and usuario and password):
        return jsonify({"ok": False, "error": "Faltan datos"}), 400
    if Usuario.query.filter_by(usuario=usuario).first():
        return jsonify({"ok": False, "error": "Ese usuario ya existe"}), 400
    rol_nuevo = data.get("rol", "display")
    # Solo el superadmin crea administradores principales (supervisor)
    permitidos = ("display", "subadmin", "supervisor", "admin") if _es_superadmin(data) else ("display", "subadmin")
    if rol_nuevo not in permitidos:
        return jsonify({"ok": False, "error": "No autorizado para crear ese tipo de usuario"}), 403
    negocio = data.get("negocio", "mundo_vegano")
    if rol_nuevo == "admin":
        negocio = "ambos"   # admin general: entra a todos los negocios
    elif isinstance(data.get("negocios"), list) and _es_superadmin(data):
        validos = set(slugs_negocios())
        negs = [x for x in dict.fromkeys(data["negocios"]) if x in validos]
        if not negs:
            return jsonify({"ok": False, "error": "Elegí al menos un negocio"}), 400
        negocio = ",".join(negs)
    elif negocio not in slugs_negocios():
        return jsonify({"ok": False, "error": "Negocio inválido"}), 400
    u = Usuario(nombre=nombre, usuario=usuario, password=password, rol=rol_nuevo, negocio=negocio)
    db.session.add(u)
    db.session.commit()
    return jsonify({"ok": True, "id": u.id})

@app.route("/api/usuarios/todos")
def usuarios_todos():
    """Superadmin: todos los usuarios de la app (Sistema → Usuarios)."""
    if request.args.get("solicitante_rol") != "superadmin":
        return jsonify({"error": "No autorizado"}), 403
    rows = Usuario.query.order_by(Usuario.nombre).all()
    return jsonify([{"id": u.id, "nombre": u.nombre, "usuario": u.usuario, "password": u.password, "rol": u.rol,
                     "todos": (u.negocio or "") == "ambos",
                     "negocios": [] if (u.negocio or "") == "ambos" else negocios_de_usuario(u)} for u in rows])

@app.route("/api/usuarios/<int:usuario_id>", methods=["PUT"])
def editar_usuario(usuario_id):
    """Superadmin: cambia nombre, tipo (rol) y negocios de un usuario."""
    data = request.json or {}
    if not _es_superadmin(data):
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    u = Usuario.query.get_or_404(usuario_id)
    if u.rol == "superadmin":
        return jsonify({"ok": False, "error": "El superadmin no se modifica desde aquí"}), 400
    nombre = (data.get("nombre") or "").strip()
    if nombre:
        u.nombre = nombre
    rol = data.get("rol", u.rol)
    if rol not in ("admin", "supervisor", "subadmin", "display"):
        return jsonify({"ok": False, "error": "Tipo de usuario inválido"}), 400
    if rol == "admin":
        u.negocio = "ambos"
    elif "negocios" in data or u.negocio == "ambos":
        validos = set(slugs_negocios())
        negs = [x for x in dict.fromkeys(data.get("negocios") or []) if x in validos]
        if not negs:
            return jsonify({"ok": False, "error": "Elegí al menos un negocio"}), 400
        u.negocio = ",".join(negs)
    u.rol = rol
    db.session.commit()
    return jsonify({"ok": True})

@app.route("/api/usuarios/<int:usuario_id>", methods=["DELETE"])
def eliminar_usuario(usuario_id):
    if not _es_supervisor(request.json):
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    u = Usuario.query.get_or_404(usuario_id)
    if u.rol == "superadmin":
        return jsonify({"ok": False, "error": "No se puede eliminar al superadmin"}), 400
    if u.rol in ("supervisor", "admin") and not _es_superadmin(request.json):
        return jsonify({"ok": False, "error": "Solo el superadmin puede eliminar administradores"}), 403
    # display en varios negocios: desde un negocio solo se le quita ESE acceso
    neg = (request.json or {}).get("negocio")
    negs = [x for x in (u.negocio or "").split(",") if x]
    if u.rol == "display" and neg in negs and len(negs) > 1:
        u.negocio = ",".join(x for x in negs if x != neg)
        db.session.commit()
        return jsonify({"ok": True, "quitado_de": neg})
    db.session.delete(u)
    db.session.commit()
    return jsonify({"ok": True})

@app.route("/api/usuarios/<int:usuario_id>/password", methods=["POST"])
def cambiar_password(usuario_id):
    data = request.json or {}
    if not _es_supervisor(data):
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    nueva = data.get("password", "").strip()
    if not nueva:
        return jsonify({"ok": False, "error": "Contraseña vacía"}), 400
    u = Usuario.query.get_or_404(usuario_id)
    if u.rol in ("supervisor", "admin", "superadmin") and not _es_superadmin(data):
        return jsonify({"ok": False, "error": "Solo el superadmin puede cambiar esa contraseña"}), 403
    u.password = nueva
    db.session.commit()
    return jsonify({"ok": True})

CLAVES_OCULTAR = ("ocultar_automercado", "ocultar_walmart", "ocultar_vindi")

def _ocultar(negocio, clave):
    """Lee {negocio}:{clave}; Mundo Vegano cae a la clave legacy sin prefijo."""
    cfg = Config.query.get(f"{negocio}:{clave}")
    if cfg is None and negocio == "mundo_vegano":
        cfg = Config.query.get(clave)
    return bool(cfg and cfg.valor == "1")

@app.route("/api/tiendas")
def tiendas():
    negocio = request.args.get("negocio", "mundo_vegano")
    rows = db.session.query(Inventario.tienda).filter(Inventario.negocio == negocio).distinct().order_by(Inventario.tienda).all()
    habilitada = filtro_tiendas(negocio)
    tiendas_list = [r.tienda for r in rows if habilitada(r.tienda)]
    if _ocultar(negocio, "ocultar_automercado"):
        tiendas_list = [t for t in tiendas_list if "AUTO" not in t.upper() and not t.upper().startswith("AM ")]
    if _ocultar(negocio, "ocultar_walmart"):
        tiendas_list = [t for t in tiendas_list if not t.upper().startswith("WM ") and "WALMART" not in t.upper()]
    if _ocultar(negocio, "ocultar_vindi"):
        tiendas_list = [t for t in tiendas_list if "VINDI" not in t.upper()]
    return jsonify(tiendas_list)

@app.route("/api/config/ocultar-cadena", methods=["POST"])
def toggle_ocultar_cadena():
    data = request.get_json(force=True)
    clave = data.get("clave")
    if clave not in CLAVES_OCULTAR:
        return jsonify({"error": "clave inválida"}), 400
    negocio = data.get("negocio", "mundo_vegano")
    valor = "1" if data.get("ocultar") else "0"
    clave_neg = f"{negocio}:{clave}"
    cfg = Config.query.get(clave_neg)
    if cfg:
        cfg.valor = valor
    else:
        db.session.add(Config(clave=clave_neg, valor=valor))
    db.session.commit()
    return jsonify({"ok": True, "ocultar": valor == "1"})

@app.route("/api/config/ocultar-cadena", methods=["GET"])
def get_ocultar_cadena():
    clave   = request.args.get("clave", "ocultar_automercado")
    negocio = request.args.get("negocio", "mundo_vegano")
    if clave not in CLAVES_OCULTAR:
        return jsonify({"error": "clave inválida"}), 400
    return jsonify({"ocultar": _ocultar(negocio, clave)})

@app.route("/api/reset-semana-actual", methods=["POST"])
def reset_semana_actual():
    data    = request.get_json(silent=True) or {}
    if data.get("rol") not in ROLES_ADMIN:
        return jsonify({"error": "No autorizado"}), 403
    negocio = (data.get("negocio") or "").strip()
    if not Negocio.query.get(negocio):
        return jsonify({"error": "Negocio no indicado"}), 400
    try:
        ValidacionIA.query.filter_by(negocio=negocio).delete(synchronize_session=False)
        Reporte.query.filter_by(negocio=negocio).delete(synchronize_session=False)
        Diferencia.query.filter_by(negocio=negocio).delete(synchronize_session=False)
        Liquidacion.query.filter_by(negocio=negocio).delete(synchronize_session=False)
        db.session.commit()
        return jsonify({"ok": True})
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": str(e)}), 500

@app.route("/api/reset-historial", methods=["POST"])
def reset_historial():
    data    = request.get_json(force=True)
    if data.get("rol") not in ROLES_ADMIN:
        return jsonify({"error": "No autorizado"}), 403
    if data.get("confirmar") != "BORRAR":
        return jsonify({"error": "Confirmación incorrecta"}), 400
    negocio = (data.get("negocio") or "").strip()
    if not Negocio.query.get(negocio):
        return jsonify({"error": "Negocio no indicado"}), 400
    try:
        sem = semana_actual(negocio)
        Diferencia.query.filter(Diferencia.semana != sem, Diferencia.negocio == negocio).delete(synchronize_session=False)
        ValidacionIA.query.filter(ValidacionIA.semana != sem, ValidacionIA.negocio == negocio).delete(synchronize_session=False)
        Reporte.query.filter(Reporte.semana != sem, Reporte.negocio == negocio).delete(synchronize_session=False)
        Liquidacion.query.filter(Liquidacion.semana != sem, Liquidacion.negocio == negocio).delete(synchronize_session=False)
        db.session.commit()
        return jsonify({"ok": True})
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": str(e)}), 500

@app.route("/api/tiendas-reportadas")
def tiendas_reportadas():
    negocio = request.args.get("negocio", "mundo_vegano")
    rows = db.session.query(Reporte.tienda).filter_by(semana=semana_captura(negocio), negocio=negocio).distinct().all()
    return jsonify([r.tienda for r in rows])

@app.route("/api/liquidacion", methods=["POST"])
def guardar_liquidacion():
    data    = request.json or {}
    tienda  = data.get("tienda", "")
    tipo    = data.get("tipo", "Liquidación")
    negocio = data.get("negocio", "mundo_vegano")
    sem     = semana_captura(negocio)
    existente = Liquidacion.query.filter_by(tienda=tienda, semana=sem, tipo=tipo, negocio=negocio).first()
    if existente:
        existente.comentario = data.get("comentario", existente.comentario)
        existente.foto1_b64  = data.get("foto1") or existente.foto1_b64
        existente.foto2_b64  = data.get("foto2") or existente.foto2_b64
        existente.foto3_b64  = data.get("foto3") or existente.foto3_b64
        existente.usuario    = data.get("usuario", existente.usuario)
        existente.fecha      = datetime.now().strftime("%Y%m%d_%H%M%S")
        db.session.commit()
        return jsonify({"ok": True, "id": existente.id})
    liq = Liquidacion(negocio=negocio, tienda=tienda, semana=sem, tipo=tipo,
                      usuario=data.get("usuario", ""),
                      comentario=data.get("comentario", ""),
                      foto1_b64=data.get("foto1") or None,
                      foto2_b64=data.get("foto2") or None,
                      foto3_b64=data.get("foto3") or None,
                      fecha=datetime.now().strftime("%Y%m%d_%H%M%S"))
    db.session.add(liq)
    db.session.commit()
    return jsonify({"ok": True, "id": liq.id})

@app.route("/api/liquidacion")
def ver_liquidacion():
    tienda  = request.args.get("tienda", "")
    negocio = request.args.get("negocio", "mundo_vegano")
    sem     = request.args.get("semana") or (semana_captura(negocio) if tienda else semana_actual(negocio))
    q = Liquidacion.query.filter_by(semana=sem, negocio=negocio)
    if tienda:
        q = q.filter_by(tienda=tienda)
    rows = q.all()
    return jsonify([{
        "id": r.id, "tienda": r.tienda, "usuario": r.usuario,
        "tipo": r.tipo or "Liquidación",
        "comentario": r.comentario, "fecha": r.fecha,
        "foto1_b64": r.foto1_b64 or "", "foto2_b64": r.foto2_b64 or "",
        "foto3_b64": r.foto3_b64 or ""
    } for r in rows])

def productos_inactivos(negocio):
    """Productos desactivados del negocio (no se piden a los displays ni entran al cierre). Se guardan por nombre,
    así siguen desactivados aunque se vuelva a cargar el inventario."""
    cfg = db.session.get(Config, f"{negocio}:productos_inactivos")
    try:
        return set(json.loads(cfg.valor)) if cfg and cfg.valor else set()
    except Exception:
        return set()

@app.route("/api/productos-negocio")
def productos_negocio():
    """Catálogo de productos del negocio (del inventario cargado) con su estado activo/inactivo."""
    negocio = request.args.get("negocio", "mundo_vegano")
    inactivos = productos_inactivos(negocio)
    habilitada = filtro_tiendas(negocio)   # sin cadenas desactivadas (ej. Vindi) ni tiendas quitadas
    cat = {}
    for t, p, item in (db.session.query(Inventario.tienda, Inventario.producto, Inventario.item_nbr)
                       .filter_by(negocio=negocio).all()):
        if not habilitada(t):
            continue
        x = cat.setdefault(p, {"producto": p, "item_nbr": "", "tiendas": 0, "activo": p not in inactivos})
        x["tiendas"] += 1
        x["item_nbr"] = x["item_nbr"] or (item or "")
    # los desactivados que ya no vienen en el inventario no se muestran, pero siguen desactivados si vuelven
    return jsonify(sorted(cat.values(), key=lambda x: (x["activo"], x["producto"])))

@app.route("/api/productos-negocio", methods=["POST"])
def activar_producto():
    data = request.json or {}
    if data.get("rol") not in ROLES_ADMIN:
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    negocio = data.get("negocio")
    producto = (data.get("producto") or "").strip()
    if negocio not in slugs_negocios() or not producto:
        return jsonify({"ok": False, "error": "Datos inválidos"}), 400
    inactivos = productos_inactivos(negocio)
    if data.get("activo"):
        inactivos.discard(producto)
    else:
        inactivos.add(producto)
    clave = f"{negocio}:productos_inactivos"
    cfg = db.session.get(Config, clave)
    if not cfg:
        cfg = Config(clave=clave, valor="")
        db.session.add(cfg)
    cfg.valor = json.dumps(sorted(inactivos), ensure_ascii=False)
    db.session.commit()
    return jsonify({"ok": True, "inactivos": len(inactivos)})

@app.route("/api/productos/<tienda>")
def productos(tienda):
    negocio = request.args.get("negocio", "mundo_vegano")
    inactivos = productos_inactivos(negocio)
    rows = [r for r in Inventario.query.filter_by(tienda=tienda, negocio=negocio).order_by(Inventario.cantidad.desc()).all()
            if r.producto not in inactivos]
    return jsonify([{"nombre": r.producto, "cantidad": r.cantidad,
                     "nombre_empaque": r.nombre_empaque or r.producto,
                     "item_nbr": r.item_nbr or ""} for r in rows])

@app.route("/api/sync", methods=["POST"])
def sync_inventario():
    data = request.json
    if data.get("key") != SYNC_KEY:
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    inventario = data.get("inventario", {})
    negocio    = data.get("negocio", "mundo_vegano")
    Inventario.query.filter_by(negocio=negocio).delete(synchronize_session=False)
    total = 0
    for tienda, prods in inventario.items():
        for p in prods:
            db.session.add(Inventario(negocio=negocio, tienda=tienda, producto=p["nombre"], cantidad=p["cantidad"],
                                       nombre_empaque=p.get("nombre_empaque") or None))
            total += 1
    db.session.commit()
    return jsonify({"ok": True, "productos": total})

@app.route("/api/reporte", methods=["POST"])
def guardar_reporte():
    data        = request.json
    tienda      = data.get("tienda", "")
    producto    = data.get("producto", "")
    comentario  = data.get("comentario", "")
    foto_b64    = data.get("foto", "")
    usuario     = data.get("usuario", "")
    negocio     = data.get("negocio", "mundo_vegano")
    ahora       = datetime.now().strftime("%Y%m%d_%H%M%S")
    sem         = semana_captura(negocio)

    existente = Reporte.query.filter_by(tienda=tienda, producto=producto, semana=sem, negocio=negocio).first()

    nombre_foto = ""
    if foto_b64:
        nombre_foto = f"{ahora}_{tienda}_{producto[:20]}.jpg".replace(" ", "_")
        with open(os.path.join(FOTOS_DIR, nombre_foto), "wb") as f:
            f.write(base64.b64decode(foto_b64))

    if existente:
        # UPSERT: actualizar el existente
        existente.comentario    = comentario
        existente.foto_b64      = foto_b64 or existente.foto_b64
        existente.foto          = nombre_foto or existente.foto
        existente.usuario       = usuario
        existente.fecha         = ahora
        existente.modificaciones = (existente.modificaciones or 0) + 1
        existente.modificado_en  = ahora
        db.session.commit()
        return jsonify({"ok": True, "id": existente.id, "modificado": True})
    else:
        rep = Reporte(negocio=negocio, tienda=tienda, producto=producto, comentario=comentario,
                      foto=nombre_foto, foto_b64=foto_b64 or None, usuario=usuario,
                      fecha=ahora, semana=sem, modificaciones=0)
        db.session.add(rep)
        db.session.commit()
        return jsonify({"ok": True, "id": rep.id, "modificado": False})

@app.route("/api/reportes")
def ver_reportes():
    tienda  = request.args.get("tienda", "")
    negocio = request.args.get("negocio", "mundo_vegano")
    semana  = request.args.get("semana") or (semana_captura(negocio) if tienda else semana_actual(negocio))
    q = Reporte.query.filter_by(semana=semana, negocio=negocio)
    if tienda:
        q = q.filter_by(tienda=tienda)
    # Cuando se filtra por tienda (pantalla del display, pocas fotos) se incluye foto_b64.
    # Sin filtro (panel supervisor) NO se cargan las fotos en memoria: solo si tiene o no.
    include_b64 = bool(tienda)
    tiene_b64 = (Reporte.foto_b64.isnot(None) & (Reporte.foto_b64 != "")).label("tiene_b64")
    if include_b64:
        rows = [(r, bool(r.foto_b64)) for r in q.order_by(Reporte.id.desc()).all()]
    else:
        rows = q.options(db.defer(Reporte.foto_b64)).add_columns(tiene_b64).order_by(Reporte.id.desc()).all()
    habilitada = filtro_tiendas(negocio)
    rows = [x for x in rows if habilitada(x[0].tienda)]
    return jsonify([{"id": r.id, "tienda": r.tienda, "producto": r.producto,
                     "comentario": r.comentario, "foto": r.foto,
                     "tiene_foto": bool(tb or r.foto),
                     "foto_b64": (r.foto_b64 or "") if include_b64 else "",
                     "usuario": r.usuario, "fecha": r.fecha,
                     "modificaciones": r.modificaciones or 0,
                     "modificado_en": r.modificado_en or ""} for r, tb in rows])

def _grupo_excel(codigo):
    """Prefijo de hoja del Excel de cierre: Walmart y sus formatos van juntos (WM), Auto Mercado (AM), el resto por su código."""
    c = next((c for c in CADENAS_CATALOGO if c["codigo"] == codigo), None)
    if c and c["grupo"] == "walmart":
        return "WM"
    if c and c["grupo"] == "automercado":
        return "AM"
    return codigo

def _dividir_excel_por_cadena(wb, negocio):
    """Cada hoja con columna 'Tienda' se separa en una hoja por cadena del negocio ('WM …', 'AM …', …).
    Filas de cadenas desactivadas se quitan. Hojas sin columna Tienda quedan igual, al final."""
    from copy import copy
    n = Negocio.query.get(negocio)
    activas = parse_cadenas(n.cadenas) if n else []
    extra = [c for c in activas if c not in {x["codigo"] for x in CADENAS_CATALOGO}]
    grupos = list(dict.fromkeys(_grupo_excel(c) for c in activas)) or ["WM", "AM"]

    def grupo_de(tienda):
        cod = cadena_de_tienda(tienda, extra)
        return _grupo_excel(cod) if cod else grupos[0]   # tienda sin prefijo reconocido: primera cadena

    nuevas = {g: [] for g in grupos}
    sin_tienda = []
    for ws in list(wb.worksheets):
        hdr = [c.value for c in ws[1]]
        it = next((i for i, h in enumerate(hdr) if h and str(h).strip().lower() == "tienda"), None)
        if it is None:
            sin_tienda.append(ws)
            continue
        anchos = {k: v.width for k, v in ws.column_dimensions.items()}
        for g in grupos:
            dest = wb.create_sheet(f"{g} {ws.title}"[:31])
            for k, w in anchos.items():
                dest.column_dimensions[k].width = w
            dest.freeze_panes = ws.freeze_panes
            fila_dest = 0
            for fila in ws.iter_rows():
                if fila[0].row > 1:
                    tienda = fila[it].value
                    if not tienda or grupo_de(tienda) != g:
                        continue
                fila_dest += 1
                for c in fila:
                    d = dest.cell(row=fila_dest, column=c.column, value=c.value)
                    if c.has_style:
                        d.font, d.fill, d.border = copy(c.font), copy(c.fill), copy(c.border)
                        d.alignment, d.number_format = copy(c.alignment), c.number_format
            nuevas[g].append(dest)
        wb.remove(ws)
    wb._sheets = [h for g in grupos for h in nuevas[g]] + sin_tienda

@app.route("/api/cerrar-semana", methods=["POST"])
def cerrar_semana():
    data    = request.json or {}
    if data.get("rol") not in ROLES_ADMIN:
        return jsonify({"ok": False, "error": "Solo un administrador puede cerrar la semana"}), 403
    negocio = (data.get("negocio") or "mundo_vegano").strip()
    semana  = semana_actual(negocio)
    if semana >= semana_captura(negocio):
        return jsonify({"ok": False, "error": "Esa semana todavía no termina: se puede cerrar desde el lunes. "
                                              "Los displays siguen mandando reportes a esa semana."}), 409
    habilitada = filtro_tiendas(negocio)   # cadenas desactivadas / tiendas quitadas no entran al cierre
    # Las fotos (foto_b64) NO se cargan aquí: se leen de a una al armar el ZIP (ahorra memoria)
    reportes = [r for r in Reporte.query.options(db.defer(Reporte.foto_b64))
                .filter_by(semana=semana, negocio=negocio).all() if habilitada(r.tienda)]

    # Seguridad: si la semana ya se cerró y no tiene reportes, cerrarla de nuevo borraría el
    # historial correcto y lo reemplazaría por "sin foto" en todo.
    if not reportes and Diferencia.query.filter_by(semana=semana, negocio=negocio).first():
        return jsonify({"ok": False, "error": f"La semana {semana} ya fue cerrada y no tiene reportes. "
                        "No se vuelve a cerrar para no perder el historial. "
                        "Si el ZIP no se descargó, usá 'Descargar último cierre'."}), 409

    # Construir mapa de reportes: {(tienda, producto): reporte}
    mapa = {}
    for r in reportes:
        key = (r.tienda, r.producto)
        if key not in mapa or r.foto:  # preferir el que tiene foto
            mapa[key] = r

    # Semana anterior para detectar reincidentes
    semana_ant = Diferencia.query.filter(Diferencia.semana < semana, Diferencia.negocio == negocio).order_by(Diferencia.semana.desc()).with_entities(Diferencia.semana).first()
    semana_anterior = semana_ant[0] if semana_ant else None
    difs_anteriores = set()
    if semana_anterior:
        for d in Diferencia.query.filter_by(semana=semana_anterior, negocio=negocio).all():
            if d.estado != "OK":
                difs_anteriores.add((d.tienda, d.producto))

    # Mapa de validaciones IA/supervisor para esta semana
    validaciones = {}
    for v in ValidacionIA.query.filter_by(semana=semana, negocio=negocio).all():
        rep = mapa.get((v.tienda, v.producto))
        if rep:
            validaciones[(v.tienda, v.producto)] = v.estado  # APROBADO / RECHAZADO / REVISAR / etc.

    rurales_negocio = rurales_de(negocio)

    # Tiendas rurales que reportaron esta semana
    tiendas_rurales_esta_semana = {r.tienda for r in reportes if r.tienda in rurales_negocio}

    # Tiendas rurales que visitaron en las últimas 3 semanas (historial Diferencia)
    semanas_hist = sorted({
        d.semana for d in Diferencia.query
        .filter(Diferencia.semana < semana, Diferencia.negocio == negocio)
        .with_entities(Diferencia.semana).distinct().all()
    }, reverse=True)[:3]
    tiendas_rurales_visitaron_mes = set(tiendas_rurales_esta_semana)
    if semanas_hist:
        for d in Diferencia.query.filter(
            Diferencia.semana.in_(semanas_hist),
            Diferencia.negocio == negocio,
            Diferencia.estado != "SIN_FOTO"
        ).with_entities(Diferencia.tienda).distinct().all():
            if d.tienda in rurales_negocio:
                tiendas_rurales_visitaron_mes.add(d.tienda)

    # Calcular diferencias (solo tiendas de cadenas activas del negocio)
    inactivos = productos_inactivos(negocio)   # productos desactivados no cuentan como diferencia
    inventario = [i for i in Inventario.query.filter_by(negocio=negocio).all()
                  if habilitada(i.tienda) and i.producto not in inactivos]
    Diferencia.query.filter_by(semana=semana, negocio=negocio).delete()
    for inv in inventario:
        if inv.cantidad <= 0:
            continue
        key = (inv.tienda, inv.producto)
        rep = mapa.get(key)
        val_estado = validaciones.get(key)

        # Rural que no visitó esta semana pero sí este mes → no es diferencia esta semana
        if inv.tienda in rurales_negocio and inv.tienda not in tiendas_rurales_esta_semana:
            if inv.tienda in tiendas_rurales_visitaron_mes:
                continue  # visitó este mes, no contar como diferencia

        if rep and rep.foto and val_estado == "RECHAZADO":
            estado = "RECHAZADO"
            comentario = rep.comentario or ""
        elif rep and rep.foto and val_estado == "REVISAR":
            estado = "REVISAR"
            comentario = rep.comentario or ""
        elif rep and rep.foto and val_estado == "APROBADO":
            estado = "OK"
            comentario = rep.comentario or ""
        elif rep and rep.comentario:
            estado = "CON_JUSTIFICACION"
            comentario = rep.comentario
        else:
            estado = "SIN_FOTO"
            comentario = ""
        db.session.add(Diferencia(negocio=negocio, semana=semana, tienda=inv.tienda, producto=inv.producto,
                                   cantidad=inv.cantidad, estado=estado, comentario=comentario))
    db.session.commit()

    # Generar Excel
    wb = openpyxl.Workbook()

    # Estilos
    fill_rojo     = PatternFill("solid", fgColor="FFD7D7")
    fill_amarillo = PatternFill("solid", fgColor="FFF3CD")
    fill_verde    = PatternFill("solid", fgColor="D4EDDA")
    fill_header   = PatternFill("solid", fgColor="1E2230")
    fill_reinc    = PatternFill("solid", fgColor="FFB3B3")
    font_header   = Font(bold=True, color="00BF8F", size=11)
    font_reinc    = Font(bold=True, color="C0392B")
    thin          = Side(style="thin", color="CCCCCC")
    border        = Border(left=thin, right=thin, top=thin, bottom=thin)
    center        = Alignment(horizontal="center", vertical="center")

    def estilo_header(ws, cols):
        for cell in ws[1]:
            cell.fill   = fill_header
            cell.font   = font_header
            cell.alignment = center
            cell.border = border
        for i, w in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    rango_archivo, rango_legible = rango_de_codigo(semana)

    # Hoja 1 — Diferencias actuales
    ws1 = wb.active
    ws1.title = f"Diferencias {rango_legible}"
    ws1.append(["Tienda", "Producto", "Stock sistema", "Estado", "Comentario", "Reincidente"])
    estilo_header(ws1, [35, 40, 14, 18, 40, 12])

    difs = Diferencia.query.filter_by(semana=semana, negocio=negocio).filter(Diferencia.estado != "OK").order_by(Diferencia.tienda, Diferencia.producto).all()
    for d in difs:
        reinc = "SI" if (d.tienda, d.producto) in difs_anteriores else "No"
        if d.estado == "SIN_FOTO":
            estado_txt = "Sin foto"
        elif d.estado == "RECHAZADO":
            estado_txt = "Foto rechazada"
        else:
            estado_txt = "Con justificación"
        row = [d.tienda, d.producto, d.cantidad, estado_txt, d.comentario or "", reinc]
        ws1.append(row)
        r = ws1.max_row
        if d.estado == "RECHAZADO":
            fill = PatternFill("solid", fgColor="FFB3B3")  # rojo más oscuro
        elif d.estado == "SIN_FOTO":
            fill = fill_rojo
        else:
            fill = fill_amarillo
        for c in range(1, 7):
            ws1.cell(r, c).fill   = fill_reinc if reinc == "SI" else fill
            ws1.cell(r, c).border = border
            ws1.cell(r, c).alignment = Alignment(vertical="center", wrap_text=True)
        if reinc == "SI":
            ws1.cell(r, 6).font = font_reinc

    # Hoja 2 — Solo reincidentes
    ws2 = wb.create_sheet("Reincidentes")
    ws2.append(["Tienda", "Producto", "Stock sistema", "Estado", "Comentario"])
    estilo_header(ws2, [35, 40, 14, 18, 40])
    for d in difs:
        if (d.tienda, d.producto) in difs_anteriores:
            if d.estado == "SIN_FOTO":
                estado_txt = "Sin foto"
            elif d.estado == "RECHAZADO":
                estado_txt = "Foto rechazada"
            else:
                estado_txt = "Con justificación"
            ws2.append([d.tienda, d.producto, d.cantidad, estado_txt, d.comentario or ""])
            r = ws2.max_row
            for c in range(1, 6):
                ws2.cell(r, c).fill   = fill_reinc
                ws2.cell(r, c).font   = font_reinc
                ws2.cell(r, c).border = border

    # Hoja 3 — Todo OK
    ws3 = wb.create_sheet("Exhibidos OK")
    ws3.append(["Tienda", "Producto", "Stock sistema", "Comentario"])
    estilo_header(ws3, [35, 40, 14, 40])
    for d in Diferencia.query.filter_by(semana=semana, negocio=negocio, estado="OK").order_by(Diferencia.tienda).all():
        ws3.append([d.tienda, d.producto, d.cantidad, d.comentario or ""])
        r = ws3.max_row
        for c in range(1, 5):
            ws3.cell(r, c).fill   = fill_verde
            ws3.cell(r, c).border = border

    # Hoja 4 — Historial matricial: todas las tiendas, ordenado por total de diferencias
    inventario_actual = {(inv.tienda, inv.producto): inv.cantidad for inv in Inventario.query.filter_by(negocio=negocio).all()
                         if inv.producto not in inactivos}

    semanas_arch = sorted(set(
        d.semana for d in Diferencia.query.filter_by(negocio=negocio).with_entities(Diferencia.semana).distinct().all()
    ))

    # {(tienda, producto): {semana: 1/0}}
    matriz = {}
    for d in Diferencia.query.filter_by(negocio=negocio).order_by(Diferencia.semana).all():
        key = (d.tienda, d.producto)
        if key not in matriz:
            matriz[key] = {}
        if d.estado != "OK":
            matriz[key][d.semana] = 1
        elif d.semana not in matriz[key]:
            matriz[key][d.semana] = 0

    fill_uno        = PatternFill("solid", fgColor="C62828")
    fill_cero       = PatternFill("solid", fgColor="1B5E20")
    fill_vacio      = PatternFill("solid", fgColor="2a2a2a")
    font_blanco     = Font(bold=True, color="FFFFFF", size=10)
    font_header_col = Font(bold=True, color="FFFFFF", size=9)
    fill_header_col = PatternFill("solid", fgColor="1e4060")
    fill_total_alto = PatternFill("solid", fgColor="5C0000")

    ws4 = wb.create_sheet("Historial diferencias")

    cabeceras = ["Tienda", "Producto", "Stock"]
    for sem in semanas_arch:
        _, leg = rango_de_codigo(sem)
        cabeceras.append(leg)
    cabeceras.append("Total")
    cabeceras.append("Comentario semana actual")
    ws4.append(cabeceras)

    anchos = [32, 40, 8] + [16] * len(semanas_arch) + [8, 45]
    for i, ancho in enumerate(anchos, 1):
        ws4.column_dimensions[ws4.cell(1, i).column_letter].width = ancho
        ws4.cell(1, i).fill = fill_header_col
        ws4.cell(1, i).font = font_header_col
        ws4.cell(1, i).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws4.row_dimensions[1].height = 40

    # Comentarios de la semana que se está cerrando — directo del mapa de reportes
    comentarios_semana = {
        key: (rep.comentario or "")
        for key, rep in mapa.items()
        if rep.comentario
    }

    # Calcular total por fila y ordenar de mayor a menor
    filas_mat = []
    for (tienda, producto), sem_map in matriz.items():
        stock = inventario_actual.get((tienda, producto), "")
        vals = [sem_map.get(s, None) for s in semanas_arch]
        total = sum(1 for v in vals if v == 1)
        comentario = comentarios_semana.get((tienda, producto), "")
        filas_mat.append((tienda, producto, stock, vals, total, comentario))
    filas_mat.sort(key=lambda x: x[4], reverse=True)

    for tienda, producto, stock, vals, total, comentario in filas_mat:
        fila = [tienda, producto, stock] + [""] * len(semanas_arch) + [total, comentario]
        ws4.append(fila)
        r = ws4.max_row
        # Tienda / Producto / Stock
        for c in range(1, 4):
            ws4.cell(r, c).border = border
            ws4.cell(r, c).alignment = Alignment(vertical="center")
        # Celdas de semana
        for i, v in enumerate(vals):
            c = 4 + i
            ws4.cell(r, c).border = border
            ws4.cell(r, c).alignment = Alignment(horizontal="center", vertical="center")
            if v == 1:
                ws4.cell(r, c).fill = fill_uno
                ws4.cell(r, c).font = font_blanco
                ws4.cell(r, c).value = "1"
            elif v == 0:
                ws4.cell(r, c).fill = fill_cero
                ws4.cell(r, c).font = font_blanco
                ws4.cell(r, c).value = "0"
            else:
                ws4.cell(r, c).fill = fill_vacio
        # Total
        c_total = 4 + len(semanas_arch)
        ws4.cell(r, c_total).border = border
        ws4.cell(r, c_total).alignment = Alignment(horizontal="center", vertical="center")
        if total >= 3:
            ws4.cell(r, c_total).fill = fill_total_alto
            ws4.cell(r, c_total).font = Font(bold=True, color="FFFFFF")
        # Comentario
        c_com = c_total + 1
        ws4.cell(r, c_com).border = border
        ws4.cell(r, c_com).alignment = Alignment(vertical="center", wrap_text=True)

    # Hoja 5 — Tiendas sin visita
    todas_tiendas = set(inv.tienda for inv in inventario)
    tiendas_con_reporte = set(r.tienda for r in reportes)
    tiendas_sin_visita = sorted(todas_tiendas - tiendas_con_reporte)

    fill_naranja = PatternFill("solid", fgColor="FFE0B2")
    font_naranja = Font(bold=True, color="BF360C")

    ws5 = wb.create_sheet("Tiendas sin visita")
    ws5.append(["Tienda", "Productos en sistema", "Observación"])
    estilo_header(ws5, [40, 20, 40])
    for t in tiendas_sin_visita:
        prods = sum(1 for i in Inventario.query.filter_by(tienda=t, negocio=negocio).all() if i.producto not in inactivos)
        ws5.append([t, prods, "Ningún display reportó esta tienda en la semana"])
        r = ws5.max_row
        for c in range(1, 4):
            ws5.cell(r, c).fill   = fill_naranja
            ws5.cell(r, c).font   = font_naranja
            ws5.cell(r, c).border = border

    # Hoja 6 — Resumen por tienda con % cumplimiento
    ws6 = wb.create_sheet("Resumen por tienda")
    ws6.append(["Tienda", "Total productos", "Con foto", "Sin foto", "% Cumplimiento"])
    estilo_header(ws6, [40, 16, 12, 12, 16])

    resumen_tiendas = {}
    for inv in inventario:
        if inv.cantidad <= 0:
            continue
        t = inv.tienda
        if t not in resumen_tiendas:
            resumen_tiendas[t] = {"total": 0, "ok": 0}
        resumen_tiendas[t]["total"] += 1

    for d in Diferencia.query.filter_by(semana=semana, negocio=negocio, estado="OK").all():
        if d.tienda in resumen_tiendas:
            resumen_tiendas[d.tienda]["ok"] += 1

    filas_resumen = []
    for t, datos in resumen_tiendas.items():
        pct = round(datos["ok"] / datos["total"] * 100) if datos["total"] else 0
        filas_resumen.append((t, datos["total"], datos["ok"], datos["total"] - datos["ok"], pct))
    filas_resumen.sort(key=lambda x: x[4])  # menor % primero

    for t, total, ok, sin_foto, pct in filas_resumen:
        ws6.append([t, total, ok, sin_foto, f"{pct}%"])
        r = ws6.max_row
        fill = fill_verde if pct >= 80 else (fill_amarillo if pct >= 50 else fill_rojo)
        for c in range(1, 6):
            ws6.cell(r, c).fill   = fill
            ws6.cell(r, c).border = border
            ws6.cell(r, c).alignment = Alignment(vertical="center")

    # Hoja 7 — Resumen por display
    ws7 = wb.create_sheet("Resumen por display")
    ws7.append(["Display", "Tiendas visitadas", "Productos reportados", "Con foto", "Última actividad"])
    estilo_header(ws7, [25, 18, 20, 12, 20])

    resumen_display = {}
    for rep in reportes:
        u = rep.usuario or "Desconocido"
        if u not in resumen_display:
            resumen_display[u] = {"tiendas": set(), "total": 0, "fotos": 0, "ultima": ""}
        resumen_display[u]["tiendas"].add(rep.tienda)
        resumen_display[u]["total"] += 1
        if rep.foto:
            resumen_display[u]["fotos"] += 1
        if rep.fecha > resumen_display[u]["ultima"]:
            resumen_display[u]["ultima"] = rep.fecha

    for disp, datos in sorted(resumen_display.items()):
        fecha_fmt = datos["ultima"][:8]
        if len(fecha_fmt) == 8:
            fecha_fmt = f"{fecha_fmt[6:8]}/{fecha_fmt[4:6]}/{fecha_fmt[:4]}"
        ws7.append([disp, len(datos["tiendas"]), datos["total"], datos["fotos"], fecha_fmt])
        r = ws7.max_row
        for c in range(1, 6):
            ws7.cell(r, c).fill   = fill_verde
            ws7.cell(r, c).border = border
            ws7.cell(r, c).alignment = Alignment(vertical="center")

    # Hoja 8 — Visitas Rurales (frecuencia mensual)
    ws8 = wb.create_sheet("Visitas Rurales")
    ws8.append(["Tienda", "Visitó este mes", "Visitó esta semana", "Última semana registrada"])
    estilo_header(ws8, [40, 16, 18, 25])

    fill_rural_ok  = PatternFill("solid", fgColor="D4EDDA")
    fill_rural_mal = PatternFill("solid", fgColor="FFD7D7")
    font_rural_mal = Font(bold=True, color="C0392B")

    todas_rurales_inventario = sorted({inv.tienda for inv in inventario if inv.tienda in rurales_negocio})
    for tienda_r in todas_rurales_inventario:
        visito_mes    = "Sí" if tienda_r in tiendas_rurales_visitaron_mes else "No"
        visito_semana = "Sí" if tienda_r in tiendas_rurales_esta_semana   else "No"
        ultima = "—"
        if semanas_hist:
            for sh in semanas_hist:
                tiene = Diferencia.query.filter_by(semana=sh, tienda=tienda_r, negocio=negocio).filter(Diferencia.estado != "SIN_FOTO").first()
                if tiene:
                    _, leg = rango_de_codigo(sh)
                    ultima = leg
                    break
        if tienda_r in tiendas_rurales_esta_semana:
            _, leg = rango_de_codigo(semana)
            ultima = leg
        ws8.append([tienda_r, visito_mes, visito_semana, ultima])
        r = ws8.max_row
        fill_r = fill_rural_ok if visito_mes == "Sí" else fill_rural_mal
        for c in range(1, 5):
            ws8.cell(r, c).fill      = fill_r
            ws8.cell(r, c).border    = border
            ws8.cell(r, c).alignment = Alignment(vertical="center")
        if visito_mes == "No":
            for c in range(1, 5):
                ws8.cell(r, c).font = font_rural_mal

    # Hojas separadas por cadena del negocio (Walmart / Auto Mercado / otras)
    _dividir_excel_por_cadena(wb, negocio)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    # Mapa de estado por producto para separar carpetas en el ZIP
    estado_cierre = {
        (d.tienda, d.producto): d.estado
        for d in Diferencia.query.filter_by(semana=semana, negocio=negocio).all()
    }

    # ── 1) ARMAR EL ZIP COMPLETO EN DISCO, ANTES DE BORRAR NADA ──────────────
    os.makedirs(CIERRES_DIR, exist_ok=True)
    nombre_zip = f"Cierre_{rango_archivo}.zip"
    ruta_zip = os.path.join(CIERRES_DIR, f"{negocio}__{nombre_zip}")
    ruta_tmp = ruta_zip + ".parcial"
    with zipfile.ZipFile(ruta_tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        # Excel dentro del ZIP
        zf.writestr(f"Diferencias_{rango_archivo}.xlsx", buf.getvalue())

        # Fotos separadas en carpetas (OK / Rechazadas / Diferencias), leídas de a una desde la base
        for (tienda, producto), rep in mapa.items():
            # Solo se salta una foto dañada o inexistente; si falla la ESCRITURA del ZIP
            # (memoria, disco), el error sube y el cierre se detiene SIN borrar nada.
            try:
                foto_b64 = db.session.query(Reporte.foto_b64).filter(Reporte.id == rep.id).scalar()
                if foto_b64:
                    raw = foto_b64.split(",", 1)[1] if "," in foto_b64 else foto_b64
                    foto_bytes = base64.b64decode(raw)
                    del foto_b64, raw
                elif rep.foto:
                    with open(os.path.join(FOTOS_DIR, rep.foto), "rb") as f:
                        foto_bytes = f.read()
                else:
                    continue
            except (ValueError, TypeError, OSError):
                continue  # foto corrupta o archivo inexistente
            tienda_safe   = re.sub(r'[\\/:*?"<>|]', "_", tienda).strip()
            producto_safe = re.sub(r'[\\/:*?"<>|]', "_", producto).strip()[:60]
            estado = estado_cierre.get((tienda, producto), "OK")
            if estado == "RECHAZADO":
                carpeta = f"Fotos_Rechazadas_{rango_archivo}"
            elif estado == "OK":
                carpeta = f"Fotos_OK_{rango_archivo}"
            else:
                carpeta = f"Fotos_Diferencias_{rango_archivo}"
            zf.writestr(f"{carpeta}/{tienda_safe}/{producto_safe}.jpg", foto_bytes, compress_type=zipfile.ZIP_STORED)
            del foto_bytes

        # Fotos de novedades (liquidación, vencidos, etc.) en carpetas por tipo
        for liq in Liquidacion.query.filter_by(semana=semana, negocio=negocio).all():
            tipo_safe   = re.sub(r'[\\/:*?"<>|]', "_", liq.tipo or "Liquidacion").strip()
            tienda_safe = re.sub(r'[\\/:*?"<>|]', "_", liq.tienda).strip()
            for idx, foto_b64 in enumerate([liq.foto1_b64, liq.foto2_b64, liq.foto3_b64], 1):
                if not foto_b64:
                    continue
                try:
                    raw = foto_b64.split(",", 1)[1] if "," in foto_b64 else foto_b64
                    foto_bytes = base64.b64decode(raw)
                except (ValueError, TypeError):
                    continue  # foto corrupta
                zf.writestr(f"{tipo_safe}_{rango_archivo}/{tienda_safe}/foto{idx}.jpg", foto_bytes, compress_type=zipfile.ZIP_STORED)
    # el ZIP quedó completo: recién ahora pasa a ser "el último cierre"
    for viejo in os.listdir(CIERRES_DIR):
        if viejo.startswith(f"{negocio}__") and viejo.endswith(".zip") and viejo != os.path.basename(ruta_zip):
            try: os.remove(os.path.join(CIERRES_DIR, viejo))
            except Exception: pass
    os.replace(ruta_tmp, ruta_zip)
    db.session.expunge_all()   # suelta de memoria los objetos ya usados

    # ── 2) RECIÉN AHORA SE LIMPIA LA SEMANA ──────────────────────────────────
    # Borrar validaciones primero (tienen FK hacia reporte)
    ValidacionIA.query.filter(ValidacionIA.semana <= semana, ValidacionIA.negocio == negocio).delete()
    # Borrar reportes de la semana cerrada y de cualquier semana anterior (limpieza completa)
    Reporte.query.filter(Reporte.semana <= semana, Reporte.negocio == negocio).delete()
    # El panel vuelve a la semana del calendario (la que están llenando los displays).
    # Los reportes de semanas posteriores a la cerrada NO se tocan (llegaron durante el cierre).
    for clave_override in (f"{negocio}:semana_override", "semana_override"):
        cfg_override = Config.query.get(clave_override)
        if cfg_override and (clave_override != "semana_override" or negocio == "mundo_vegano"):
            cfg_override.valor = ""
    # Eliminar liquidaciones de la semana cerrada (sus fotos ya están en el ZIP)
    Liquidacion.query.filter_by(semana=semana, negocio=negocio).delete()
    db.session.commit()


    # Limpiar archivos de fotos del disco ahora que ya están en el ZIP
    try:
        for nombre_archivo in os.listdir(FOTOS_DIR):
            ruta = os.path.join(FOTOS_DIR, nombre_archivo)
            if os.path.isfile(ruta):
                os.remove(ruta)
    except Exception:
        pass

    return send_file(ruta_zip, as_attachment=True, download_name=nombre_zip,
                     mimetype="application/zip")

@app.route("/api/ultimo-cierre")
def ultimo_cierre():
    """Vuelve a descargar el ZIP del último cierre del negocio (por si la descarga se cortó)."""
    negocio = request.args.get("negocio", "mundo_vegano")
    try:
        zips = sorted((f for f in os.listdir(CIERRES_DIR) if f.startswith(f"{negocio}__") and f.endswith(".zip")),
                      key=lambda f: os.path.getmtime(os.path.join(CIERRES_DIR, f)), reverse=True)
    except FileNotFoundError:
        zips = []
    if not zips:
        return jsonify({"ok": False, "error": "No hay un cierre guardado en el servidor"}), 404
    return send_file(os.path.join(CIERRES_DIR, zips[0]), as_attachment=True,
                     download_name=zips[0].split("__", 1)[1], mimetype="application/zip")

@app.route("/api/cargar-inventario", methods=["POST"])
def cargar_inventario():
    """Recibe un Excel con columnas Tienda y Producto, crea registros SIN_FOTO para la semana activa."""
    if request.form.get("rol") not in ROLES_ADMIN:
        return jsonify({"ok": False, "error": "Solo un administrador puede hacer esto"}), 403
    if "archivo" not in request.files:
        return jsonify({"error": "No se recibió archivo"}), 400
    archivo = request.files["archivo"]
    if not archivo.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Solo se aceptan archivos Excel (.xlsx, .xls)"}), 400

    negocio = request.form.get("negocio", "mundo_vegano")
    semana  = semana_actual(negocio)

    try:
        wb = openpyxl.load_workbook(archivo, data_only=True)
        ws = wb.active

        # Buscar la fila que contiene los headers reales (puede no ser la fila 1)
        header_row_idx = None
        headers = []
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
            vals = [str(v).strip().lower() if v is not None else "" for v in row]
            # La fila de headers tiene al menos 3 celdas no vacías y contiene palabras clave
            non_empty = [v for v in vals if v]
            if len(non_empty) >= 3 and any(k in " ".join(non_empty) for k in ["store", "tienda", "producto", "signing", "item", "descripcion"]):
                headers = vals
                header_row_idx = row_idx
                break

        if not headers or header_row_idx is None:
            return jsonify({"error": "No se encontró la fila de encabezados en el archivo"}), 400

        # Detectar columnas por nombre flexible
        def col_idx(nombres):
            for n in nombres:
                for i, h in enumerate(headers):
                    if n in h:
                        return i
            return None

        idx_tienda   = col_idx(["store name", "tienda", "sucursal", "cadena", "store"])
        idx_producto = col_idx(["signing desc", "producto", "product", "descripcion", "nombre", "artículo", "articulo", "item desc"])
        idx_item_nbr = col_idx(["item nbr", "item_nbr", "itemnbr", "codigo", "barcode", "ean", "upc"])

        if idx_tienda is None or idx_producto is None:
            return jsonify({"error": f"No se encontraron columnas de tienda/producto. Columnas detectadas: {[h for h in headers if h][:15]}"}), 400

        # Filtro opcional: solo items activos (columna Item Status = 'A' si existe)
        idx_status = col_idx(["item status", "status", "estado"])

        creados = 0
        omitidos = 0
        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            tienda   = str(row[idx_tienda]).strip()   if len(row) > idx_tienda   and row[idx_tienda]   else ""
            producto = str(row[idx_producto]).strip()  if len(row) > idx_producto and row[idx_producto]  else ""
            if not tienda or not producto or tienda == "None" or producto == "None":
                continue
            # Saltar items inactivos si hay columna de status
            if idx_status is not None and len(row) > idx_status:
                status = str(row[idx_status]).strip().upper() if row[idx_status] else ""
                if status and status not in ("A", "ACTIVO", "ACTIVE", "1"):
                    continue
            item_nbr = ""
            if idx_item_nbr is not None and len(row) > idx_item_nbr and row[idx_item_nbr]:
                item_nbr = str(row[idx_item_nbr]).strip().split(".")[0]  # quitar decimales si viene como float
            # Solo crear si no existe ya en esta semana
            existe = Reporte.query.filter_by(semana=semana, tienda=tienda, producto=producto, negocio=negocio).first()
            if existe:
                omitidos += 1
                continue
            r = Reporte(negocio=negocio, tienda=tienda, producto=producto, item_nbr=item_nbr or None, semana=semana,
                        usuario="inventario", fecha=datetime.now().strftime("%Y-%m-%d"))
            db.session.add(r)
            creados += 1

        db.session.commit()
        # Guardar timestamp de última actualización de inventario
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        clave_ts = f"{negocio}:inventario_ultima_actualizacion"
        cfg_ts = Config.query.get(clave_ts)
        if cfg_ts:
            cfg_ts.valor = ts
        else:
            db.session.add(Config(clave=clave_ts, valor=ts))
        db.session.commit()
        return jsonify({"ok": True, "creados": creados, "omitidos": omitidos, "semana": semana, "modo": "agregar", "timestamp": ts})
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": str(e)}), 500

@app.route("/api/inventario-historial")
def inventario_historial():
    negocio  = request.args.get("negocio", "mundo_vegano")
    clave_ts = f"{negocio}:inventario_ultima_actualizacion"
    cfg = Config.query.get(clave_ts)
    # fallback legacy para mundo_vegano
    if not cfg and negocio == "mundo_vegano":
        cfg = Config.query.get("inventario_ultima_actualizacion")
    return jsonify({"ultima": cfg.valor if cfg else None})

def _buscar_encabezados(wb, requeridas, max_filas=40):
    """Busca en TODAS las hojas (empezando por la activa) la fila que contiene las columnas requeridas.
    Así funcionan exportes de Power BI / Retail Link con notas de filtros u hojas extra."""
    hojas = [wb.active] + [w for w in wb.worksheets if w is not wb.active]
    for ws in hojas:
        for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_filas, values_only=True), 1):
            vals = [str(v).strip() if v is not None else "" for v in row]
            if all(r in vals for r in requeridas):
                return ws, idx, vals
    return None, None, None

def _leer_excel_walmart(archivo):
    """Lee INV MV de Walmart. Devuelve lista de {tienda, producto, cantidad}."""
    wb = openpyxl.load_workbook(io.BytesIO(archivo.read()), data_only=True)
    # Fila de headers (tiene 'Store Name' y 'Signing Desc'), en cualquier hoja
    ws, header_row_idx, headers = _buscar_encabezados(wb, ("Store Name", "Signing Desc"))
    if not headers:
        raise ValueError("No se encontró la fila de encabezados (Store Name / Signing Desc)")

    idx_tienda   = headers.index("Store Name")
    idx_producto = headers.index("Signing Desc")
    idx_status   = headers.index("Item Status") if "Item Status" in headers else None
    headers_lower = [h.lower().replace(" ", "") for h in headers]
    idx_item_nbr = next((i for i, h in enumerate(headers_lower) if h in ("itemnbr", "itemnumber", "itemno")), None)
    # Usar la ÚLTIMA columna 'Curr Str On Hand Qty' (semana más reciente)
    idx_qty = next((i for i, h in reversed(list(enumerate(headers))) if "Curr Str On Hand Qty" in h), None)

    resultados = []
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        tienda   = str(row[idx_tienda]).strip()   if len(row) > idx_tienda   and row[idx_tienda]   else ""
        producto = str(row[idx_producto]).strip()  if len(row) > idx_producto and row[idx_producto]  else ""
        if not tienda or not producto or tienda == "None" or producto == "None":
            continue
        if idx_status is not None and len(row) > idx_status:
            status = str(row[idx_status]).strip().upper() if row[idx_status] else ""
            if status and status not in ("A", "ACTIVO", "ACTIVE", "1"):
                continue
        cantidad = 0
        if idx_qty is not None and len(row) > idx_qty and row[idx_qty] is not None:
            try:
                cantidad = int(float(row[idx_qty]))
            except (ValueError, TypeError):
                cantidad = 0
        item_nbr = ""
        if idx_item_nbr is not None and len(row) > idx_item_nbr and row[idx_item_nbr]:
            item_nbr = str(row[idx_item_nbr]).strip().split(".")[0]
        resultados.append({"tienda": tienda, "producto": producto, "cantidad": cantidad, "item_nbr": item_nbr})
    return resultados


def _leer_excel_automercado(archivo):
    """Lee Excel de Automercado. Tienda=NombrePDX, Producto=NombreProducto, Cantidad=Contenido,
    Código=CodigoProducto (o CodigoBarras). Los encabezados pueden estar en cualquier hoja y fila."""
    wb = openpyxl.load_workbook(io.BytesIO(archivo.read()), data_only=True)
    ws, header_row_idx, headers = _buscar_encabezados(wb, ("NombrePDX", "NombreProducto"))
    if not headers:
        hojas = ", ".join(w.title for w in wb.worksheets)
        raise ValueError(f"No se encontraron las columnas NombrePDX y NombreProducto en ninguna hoja ({hojas})")

    def col_idx(nombre):
        try:
            return headers.index(nombre)
        except ValueError:
            return None

    idx_tienda   = col_idx("NombrePDX")
    idx_producto = col_idx("NombreProducto")
    idx_cantidad = col_idx("Contenido")
    idx_codigo   = col_idx("CodigoProducto")
    if idx_codigo is None:
        idx_codigo = col_idx("CodigoBarras")

    resultados = []
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        tienda   = str(row[idx_tienda]).strip()   if len(row) > idx_tienda   and row[idx_tienda]   else ""
        producto = str(row[idx_producto]).strip()  if len(row) > idx_producto and row[idx_producto]  else ""
        if not tienda or not producto or tienda == "None" or producto == "None":
            continue
        cantidad = 0
        if idx_cantidad is not None and len(row) > idx_cantidad and row[idx_cantidad] is not None:
            try:
                cantidad = int(float(row[idx_cantidad]))
            except (ValueError, TypeError):
                cantidad = 0
        item_nbr = ""
        if idx_codigo is not None and len(row) > idx_codigo and row[idx_codigo] is not None:
            item_nbr = str(row[idx_codigo]).strip().split(".")[0]
        resultados.append({"tienda": tienda, "producto": producto, "cantidad": cantidad, "item_nbr": item_nbr})
    return resultados


@app.route("/api/actualizar-inventario", methods=["POST"])
def actualizar_inventario():
    """Actualiza tabla Inventario con cantidades del Excel. NO toca Reporte (fotos/comentarios intactos)."""
    if request.form.get("rol") not in ROLES_ADMIN:
        return jsonify({"ok": False, "error": "Solo un administrador puede hacer esto"}), 403
    archivo_wm = request.files.get("archivo_wm")
    archivo_am = request.files.get("archivo_am")
    negocio    = request.form.get("negocio", "mundo_vegano")
    if not archivo_wm and not archivo_am:
        return jsonify({"error": "Se requiere al menos un archivo"}), 400

    try:
        registros = []
        if archivo_wm:
            registros += _leer_excel_walmart(archivo_wm)
        if archivo_am:
            registros += _leer_excel_automercado(archivo_am)

        # Cargar inventario del negocio en memoria (evita N+1 queries)
        mapa_inv = {(inv.tienda, inv.producto): inv for inv in Inventario.query.filter_by(negocio=negocio).all()}

        actualizados = 0
        nuevos = 0
        for r in registros:
            key = (r["tienda"], r["producto"])
            inv = mapa_inv.get(key)
            if inv:
                inv.cantidad = r["cantidad"]
                if r.get("item_nbr"):
                    inv.item_nbr = r["item_nbr"]
                actualizados += 1
            else:
                nuevo = Inventario(negocio=negocio, tienda=r["tienda"], producto=r["producto"],
                                   cantidad=r["cantidad"], item_nbr=r.get("item_nbr") or None)
                db.session.add(nuevo)
                nuevos += 1

        # Quitar lo que ya no viene en el archivo, solo de la(s) cadena(s) del archivo subido
        # (WM/MXM/MP/PALI con el de Walmart, Auto Mercado con el suyo). Si un archivo no trajo
        # ninguna tienda de su grupo, no se borra nada de ese grupo (archivo vacío o equivocado).
        grupo = lambda t: _grupo_excel(cadena_de_tienda(t))
        llegaron = {(r["tienda"], r["producto"]) for r in registros}
        grupos = {grupo(r["tienda"]) for r in registros} & ({"WM"} if archivo_wm else set()) | \
                 ({grupo(r["tienda"]) for r in registros} & ({"AM"} if archivo_am else set()))
        quitados = 0
        for key, inv in mapa_inv.items():
            if key not in llegaron and grupo(inv.tienda) in grupos:
                db.session.delete(inv)
                quitados += 1

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        clave_ts = f"{negocio}:inventario_ultima_actualizacion"
        cfg_ts = Config.query.get(clave_ts)
        if cfg_ts:
            cfg_ts.valor = ts
        else:
            db.session.add(Config(clave=clave_ts, valor=ts))
        db.session.commit()
        return jsonify({"ok": True, "actualizados": actualizados, "nuevos": nuevos, "quitados": quitados,
                        "total": actualizados + nuevos, "timestamp": ts})
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": str(e)}), 500


@app.route("/api/reemplazar-inventario", methods=["POST"])
def reemplazar_inventario():
    """Borra reportes SIN_FOTO de la semana activa y carga el Excel desde cero."""
    if request.form.get("rol") not in ROLES_ADMIN:
        return jsonify({"ok": False, "error": "Solo un administrador puede hacer esto"}), 403
    if "archivo" not in request.files:
        return jsonify({"error": "No se recibió archivo"}), 400
    archivo = request.files["archivo"]
    if not archivo.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Solo se aceptan archivos Excel (.xlsx, .xls)"}), 400

    negocio = request.form.get("negocio", "mundo_vegano")
    semana  = semana_actual(negocio)

    try:
        wb = openpyxl.load_workbook(archivo, data_only=True)
        ws = wb.active

        header_row_idx = None
        headers = []
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
            vals = [str(v).strip().lower() if v is not None else "" for v in row]
            non_empty = [v for v in vals if v]
            if len(non_empty) >= 3 and any(k in " ".join(non_empty) for k in ["store", "tienda", "producto", "signing", "item", "descripcion"]):
                headers = vals
                header_row_idx = row_idx
                break

        if not headers or header_row_idx is None:
            return jsonify({"error": "No se encontró la fila de encabezados en el archivo"}), 400

        def col_idx(nombres):
            for n in nombres:
                for i, h in enumerate(headers):
                    if n in h:
                        return i
            return None

        idx_tienda   = col_idx(["store name", "tienda", "sucursal", "cadena", "store"])
        idx_producto = col_idx(["signing desc", "producto", "product", "descripcion", "nombre", "artículo", "articulo", "item desc"])

        if idx_tienda is None or idx_producto is None:
            return jsonify({"error": f"No se encontraron columnas. Detectadas: {[h for h in headers if h][:15]}"}), 400

        idx_status   = col_idx(["item status", "status", "estado"])
        idx_item_nbr = col_idx(["item nbr", "item_nbr", "itemnbr", "codigo", "barcode", "ean", "upc"])

        # Borrar solo los registros SIN foto de esta semana (los que ya tienen foto se conservan)
        borrados = Reporte.query.filter_by(semana=semana, negocio=negocio, foto=None, foto_b64=None).delete(synchronize_session=False)
        db.session.flush()

        creados = 0
        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            tienda   = str(row[idx_tienda]).strip()   if len(row) > idx_tienda   and row[idx_tienda]   else ""
            producto = str(row[idx_producto]).strip()  if len(row) > idx_producto and row[idx_producto]  else ""
            if not tienda or not producto or tienda == "None" or producto == "None":
                continue
            if idx_status is not None and len(row) > idx_status:
                status = str(row[idx_status]).strip().upper() if row[idx_status] else ""
                if status and status not in ("A", "ACTIVO", "ACTIVE", "1"):
                    continue
            item_nbr = ""
            if idx_item_nbr is not None and len(row) > idx_item_nbr and row[idx_item_nbr]:
                item_nbr = str(row[idx_item_nbr]).strip().split(".")[0]
            # No duplicar los que ya tienen foto
            existe = Reporte.query.filter_by(semana=semana, tienda=tienda, producto=producto, negocio=negocio).first()
            if existe:
                continue
            r = Reporte(negocio=negocio, tienda=tienda, producto=producto, item_nbr=item_nbr or None, semana=semana,
                        usuario="inventario", fecha=datetime.now().strftime("%Y-%m-%d"))
            db.session.add(r)
            creados += 1

        db.session.commit()
        return jsonify({"ok": True, "borrados": borrados, "creados": creados, "semana": semana, "modo": "reemplazar"})
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": str(e)}), 500


@app.route("/api/limpiar-reportes", methods=["POST"])
def limpiar_reportes():
    data = request.json or {}
    if data.get("rol") not in ROLES_ADMIN:
        return jsonify({"error": "No autorizado"}), 403
    negocio = data.get("negocio", "mundo_vegano")
    count = Reporte.query.filter_by(negocio=negocio).count()
    ValidacionIA.query.filter_by(negocio=negocio).delete(synchronize_session=False)
    Reporte.query.filter_by(negocio=negocio).delete(synchronize_session=False)
    db.session.commit()
    return jsonify({"ok": True, "borrados": count})

@app.route("/api/limpiar-historial", methods=["POST"])
def limpiar_historial():
    data = request.json or {}
    if data.get("rol") not in ROLES_ADMIN:
        return jsonify({"error": "No autorizado"}), 403
    negocio = data.get("negocio", "mundo_vegano")
    count = Diferencia.query.filter_by(negocio=negocio).count()
    Diferencia.query.filter_by(negocio=negocio).delete(synchronize_session=False)
    db.session.commit()
    return jsonify({"ok": True, "borrados": count})

@app.route("/api/limpiar-todo", methods=["POST"])
def limpiar_todo():
    data = request.json or {}
    if data.get("rol") not in ROLES_ADMIN:
        return jsonify({"error": "No autorizado"}), 403
    negocio = data.get("negocio", "mundo_vegano")
    r = Reporte.query.filter_by(negocio=negocio).count()
    d = Diferencia.query.filter_by(negocio=negocio).count()
    ValidacionIA.query.filter_by(negocio=negocio).delete(synchronize_session=False)
    Reporte.query.filter_by(negocio=negocio).delete(synchronize_session=False)
    Diferencia.query.filter_by(negocio=negocio).delete(synchronize_session=False)
    db.session.commit()
    return jsonify({"ok": True, "reportes_borrados": r, "historial_borrado": d})

@app.route("/api/reporte/<int:reporte_id>/foto")
def ver_foto_reporte(reporte_id):
    rep = Reporte.query.get_or_404(reporte_id)
    return jsonify({"foto_b64": rep.foto_b64 or ""})

@app.route("/api/diferencias-semana")
def diferencias_semana():
    negocio = request.args.get("negocio", "mundo_vegano")
    semana  = semana_actual(negocio)
    habilitada = filtro_tiendas(negocio)
    tiene_b64 = (Reporte.foto_b64.isnot(None) & (Reporte.foto_b64 != "")).label("tiene_b64")
    # sin cargar las fotos en memoria
    reportes = (db.session.query(Reporte.tienda, Reporte.foto, Reporte.comentario, tiene_b64)
                .filter(Reporte.semana == semana, Reporte.negocio == negocio).all())

    # Resumen por tienda
    tiendas_reporte = {}  # {tienda: {fotos, comentarios}}
    for r in reportes:
        t = r.tienda
        if not habilitada(t):
            continue
        if t not in tiendas_reporte:
            tiendas_reporte[t] = {"fotos": 0, "comentarios": 0}
        if r.foto or r.tiene_b64:
            tiendas_reporte[t]["fotos"] += 1
        if r.comentario:
            tiendas_reporte[t]["comentarios"] += 1

    # Productos esperados por tienda según inventario
    inv_por_tienda = {}
    habilitada = filtro_tiendas(negocio)
    inactivos = productos_inactivos(negocio)
    for inv in Inventario.query.filter(Inventario.cantidad > 0, Inventario.negocio == negocio).all():
        if not habilitada(inv.tienda) or inv.producto in inactivos:
            continue
        inv_por_tienda.setdefault(inv.tienda, 0)
        inv_por_tienda[inv.tienda] += 1

    todas_tiendas = sorted(set(inv_por_tienda.keys()))
    resultado = []
    for tienda in todas_tiendas:
        visitada = tienda in tiendas_reporte
        datos = tiendas_reporte.get(tienda, {"fotos": 0, "comentarios": 0})
        resultado.append({
            "tienda":       tienda,
            "visitada":     visitada,
            "productos":    inv_por_tienda.get(tienda, 0),
            "con_foto":     datos["fotos"],
            "con_comentario": datos["comentarios"],
        })

    return jsonify({"tiendas": resultado})

@app.route("/api/reportes/<int:reporte_id>", methods=["DELETE"])
def eliminar_reporte(reporte_id):
    rep = Reporte.query.get_or_404(reporte_id)
    db.session.delete(rep)
    db.session.commit()
    return jsonify({"ok": True})

# ═══════════════════════════════════════════════════════════════════════════════
# VALIDACIÓN IA — Fase 1: encolar y consultar progreso
# ═══════════════════════════════════════════════════════════════════════════════

@app.route("/api/validacion/encolar", methods=["POST"])
def validacion_encolar():
    """Crea registros PENDIENTE en ValidacionIA para todos los reportes con foto de la semana actual."""
    data    = request.json or {}
    if data.get("rol") not in ROLES_ADMIN:
        return jsonify({"error": "No autorizado"}), 403
    negocio = data.get("negocio", "mundo_vegano")
    semana  = semana_actual(negocio)
    reportes = Reporte.query.filter_by(semana=semana, negocio=negocio).all()
    encolados = 0
    for rep in reportes:
        if not rep.foto_b64 and not rep.foto:
            continue
        existe = ValidacionIA.query.filter_by(reporte_id=rep.id, semana=semana).first()
        forzar = data.get("forzar", False)
        if existe:
            if forzar or existe.estado in ("ERROR", "PROCESANDO", "RECHAZADO"):
                existe.estado = "PENDIENTE"
                existe.motivo = None
                existe.procesado_en = None
                encolados += 1
            continue
        inv = Inventario.query.filter_by(tienda=rep.tienda, producto=rep.producto, negocio=negocio).first()
        val = ValidacionIA(
            negocio         = negocio,
            reporte_id      = rep.id,
            semana          = semana,
            tienda          = rep.tienda,
            producto        = rep.producto,
            nombre_empaque  = inv.nombre_empaque if inv and inv.nombre_empaque else None,
            marca           = "",
            estado          = "PENDIENTE",
            creado_en       = datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )
        db.session.add(val)
        encolados += 1

    db.session.commit()
    return jsonify({"ok": True, "encolados": encolados, "semana": semana})


@app.route("/api/validacion/progreso")
def validacion_progreso():
    """Retorna el progreso de validación IA de la semana actual."""
    negocio = request.args.get("negocio", "mundo_vegano")
    semana  = request.args.get("semana", semana_actual(negocio))
    total      = ValidacionIA.query.filter_by(semana=semana, negocio=negocio).count()
    pendientes = ValidacionIA.query.filter_by(semana=semana, negocio=negocio, estado="PENDIENTE").count()
    procesando = ValidacionIA.query.filter_by(semana=semana, negocio=negocio, estado="PROCESANDO").count()
    aprobados  = ValidacionIA.query.filter_by(semana=semana, negocio=negocio, estado="APROBADO").count()
    rechazados = ValidacionIA.query.filter_by(semana=semana, negocio=negocio, estado="RECHAZADO").count()
    revisar    = ValidacionIA.query.filter_by(semana=semana, negocio=negocio, estado="REVISAR").count()
    errores    = ValidacionIA.query.filter_by(semana=semana, negocio=negocio, estado="ERROR").count()

    costo_total = db.session.query(db.func.sum(ValidacionIA.costo_usd))\
        .filter_by(semana=semana, negocio=negocio).scalar() or 0.0

    return jsonify({
        "semana": semana,
        "total": total,
        "pendientes": pendientes,
        "procesando": procesando,
        "aprobados": aprobados,
        "rechazados": rechazados,
        "revisar": revisar,
        "errores": errores,
        "procesadas": aprobados + rechazados + revisar + errores,
        "costo_usd": round(costo_total, 4),
    })


@app.route("/api/validacion/resultados")
def validacion_resultados():
    """Lista de validaciones con filtro por estado para revisión manual."""
    negocio = request.args.get("negocio", "mundo_vegano")
    semana  = request.args.get("semana", semana_actual(negocio))
    estado  = request.args.get("estado", "")

    q = ValidacionIA.query.filter_by(semana=semana, negocio=negocio)
    if estado:
        q = q.filter_by(estado=estado)
    q = q.order_by(ValidacionIA.estado, ValidacionIA.tienda)

    rows = q.all()
    resultado = []
    for v in rows:
        foto_b64 = ""
        if v.reporte_id:
            rep = Reporte.query.get(v.reporte_id)
            if rep:
                foto_b64 = rep.foto_b64 or ""
        resultado.append({
            "id": v.id,
            "reporte_id": v.reporte_id,
            "tienda": v.tienda,
            "producto": v.producto,
            "estado": v.estado,
            "confianza": v.confianza or "",
            "motivo": v.motivo or "",
            "costo_usd": v.costo_usd or 0,
            "tiene_foto": bool(foto_b64),
        })
    return jsonify(resultado)


# ═══════════════════════════════════════════════════════════════════════════════
# VALIDACIÓN IA — Fase 2: worker que llama a Claude Haiku
# ═══════════════════════════════════════════════════════════════════════════════

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

_worker_lock = threading.Lock()
_worker_running = False


PROMPT_IA = """Eres un auditor de displays de productos veganos en supermercados de Costa Rica.
Tu trabajo es verificar que el producto "{producto}" está presente en la foto.

Las fotos son de góndolas con VARIOS productos juntos. Es normal ver múltiples productos en la misma foto.

══ CÓMO EVALUAR ══

PASO 1 — Identificá el empaque del producto solicitado según su color y forma (guía abajo).
PASO 2 — Buscá ese empaque específico en cualquier parte de la foto.
PASO 3 — Verificá que al menos UNA palabra clave del producto es visible en ese empaque.

→ Color correcto + palabra clave visible = APROBADO
→ Color correcto pero texto ilegible = REVISAR
→ El empaque con ese color no aparece en ninguna parte de la foto = RECHAZADO

IMPORTANTE: Cada producto tiene su propio color. Si en la foto hay varios productos veganos, identificá el que corresponde al color del producto solicitado e ignorá los demás.

══ GUÍA DE PRODUCTOS ══

- "JAMON HICKORY VEGANO TOFURKY" → empaque NARANJA (caja cuadrada naranja con personaje) · clave: "Hickory Smoked" o "Deli Slices". OJO: ignorá el Tofurky CELESTE que es otro producto (Chick'n).
- "JAMON VEGANO TOFURKY" → empaque NARANJA (sin "Hickory") · clave: "Deli Slices" o "TOFURKY"
- "PAVO VEGANO TOFURKY" → empaque TOFURKY diferente al naranja del jamón · clave: "Turkey" o "Roasted Turkey"
- "TROZOS DE POLLO VEGANO TOFURKY" → empaque CELESTE/AZUL · clave: "Chick'n" o "Plant-Based Chick'n". OJO: ignorá el Tofurky NARANJA que es jamón.
- "MOZZARELLA FOLLOW YOUR HEART" → empaque VERDE brillante (paquete plano) · clave: "Mozzarella". OJO: ignorá el Follow Your Heart CELESTE que es Provolone.
- "PROVOLONE SLICE FOLLOW YOUR HEART" → empaque CELESTE/TURQUESA (caja rectangular) · clave: "Provolone". OJO: ignorá el Follow Your Heart VERDE que es Mozzarella.
- "RES TORTA FRG" → empaque NARANJA intenso (bolsa naranja grande O bandeja circular naranja, marca FRUGAL) · clave: "TORTAS" o "HAMBURGUESA" o "FRUGAL". OJO: ignorá el Beyond Burger AMARILLO MOSTAZA que es otro producto.
- "ORGANIC BUTTER RICHANDCREAMY MELT" → pote AZUL MARINO oscuro cuadrado · clave: "melt" o "Plant Butter" o "Melt Organic"
- "TORTA BEYOND MEAT / TORTA PROTEINA GUISANTE / BEYOND BURGER" → caja AMARILLO MOSTAZA con letras negras grandes · clave: "Beyond Burger" o "Beyond Meat". OJO: el empaque puede verse opaco o con brillo si hay plástico de freezer encima, pero el tono sigue siendo amarillo/dorado. Si el texto "BEYOND BURGER" es visible aunque el color se vea lavado por el reflejo, APROBADO. Ignorá la Torta Frugal NARANJA que es otro producto.
- "NUGGETS BEYOND / BEYOND CHICKEN NUGGETS" → bolsa VERDE brillante · clave: "Nuggets" o "Chicken Nuggets". OJO: ignorá el Follow Your Heart VERDE que es queso, no nuggets.
- "SALCHICHA BEYOND MEAT" → empaque BEYOND MEAT distinto al amarillo mostaza y al verde · clave: "Sausage"
- "VEGACREAM ZIGGYS" → pote redondo BEIGE/CREMA con etiqueta oscura · clave: "ZIGGY'S" o "VEGACREAM"
- "VEGACREAM ZIGGYS HIERBAS" → pote redondo BEIGE/CREMA con etiqueta VERDE OLIVA · clave: "HIERBAS" o "HIERBAS MIXTAS" o "ZIGGY'S"

══ CUÁNDO RECHAZAR ══
Solo rechazá si el empaque con el color del producto solicitado definitivamente no aparece en ninguna parte de la foto.

Responde EXACTAMENTE en este formato (sin texto adicional):
ESTADO: [APROBADO|RECHAZADO|REVISAR]
CONFIANZA: [alta|media|baja]
MOTIVO: [una sola oración: color visto, palabra clave identificada y decisión]"""

PROMPT_IA_GENERICO = """Eres un auditor de displays de productos {negocio} en supermercados de Costa Rica.
Tu trabajo es verificar que el producto "{producto}" está presente en la foto.

Las fotos son de góndolas con VARIOS productos juntos. Es normal ver múltiples productos en la misma foto.

══ CÓMO EVALUAR ══

PASO 1 — Identificá el empaque del producto solicitado según su color y forma (guía abajo).
PASO 2 — Buscá ese empaque específico en cualquier parte de la foto.
PASO 3 — Verificá que al menos UNA palabra clave del producto es visible en ese empaque.

→ Color correcto + palabra clave visible = APROBADO
→ Color correcto pero texto ilegible = REVISAR
→ El empaque con ese color no aparece en ninguna parte de la foto = RECHAZADO

══ GUÍA DE PRODUCTOS ══
(Agregá una línea por producto, con el nombre como aparece en el inventario. Ejemplo:)
- "NOMBRE DEL PRODUCTO" → empaque COLOR (forma) · clave: "palabra visible". OJO: ignorá el producto parecido de otro color.

══ CUÁNDO RECHAZAR ══
Solo rechazá si el empaque con el color del producto solicitado definitivamente no aparece en ninguna parte de la foto.

Responde EXACTAMENTE en este formato (sin texto adicional):
ESTADO: [APROBADO|RECHAZADO|REVISAR]
CONFIANZA: [alta|media|baja]
MOTIVO: [una sola oración: color visto, palabra clave identificada y decisión]"""

def _prompt_por_defecto(slug, nombre=""):
    if slug == "mundo_vegano":
        return PROMPT_IA
    return PROMPT_IA_GENERICO.replace("{negocio}", nombre or "del cliente")

def prompt_de(slug):
    """Prompt que usa la IA para un negocio: el editado en Sistema → Negocios, o el de fábrica."""
    n = Negocio.query.get(slug)
    if n and (n.prompt or "").strip():
        return n.prompt
    return _prompt_por_defecto(slug, n.nombre if n else "")

PARALELO = 5  # fotos simultáneas


def _procesar_una(val_id, client):
    """Procesa una sola validación. Llamado desde threads paralelos."""
    with app.app_context():
        val = ValidacionIA.query.get(val_id)
        if not val:
            return
        try:
            rep = Reporte.query.get(val.reporte_id) if val.reporte_id else None
            if not rep or (not rep.foto_b64 and not rep.foto):
                val.estado = "ERROR"
                val.motivo = "Sin foto disponible"
                val.procesado_en = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                db.session.commit()
                return

            if rep.foto_b64:
                foto_raw = rep.foto_b64
                if "," in foto_raw:
                    foto_raw = foto_raw.split(",", 1)[1]
            else:
                ruta = os.path.join(FOTOS_DIR, rep.foto)
                with open(ruta, "rb") as f:
                    foto_raw = base64.b64encode(f.read()).decode()

            prompt = prompt_de(val.negocio).replace("{producto}", val.producto or "")
            response = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=150,
                messages=[{"role": "user", "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": foto_raw}},
                    {"type": "text",  "text": prompt},
                ]}],
            )

            texto      = response.content[0].text.strip()
            tokens_in  = response.usage.input_tokens
            tokens_out = response.usage.output_tokens
            costo      = (tokens_in * 1.00 + tokens_out * 5.00) / 1_000_000   # Haiku 4.5

            estado_ia, confianza, motivo = "REVISAR", "baja", texto
            for linea in texto.splitlines():
                linea = linea.strip()
                if linea.startswith("ESTADO:"):
                    v = linea.split(":", 1)[1].strip().upper()
                    if v in ("APROBADO", "RECHAZADO", "REVISAR"):
                        estado_ia = v
                elif linea.startswith("CONFIANZA:"):
                    c = linea.split(":", 1)[1].strip().lower()
                    if c in ("alta", "media", "baja"):
                        confianza = c
                elif linea.startswith("MOTIVO:"):
                    motivo = linea.split(":", 1)[1].strip()

            val.estado        = estado_ia
            val.confianza     = confianza
            val.motivo        = motivo
            val.tokens_input  = tokens_in
            val.tokens_output = tokens_out
            val.costo_usd     = costo
            val.procesado_en  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            db.session.commit()

        except Exception as e:
            val.estado       = "ERROR"
            val.motivo       = str(e)[:300]
            val.procesado_en = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            db.session.commit()


def _procesar_validaciones():
    """Procesa en background todos los registros PENDIENTE con PARALELO threads simultáneos."""
    global _worker_running
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        with app.app_context():
            while True:
                # Tomar hasta PARALELO pendientes y marcarlos como PROCESANDO
                pendientes = ValidacionIA.query.filter_by(estado="PENDIENTE").limit(PARALELO).all()
                if not pendientes:
                    break

                for val in pendientes:
                    val.estado  = "PROCESANDO"
                    val.intentos = (val.intentos or 0) + 1
                db.session.commit()

                ids = [val.id for val in pendientes]

            # Lanzar en paralelo
                hilos = [threading.Thread(target=_procesar_una, args=(vid, client), daemon=True)
                         for vid in ids]
                for h in hilos:
                    h.start()
                for h in hilos:
                    h.join()
    finally:
        with _worker_lock:
            _worker_running = False


@app.route("/api/validacion/<int:val_id>/decision", methods=["POST"])
def validacion_decision(val_id):
    """Supervisor aprueba o deniega manualmente una validación REVISAR o RECHAZADO."""
    data = request.json or {}
    if data.get("rol") not in ROLES_ADMIN:
        return jsonify({"error": "No autorizado"}), 403
    decision = data.get("decision", "").upper()
    if decision not in ("APROBADO", "RECHAZADO"):
        return jsonify({"error": "decision debe ser APROBADO o RECHAZADO"}), 400
    val = ValidacionIA.query.get_or_404(val_id)
    val.estado = decision
    val.motivo = (val.motivo or "") + f" [Revisado manualmente por supervisor]"
    val.procesado_en = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    db.session.commit()
    return jsonify({"ok": True, "estado": val.estado})


@app.route("/api/validacion/<int:val_id>/reintentar", methods=["POST"])
def validacion_reintentar(val_id):
    """Marca una validación en ERROR como PENDIENTE para que el worker la reprocese."""
    if (request.get_json(silent=True) or {}).get("rol") not in ROLES_ADMIN:
        return jsonify({"ok": False, "error": "Solo un administrador puede hacer esto"}), 403
    val = ValidacionIA.query.get_or_404(val_id)
    if val.estado != "ERROR":
        return jsonify({"error": "Solo se pueden reintentar validaciones en estado ERROR"}), 400
    val.estado = "PENDIENTE"
    val.motivo = ""
    db.session.commit()
    # Lanzar worker si no está corriendo
    global _worker_running
    with _worker_lock:
        if not _worker_running:
            import anthropic as _anthropic
            client = _anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
            t = threading.Thread(target=_procesar_validaciones, daemon=True)
            t.start()
    return jsonify({"ok": True})

@app.route("/api/validacion/procesar", methods=["POST"])
def validacion_procesar():
    """Lanza el worker en background para procesar fotos PENDIENTE con Claude Haiku."""
    global _worker_running
    if not ANTHROPIC_API_KEY:
        return jsonify({"error": "ANTHROPIC_API_KEY no configurada"}), 500

    data = request.json or {}
    if data.get("rol") not in ROLES_ADMIN:
        return jsonify({"error": "No autorizado"}), 403
    negocio   = data.get("negocio", "mundo_vegano")
    pendientes = ValidacionIA.query.filter_by(estado="PENDIENTE", negocio=negocio).count()
    if pendientes == 0:
        return jsonify({"ok": True, "mensaje": "No hay fotos pendientes", "iniciado": False})

    with _worker_lock:
        if _worker_running:
            return jsonify({"ok": True, "mensaje": "Worker ya en ejecución", "iniciado": False})
        _worker_running = True

    t = threading.Thread(target=_procesar_validaciones, daemon=True)
    t.start()
    return jsonify({"ok": True, "mensaje": f"Procesando {pendientes} fotos en background", "iniciado": True})


DB_LIMITE_MB = int(os.environ.get("DB_LIMITE_MB", "5120"))   # disco del Postgres en Render (displays-db: 5 GB + autoscaling)
TAM_FOTO_ESTIMADO = 222_000   # caracteres base64 de una foto real (promedio medido en producción, 1200px), si aún no hay fotos

def tamano_db_mb():
    try:
        if db.engine.dialect.name == "postgresql":
            r = db.session.execute(db.text("SELECT pg_database_size(current_database())")).fetchone()
            return round(r[0] / 1024 / 1024, 2)
        ruta = db.engine.url.database
        return round(os.path.getsize(ruta) / 1024 / 1024, 2) if ruta and os.path.exists(ruta) else 0
    except Exception:
        return None

def _suma_largo(columna, *filtros):
    return int(db.session.query(db.func.coalesce(db.func.sum(db.func.length(columna)), 0)).filter(*filtros).scalar() or 0)

@app.route("/api/sistema/consumo")
def sistema_consumo():
    """Superadmin: cuánto ocupa cada negocio en la base y cuánto gasta en IA."""
    if request.args.get("solicitante_rol") != "superadmin":
        return jsonify({"error": "No autorizado"}), 403
    MB = 1024 * 1024
    negocios = []
    for n in Negocio.query.order_by(Negocio.orden, Negocio.nombre).all():
        slug = n.slug
        fotos_b = _suma_largo(Reporte.foto_b64, Reporte.negocio == slug)
        n_fotos = Reporte.query.filter(Reporte.negocio == slug, Reporte.foto_b64.isnot(None), Reporte.foto_b64 != "").count()
        liq_b = (_suma_largo(Liquidacion.foto1_b64, Liquidacion.negocio == slug)
                 + _suma_largo(Liquidacion.foto2_b64, Liquidacion.negocio == slug)
                 + _suma_largo(Liquidacion.foto3_b64, Liquidacion.negocio == slug))
        filas = {
            "reportes":      Reporte.query.filter_by(negocio=slug).count(),
            "inventario":    Inventario.query.filter_by(negocio=slug).count(),
            "diferencias":   Diferencia.query.filter_by(negocio=slug).count(),
            "validaciones":  ValidacionIA.query.filter_by(negocio=slug).count(),
            "liquidaciones": Liquidacion.query.filter_by(negocio=slug).count(),
            "usuarios":      Usuario.query.filter_by(negocio=slug).count(),
        }
        datos_b = sum(filas.values()) * 250 + len(n.logo_b64 or "")   # filas sin foto: ~250 bytes c/u
        ia_costo = float(db.session.query(db.func.coalesce(db.func.sum(ValidacionIA.costo_usd), 0.0))
                         .filter(ValidacionIA.negocio == slug).scalar() or 0)
        ia_fotos = ValidacionIA.query.filter(ValidacionIA.negocio == slug,
                                             ValidacionIA.estado.in_(["APROBADO", "RECHAZADO", "REVISAR"])).count()
        # proyección: una semana completa = una foto por producto con existencia en tiendas habilitadas
        habilitada = filtro_tiendas(slug)
        esperadas = sum(1 for i in Inventario.query.filter(Inventario.negocio == slug, Inventario.cantidad > 0)
                        .with_entities(Inventario.tienda).all() if habilitada(i.tienda))
        prom_foto = (fotos_b / n_fotos) if n_fotos else TAM_FOTO_ESTIMADO
        # proyección de fotos: tiendas semanales cada semana, rurales una vez al mes (semana pico = normal + rurales)
        rur, inact = rurales_de(slug), productos_inactivos(slug)
        f_sem = f_rur = 0; t_sem, t_rur, prods = set(), set(), set()
        for tienda, prod in (Inventario.query.filter(Inventario.negocio == slug, Inventario.cantidad > 0)
                             .with_entities(Inventario.tienda, Inventario.producto).all()):
            if not habilitada(tienda) or prod in inact:
                continue
            prods.add(prod)
            if (tienda or "").strip().upper() in rur:
                f_rur += 1; t_rur.add(tienda)
            else:
                f_sem += 1; t_sem.add(tienda)
        negocios.append({
            "slug": slug, "nombre": n.nombre, "activo": bool(n.activo),
            "logo": f"/api/negocios/{slug}/logo" if n.logo_b64 else "",
            "fotos_mb": round(fotos_b / MB, 2), "n_fotos": n_fotos,
            "liquidaciones_mb": round(liq_b / MB, 2),
            "datos_mb": round(datos_b / MB, 2),
            "total_mb": round((fotos_b + liq_b + datos_b) / MB, 2),
            "filas": filas,
            "ia_costo_usd": round(ia_costo, 4), "ia_fotos": ia_fotos,
            "fotos_semana_esperadas": esperadas,
            "proyeccion_semana_mb": round(esperadas * prom_foto / MB, 1),
            "proy": {"tiendas_semanales": len(t_sem), "tiendas_rurales": len(t_rur), "productos": len(prods),
                     "fotos_semana": f_sem, "fotos_rurales": f_rur, "fotos_pico": f_sem + f_rur,
                     "fotos_mes": round(f_sem * 52 / 12 + f_rur), "mb_por_foto": round(prom_foto / MB, 3)},
        })
    negocios.sort(key=lambda x: -x["total_mb"])
    total = tamano_db_mb()
    medido = sum(x["total_mb"] for x in negocios)
    proy = sum(x["proyeccion_semana_mb"] for x in negocios if x["activo"])
    return jsonify({
        "db_mb": total, "limite_mb": DB_LIMITE_MB,
        "medido_mb": round(medido, 2),
        "otros_mb": round(max((total or 0) - medido, 0), 2),   # índices, historial interno, espacio libre de la base
        "proyeccion_semana_mb": round(proy, 1),
        "pico_estimado_mb": round((total or 0) - sum(x["fotos_mb"] + x["liquidaciones_mb"] for x in negocios) + proy, 1),
        "ia_costo_total_usd": round(sum(x["ia_costo_usd"] for x in negocios), 4),
        # costo real por foto validada (si todavía no hay validaciones: 0.0024 medido en producción)
        "ia_por_foto": round((sum(x["ia_costo_usd"] for x in negocios) / sum(x["ia_fotos"] for x in negocios))
                             if sum(x["ia_fotos"] for x in negocios) else 0.0024, 5),
        "motor": db.engine.dialect.name,
        "negocios": negocios,
    })

@app.route("/api/sistema/stats")
def sistema_stats():
    data = request.args
    if request.args.get("rol") not in ROLES_ADMIN:
        return jsonify({"error": "No autorizado"}), 403
    negocio = request.args.get("negocio", "mundo_vegano")

    # Conteo de registros
    n_reportes     = Reporte.query.filter_by(negocio=negocio).count()
    n_diferencias  = Diferencia.query.filter_by(negocio=negocio).count()
    n_inventario   = Inventario.query.filter_by(negocio=negocio).count()
    n_usuarios     = Usuario.query.count()
    n_validaciones = ValidacionIA.query.filter_by(negocio=negocio).count()

    # Costo acumulado total en la llave
    costo_total = db.session.query(db.func.sum(ValidacionIA.costo_usd)).scalar() or 0.0
    fotos_total = ValidacionIA.query.filter(ValidacionIA.estado.in_(["APROBADO","RECHAZADO","REVISAR"])).count()

    # Tamaño de la base de datos
    db_size_mb = tamano_db_mb()

    # Semanas con historial
    semanas = db.session.query(Diferencia.semana).distinct().count()

    return jsonify({
        "registros": {
            "reportes":     n_reportes,
            "diferencias":  n_diferencias,
            "inventario":   n_inventario,
            "usuarios":     n_usuarios,
            "validaciones": n_validaciones,
        },
        "ia": {
            "fotos_procesadas": fotos_total,
            "costo_usd":        round(costo_total, 4),
        },
        "db": {
            "size_mb":    db_size_mb,
            "limit_mb":   1024,
            "semanas_historial": semanas,
        }
    })


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
