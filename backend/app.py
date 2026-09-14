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
CORS(app)

DATABASE_URL = os.environ.get("DATABASE_URL", "")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

if DATABASE_URL:
    app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URL
else:
    DB_PATH = os.path.join(os.path.dirname(__file__), "reportes.db")
    app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"

app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

class Reporte(db.Model):
    id         = db.Column(db.Integer, primary_key=True)
    tienda     = db.Column(db.String(200))
    producto   = db.Column(db.String(200))
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
    tienda          = db.Column(db.String(200), index=True)
    producto        = db.Column(db.String(200))
    cantidad        = db.Column(db.Integer, default=0)
    nombre_empaque  = db.Column(db.String(300))

class Diferencia(db.Model):
    id         = db.Column(db.Integer, primary_key=True)
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

class Config(db.Model):
    clave = db.Column(db.String(50), primary_key=True)
    valor = db.Column(db.String(100))

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
    {"nombre": "Display 1", "usuario": "display1", "password": "1234",     "rol": "display"},
    {"nombre": "Supervisor", "usuario": "admin",    "password": "admin123", "rol": "supervisor"}
]

def _migrar_columnas():
    # db.create_all() no altera tablas existentes; agrega columnas nuevas a mano.
    with db.engine.connect() as conn:
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

def _seed_usuarios():
    if Usuario.query.count() == 0:
        for u in USUARIOS_INICIALES:
            db.session.add(Usuario(**u))
        db.session.commit()

with app.app_context():
    db.create_all()
    _migrar_columnas()
    _seed_usuarios()

SYNC_KEY = os.environ.get("SYNC_KEY", "cpfr2024")

FOTOS_DIR = os.path.join(os.path.dirname(__file__), '..', 'fotos')
os.makedirs(FOTOS_DIR, exist_ok=True)

def semana_actual():
    cfg = Config.query.get("semana_override")
    if cfg and cfg.valor:
        return cfg.valor
    return datetime.now().strftime("%Y-S%V")

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
        return jsonify({"ok": True, "id": u.id, "nombre": u.nombre, "rol": u.rol})
    return jsonify({"ok": False, "error": "Usuario o contraseña incorrectos"}), 401

@app.route("/api/semana-activa")
def get_semana_activa():
    codigo = semana_actual()
    _, legible = rango_de_codigo(codigo)
    ya_cerrada = Diferencia.query.filter_by(semana=codigo).first() is not None
    return jsonify({"semana_activa": codigo, "semana_legible": legible, "ya_cerrada": ya_cerrada})

@app.route("/api/set-semana", methods=["POST"])
def set_semana():
    data = request.json or {}
    semana = (data.get("semana") or "").strip()
    cfg = Config.query.get("semana_override")
    if semana:
        if not cfg:
            cfg = Config(clave="semana_override", valor=semana)
            db.session.add(cfg)
        else:
            cfg.valor = semana
    else:
        if cfg:
            cfg.valor = ""
    db.session.commit()
    return jsonify({"ok": True, "semana_activa": semana_actual()})

@app.route("/api/semanas")
def listar_semanas():
    rows = SemanaConfig.query.order_by(SemanaConfig.id.desc()).all()
    return jsonify([{
        "id": s.id, "codigo": s.codigo, "nombre": s.nombre,
        "fecha_ini": s.fecha_ini, "fecha_fin": s.fecha_fin,
        "activa": s.activa, "creada_en": s.creada_en
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
    s = SemanaConfig.query.get_or_404(sid)
    # Desactivar todas
    SemanaConfig.query.update({"activa": False})
    s.activa = True
    db.session.commit()
    # Actualizar semana_override en Config para que todo el sistema lo use
    cfg = Config.query.get("semana_override")
    if not cfg:
        cfg = Config(clave="semana_override", valor=s.codigo)
        db.session.add(cfg)
    else:
        cfg.valor = s.codigo
    db.session.commit()
    return jsonify({"ok": True, "semana_activa": s.codigo, "nombre": s.nombre})

@app.route("/api/semanas/<int:sid>", methods=["DELETE"])
def eliminar_semana(sid):
    s = SemanaConfig.query.get_or_404(sid)
    if s.activa:
        return jsonify({"error": "No se puede eliminar la semana activa"}), 400
    db.session.delete(s)
    db.session.commit()
    return jsonify({"ok": True})

def _es_supervisor(data):
    return (data or {}).get("solicitante_rol") == "supervisor"

@app.route("/api/usuarios")
def listar_usuarios():
    rows = Usuario.query.order_by(Usuario.rol.desc(), Usuario.nombre).all()
    return jsonify([{"id": u.id, "nombre": u.nombre, "usuario": u.usuario, "password": u.password, "rol": u.rol} for u in rows])

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
    u = Usuario(nombre=nombre, usuario=usuario, password=password, rol="display")
    db.session.add(u)
    db.session.commit()
    return jsonify({"ok": True, "id": u.id})

@app.route("/api/usuarios/<int:usuario_id>", methods=["DELETE"])
def eliminar_usuario(usuario_id):
    if not _es_supervisor(request.json):
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    u = Usuario.query.get_or_404(usuario_id)
    if u.rol == "supervisor":
        return jsonify({"ok": False, "error": "No se puede eliminar al supervisor"}), 400
    db.session.delete(u)
    db.session.commit()
    return jsonify({"ok": True})

@app.route("/api/tiendas")
def tiendas():
    rows = db.session.query(Inventario.tienda).distinct().order_by(Inventario.tienda).all()
    tiendas_list = [r.tienda for r in rows]
    ocultar_am = Config.query.get("ocultar_automercado")
    ocultar_wm = Config.query.get("ocultar_walmart")
    if ocultar_am and ocultar_am.valor == "1":
        tiendas_list = [t for t in tiendas_list if "AUTO" not in t.upper() and not t.upper().startswith("AM ")]
    if ocultar_wm and ocultar_wm.valor == "1":
        tiendas_list = [t for t in tiendas_list if not t.upper().startswith("WM ") and "WALMART" not in t.upper()]
    return jsonify(tiendas_list)

@app.route("/api/config/ocultar-cadena", methods=["POST"])
def toggle_ocultar_cadena():
    data = request.get_json(force=True)
    clave = data.get("clave")
    if clave not in ("ocultar_automercado", "ocultar_walmart"):
        return jsonify({"error": "clave inválida"}), 400
    valor = "1" if data.get("ocultar") else "0"
    cfg = Config.query.get(clave)
    if cfg:
        cfg.valor = valor
    else:
        db.session.add(Config(clave=clave, valor=valor))
    db.session.commit()
    return jsonify({"ok": True, "ocultar": valor == "1"})

@app.route("/api/config/ocultar-cadena", methods=["GET"])
def get_ocultar_cadena():
    clave = request.args.get("clave", "ocultar_automercado")
    cfg = Config.query.get(clave)
    return jsonify({"ocultar": cfg.valor == "1" if cfg else False})

@app.route("/api/reset-historial", methods=["POST"])
def reset_historial():
    data = request.get_json(force=True)
    if data.get("confirmar") != "BORRAR":
        return jsonify({"error": "Confirmación incorrecta"}), 400
    try:
        sem = semana_actual()
        Diferencia.query.filter(Diferencia.semana != sem).delete(synchronize_session=False)
        Reporte.query.filter(Reporte.semana != sem).delete(synchronize_session=False)
        ValidacionIA.query.filter(ValidacionIA.semana != sem).delete(synchronize_session=False)
        Liquidacion.query.filter(Liquidacion.semana != sem).delete(synchronize_session=False)
        db.session.commit()
        return jsonify({"ok": True})
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": str(e)}), 500

@app.route("/api/tiendas-reportadas")
def tiendas_reportadas():
    rows = db.session.query(Reporte.tienda).filter_by(semana=semana_actual()).distinct().all()
    return jsonify([r.tienda for r in rows])

@app.route("/api/liquidacion", methods=["POST"])
def guardar_liquidacion():
    data = request.json or {}
    tienda = data.get("tienda", "")
    tipo   = data.get("tipo", "Liquidación")
    sem = semana_actual()
    # UPSERT: un registro por tienda+semana+tipo
    existente = Liquidacion.query.filter_by(tienda=tienda, semana=sem, tipo=tipo).first()
    if existente:
        existente.comentario = data.get("comentario", existente.comentario)
        existente.foto1_b64  = data.get("foto1") or existente.foto1_b64
        existente.foto2_b64  = data.get("foto2") or existente.foto2_b64
        existente.foto3_b64  = data.get("foto3") or existente.foto3_b64
        existente.usuario    = data.get("usuario", existente.usuario)
        existente.fecha      = datetime.now().strftime("%Y%m%d_%H%M%S")
        db.session.commit()
        return jsonify({"ok": True, "id": existente.id})
    liq = Liquidacion(tienda=tienda, semana=sem, tipo=tipo,
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
    tienda = request.args.get("tienda", "")
    sem = request.args.get("semana", semana_actual())
    q = Liquidacion.query.filter_by(semana=sem)
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

@app.route("/api/productos/<tienda>")
def productos(tienda):
    rows = Inventario.query.filter_by(tienda=tienda).order_by(Inventario.cantidad.desc()).all()
    return jsonify([{"nombre": r.producto, "cantidad": r.cantidad,
                     "nombre_empaque": r.nombre_empaque or r.producto} for r in rows])

@app.route("/api/sync", methods=["POST"])
def sync_inventario():
    data = request.json
    if data.get("key") != SYNC_KEY:
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    inventario = data.get("inventario", {})
    Inventario.query.delete()
    total = 0
    for tienda, prods in inventario.items():
        for p in prods:
            db.session.add(Inventario(tienda=tienda, producto=p["nombre"], cantidad=p["cantidad"],
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
    ahora       = datetime.now().strftime("%Y%m%d_%H%M%S")
    sem         = semana_actual()

    # Buscar reporte existente para esta tienda+producto+semana
    existente = Reporte.query.filter_by(tienda=tienda, producto=producto, semana=sem).first()

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
        rep = Reporte(tienda=tienda, producto=producto, comentario=comentario,
                      foto=nombre_foto, foto_b64=foto_b64 or None, usuario=usuario,
                      fecha=ahora, semana=sem, modificaciones=0)
        db.session.add(rep)
        db.session.commit()
        return jsonify({"ok": True, "id": rep.id, "modificado": False})

@app.route("/api/reportes")
def ver_reportes():
    tienda = request.args.get("tienda", "")
    semana = request.args.get("semana", semana_actual())
    q = Reporte.query.filter_by(semana=semana)
    if tienda:
        q = q.filter_by(tienda=tienda)
    rows = q.order_by(Reporte.id.desc()).all()
    # Cuando se filtra por tienda (pantalla del display, pocas fotos) se incluye foto_b64.
    # Sin filtro (panel supervisor, carga masiva) solo se indica si tiene foto para ahorrar bandwidth.
    include_b64 = bool(tienda)
    return jsonify([{"id": r.id, "tienda": r.tienda, "producto": r.producto,
                     "comentario": r.comentario, "foto": r.foto,
                     "tiene_foto": bool(r.foto_b64 or r.foto),
                     "foto_b64": (r.foto_b64 or "") if include_b64 else "",
                     "usuario": r.usuario, "fecha": r.fecha,
                     "modificaciones": r.modificaciones or 0,
                     "modificado_en": r.modificado_en or ""} for r in rows])

@app.route("/api/cerrar-semana", methods=["POST"])
def cerrar_semana():
    semana = semana_actual()
    reportes = Reporte.query.filter_by(semana=semana).all()

    # Construir mapa de reportes: {(tienda, producto): reporte}
    mapa = {}
    for r in reportes:
        key = (r.tienda, r.producto)
        if key not in mapa or r.foto:  # preferir el que tiene foto
            mapa[key] = r

    # Semana anterior para detectar reincidentes
    semana_ant = Diferencia.query.filter(Diferencia.semana < semana).order_by(Diferencia.semana.desc()).with_entities(Diferencia.semana).first()
    semana_anterior = semana_ant[0] if semana_ant else None
    difs_anteriores = set()
    if semana_anterior:
        for d in Diferencia.query.filter_by(semana=semana_anterior).all():
            if d.estado != "OK":
                difs_anteriores.add((d.tienda, d.producto))

    # Mapa de validaciones IA/supervisor para esta semana
    validaciones = {}
    for v in ValidacionIA.query.filter_by(semana=semana).all():
        rep = mapa.get((v.tienda, v.producto))
        if rep:
            validaciones[(v.tienda, v.producto)] = v.estado  # APROBADO / RECHAZADO / REVISAR / etc.

    # Calcular diferencias
    inventario = Inventario.query.all()
    Diferencia.query.filter_by(semana=semana).delete()
    for inv in inventario:
        if inv.cantidad <= 0:
            continue
        key = (inv.tienda, inv.producto)
        rep = mapa.get(key)
        val_estado = validaciones.get(key)

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
        db.session.add(Diferencia(semana=semana, tienda=inv.tienda, producto=inv.producto,
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

    rango_archivo, rango_legible = rango_semana_actual()

    # Hoja 1 — Diferencias actuales
    ws1 = wb.active
    ws1.title = f"Diferencias {rango_legible}"
    ws1.append(["Tienda", "Producto", "Stock sistema", "Estado", "Comentario", "Reincidente"])
    estilo_header(ws1, [35, 40, 14, 18, 40, 12])

    difs = Diferencia.query.filter_by(semana=semana).filter(Diferencia.estado != "OK").order_by(Diferencia.tienda, Diferencia.producto).all()
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
    for d in Diferencia.query.filter_by(semana=semana, estado="OK").order_by(Diferencia.tienda).all():
        ws3.append([d.tienda, d.producto, d.cantidad, d.comentario or ""])
        r = ws3.max_row
        for c in range(1, 5):
            ws3.cell(r, c).fill   = fill_verde
            ws3.cell(r, c).border = border

    # Hoja 4 — Historial matricial: todas las tiendas, ordenado por total de diferencias
    inventario_actual = {(inv.tienda, inv.producto): inv.cantidad for inv in Inventario.query.all()}

    semanas_arch = sorted(set(
        d.semana for d in Diferencia.query.with_entities(Diferencia.semana).distinct().all()
    ))

    # {(tienda, producto): {semana: 1/0}}
    matriz = {}
    for d in Diferencia.query.order_by(Diferencia.semana).all():
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
    todas_tiendas = set(inv.tienda for inv in Inventario.query.all())
    tiendas_con_reporte = set(r.tienda for r in reportes)
    tiendas_sin_visita = sorted(todas_tiendas - tiendas_con_reporte)

    fill_naranja = PatternFill("solid", fgColor="FFE0B2")
    font_naranja = Font(bold=True, color="BF360C")

    ws5 = wb.create_sheet("Tiendas sin visita")
    ws5.append(["Tienda", "Productos en sistema", "Observación"])
    estilo_header(ws5, [40, 20, 40])
    for t in tiendas_sin_visita:
        prods = Inventario.query.filter_by(tienda=t).count()
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
    for inv in Inventario.query.all():
        if inv.cantidad <= 0:
            continue
        t = inv.tienda
        if t not in resumen_tiendas:
            resumen_tiendas[t] = {"total": 0, "ok": 0}
        resumen_tiendas[t]["total"] += 1

    for d in Diferencia.query.filter_by(semana=semana, estado="OK").all():
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

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    # Recopilar fotos + estado de cada producto antes de borrar reportes
    fotos = {}  # {(tienda, producto): foto_b64_str}
    for (tienda, producto), rep in mapa.items():
        if rep.foto_b64:
            fotos[(tienda, producto)] = rep.foto_b64
        elif rep.foto:
            try:
                ruta = os.path.join(FOTOS_DIR, rep.foto)
                with open(ruta, "rb") as f:
                    fotos[(tienda, producto)] = base64.b64encode(f.read()).decode()
            except Exception:
                pass

    # Mapa de estado por producto para separar carpetas en el ZIP
    estado_cierre = {
        (d.tienda, d.producto): d.estado
        for d in Diferencia.query.filter_by(semana=semana).all()
    }

    # Borrar reportes de la semana cerrada y de cualquier semana anterior (limpieza completa)
    Reporte.query.filter(Reporte.semana <= semana).delete()
    # Limpiar override de semana para que la próxima semana sea automática
    cfg_override = Config.query.get("semana_override")
    if cfg_override:
        cfg_override.valor = ""
    db.session.commit()

    # Construir ZIP: Excel + fotos separadas en Fotos_OK y Fotos_Diferencias
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # Excel dentro del ZIP
        zf.writestr(f"Diferencias_{rango_archivo}.xlsx", buf.getvalue())

        # Fotos separadas en 3 carpetas:
        #   Fotos_OK/           → APROBADO o sin validación con foto
        #   Fotos_Rechazadas/   → RECHAZADO por IA o supervisor
        #   Fotos_Diferencias/  → (solo referencia; sin foto no genera archivo)
        for (tienda, producto), foto_b64 in fotos.items():
            try:
                raw = foto_b64.split(",", 1)[1] if "," in foto_b64 else foto_b64
                foto_bytes = base64.b64decode(raw)
                tienda_safe   = re.sub(r'[\\/:*?"<>|]', "_", tienda).strip()
                producto_safe = re.sub(r'[\\/:*?"<>|]', "_", producto).strip()[:60]
                estado = estado_cierre.get((tienda, producto), "OK")
                if estado == "RECHAZADO":
                    carpeta = f"Fotos_Rechazadas_{rango_archivo}"
                elif estado == "OK":
                    carpeta = f"Fotos_OK_{rango_archivo}"
                else:
                    carpeta = f"Fotos_Diferencias_{rango_archivo}"
                path = f"{carpeta}/{tienda_safe}/{producto_safe}.jpg"
                zf.writestr(path, foto_bytes)
            except Exception:
                pass  # foto corrupta o vacía, se omite

        # Fotos de novedades (liquidación, vencidos, etc.) en carpetas por tipo
        liqs = Liquidacion.query.filter_by(semana=semana).all()
        for liq in liqs:
            tipo_safe   = re.sub(r'[\\/:*?"<>|]', "_", liq.tipo or "Liquidacion").strip()
            tienda_safe = re.sub(r'[\\/:*?"<>|]', "_", liq.tienda).strip()
            for idx, foto_b64 in enumerate([liq.foto1_b64, liq.foto2_b64, liq.foto3_b64], 1):
                if not foto_b64:
                    continue
                try:
                    raw = foto_b64.split(",", 1)[1] if "," in foto_b64 else foto_b64
                    foto_bytes = base64.b64decode(raw)
                    zf.writestr(f"{tipo_safe}_{rango_archivo}/{tienda_safe}/foto{idx}.jpg", foto_bytes)
                except Exception:
                    pass
        # Eliminar liquidaciones de la semana cerrada
        Liquidacion.query.filter_by(semana=semana).delete()
        db.session.commit()

    zip_buf.seek(0)
    nombre_zip = f"Cierre_{rango_archivo}.zip"

    # Limpiar archivos de fotos del disco ahora que ya están en el ZIP
    try:
        for nombre_archivo in os.listdir(FOTOS_DIR):
            ruta = os.path.join(FOTOS_DIR, nombre_archivo)
            if os.path.isfile(ruta):
                os.remove(ruta)
    except Exception:
        pass

    return send_file(zip_buf, as_attachment=True, download_name=nombre_zip,
                     mimetype="application/zip")

@app.route("/api/cargar-inventario", methods=["POST"])
def cargar_inventario():
    """Recibe un Excel con columnas Tienda y Producto, crea registros SIN_FOTO para la semana activa."""
    if "archivo" not in request.files:
        return jsonify({"error": "No se recibió archivo"}), 400
    archivo = request.files["archivo"]
    if not archivo.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Solo se aceptan archivos Excel (.xlsx, .xls)"}), 400

    cfg_semana = Config.query.get("semana_activa")
    semana = cfg_semana.valor if cfg_semana else datetime.now().strftime("%Y-W%W")

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
            # Solo crear si no existe ya en esta semana
            existe = Reporte.query.filter_by(semana=semana, tienda=tienda, producto=producto).first()
            if existe:
                omitidos += 1
                continue
            r = Reporte(tienda=tienda, producto=producto, semana=semana,
                        usuario="inventario", fecha=datetime.now().strftime("%Y-%m-%d"))
            db.session.add(r)
            creados += 1

        db.session.commit()
        # Guardar timestamp de última actualización de inventario
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        cfg_ts = Config.query.get("inventario_ultima_actualizacion")
        if cfg_ts:
            cfg_ts.valor = ts
        else:
            db.session.add(Config(clave="inventario_ultima_actualizacion", valor=ts))
        db.session.commit()
        return jsonify({"ok": True, "creados": creados, "omitidos": omitidos, "semana": semana, "modo": "agregar", "timestamp": ts})
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": str(e)}), 500

@app.route("/api/inventario-historial")
def inventario_historial():
    cfg = Config.query.get("inventario_ultima_actualizacion")
    return jsonify({"ultima": cfg.valor if cfg else None})

def _leer_excel_walmart(archivo):
    """Lee INV MV de Walmart. Devuelve lista de {tienda, producto, cantidad}."""
    wb = openpyxl.load_workbook(archivo, data_only=True)
    ws = wb.active
    # Buscar fila de headers (tiene 'Store Name' y 'Signing Desc')
    headers = None
    header_row_idx = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
        vals = [str(v).strip() if v is not None else "" for v in row]
        if "Store Name" in vals and "Signing Desc" in vals:
            headers = vals
            header_row_idx = row_idx
            break
    if not headers:
        raise ValueError("No se encontró la fila de encabezados (Store Name / Signing Desc)")

    idx_tienda   = headers.index("Store Name")
    idx_producto = headers.index("Signing Desc")
    idx_status   = headers.index("Item Status") if "Item Status" in headers else None
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
        resultados.append({"tienda": tienda, "producto": producto, "cantidad": cantidad})
    return resultados


def _leer_excel_automercado(archivo):
    """Lee Excel de Automercado. Tienda=NombrePDX, Producto=NombreProducto, Cantidad=Contenido."""
    wb = openpyxl.load_workbook(archivo, data_only=True)
    ws = wb.active
    headers = [str(v).strip() if v is not None else "" for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]

    def col_idx(nombre):
        try:
            return headers.index(nombre)
        except ValueError:
            return None

    idx_tienda   = col_idx("NombrePDX")
    idx_producto = col_idx("NombreProducto")
    idx_cantidad = col_idx("Contenido")

    if idx_tienda is None or idx_producto is None:
        raise ValueError(f"No se encontraron columnas NombrePDX/NombreProducto. Detectadas: {[h for h in headers if h][:15]}")

    resultados = []
    for row in ws.iter_rows(min_row=2, values_only=True):
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
        resultados.append({"tienda": tienda, "producto": producto, "cantidad": cantidad})
    return resultados


@app.route("/api/actualizar-inventario", methods=["POST"])
def actualizar_inventario():
    """Actualiza tabla Inventario con cantidades del Excel. NO toca Reporte (fotos/comentarios intactos)."""
    archivo_wm = request.files.get("archivo_wm")
    archivo_am = request.files.get("archivo_am")
    if not archivo_wm and not archivo_am:
        return jsonify({"error": "Se requiere al menos un archivo"}), 400

    try:
        registros = []
        if archivo_wm:
            registros += _leer_excel_walmart(archivo_wm)
        if archivo_am:
            registros += _leer_excel_automercado(archivo_am)

        # Actualizar tabla Inventario: upsert por tienda+producto
        actualizados = 0
        nuevos = 0
        for r in registros:
            inv = Inventario.query.filter_by(tienda=r["tienda"], producto=r["producto"]).first()
            if inv:
                inv.cantidad = r["cantidad"]
                actualizados += 1
            else:
                db.session.add(Inventario(tienda=r["tienda"], producto=r["producto"],
                                          cantidad=r["cantidad"]))
                nuevos += 1

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        cfg_ts = Config.query.get("inventario_ultima_actualizacion")
        if cfg_ts:
            cfg_ts.valor = ts
        else:
            db.session.add(Config(clave="inventario_ultima_actualizacion", valor=ts))
        db.session.commit()
        return jsonify({"ok": True, "actualizados": actualizados, "nuevos": nuevos,
                        "total": actualizados + nuevos, "timestamp": ts})
    except Exception as e:
        db.session.rollback()
        return jsonify({"error": str(e)}), 500


@app.route("/api/reemplazar-inventario", methods=["POST"])
def reemplazar_inventario():
    """Borra reportes SIN_FOTO de la semana activa y carga el Excel desde cero."""
    if "archivo" not in request.files:
        return jsonify({"error": "No se recibió archivo"}), 400
    archivo = request.files["archivo"]
    if not archivo.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Solo se aceptan archivos Excel (.xlsx, .xls)"}), 400

    cfg_semana = Config.query.get("semana_activa")
    semana = cfg_semana.valor if cfg_semana else datetime.now().strftime("%Y-W%W")

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

        idx_status = col_idx(["item status", "status", "estado"])

        # Borrar solo los registros SIN foto de esta semana (los que ya tienen foto se conservan)
        borrados = Reporte.query.filter_by(semana=semana, foto=None, foto_b64=None).delete()
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
            # No duplicar los que ya tienen foto
            existe = Reporte.query.filter_by(semana=semana, tienda=tienda, producto=producto).first()
            if existe:
                continue
            r = Reporte(tienda=tienda, producto=producto, semana=semana,
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
    if data.get("rol") != "supervisor":
        return jsonify({"error": "No autorizado"}), 403
    count = Reporte.query.count()
    Reporte.query.delete()
    db.session.commit()
    return jsonify({"ok": True, "borrados": count})

@app.route("/api/limpiar-historial", methods=["POST"])
def limpiar_historial():
    data = request.json or {}
    if data.get("rol") != "supervisor":
        return jsonify({"error": "No autorizado"}), 403
    count = Diferencia.query.count()
    Diferencia.query.delete()
    db.session.commit()
    return jsonify({"ok": True, "borrados": count})

@app.route("/api/limpiar-todo", methods=["POST"])
def limpiar_todo():
    data = request.json or {}
    if data.get("rol") != "supervisor":
        return jsonify({"error": "No autorizado"}), 403
    r = Reporte.query.count()
    d = Diferencia.query.count()
    Reporte.query.delete()
    Diferencia.query.delete()
    db.session.commit()
    return jsonify({"ok": True, "reportes_borrados": r, "historial_borrado": d})

@app.route("/api/reporte/<int:reporte_id>/foto")
def ver_foto_reporte(reporte_id):
    rep = Reporte.query.get_or_404(reporte_id)
    return jsonify({"foto_b64": rep.foto_b64 or ""})

@app.route("/api/diferencias-semana")
def diferencias_semana():
    semana = semana_actual()
    reportes = Reporte.query.filter_by(semana=semana).all()

    # Resumen por tienda
    tiendas_reporte = {}  # {tienda: {fotos, comentarios}}
    for r in reportes:
        t = r.tienda
        if t not in tiendas_reporte:
            tiendas_reporte[t] = {"fotos": 0, "comentarios": 0}
        if r.foto or r.foto_b64:
            tiendas_reporte[t]["fotos"] += 1
        if r.comentario:
            tiendas_reporte[t]["comentarios"] += 1

    # Productos esperados por tienda según inventario
    inv_por_tienda = {}
    for inv in Inventario.query.filter(Inventario.cantidad > 0).all():
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
    data = request.json or {}
    if data.get("rol") != "supervisor":
        return jsonify({"error": "No autorizado"}), 403

    semana = semana_actual()
    reportes = Reporte.query.filter_by(semana=semana).all()
    encolados = 0
    for rep in reportes:
        if not rep.foto_b64 and not rep.foto:
            continue
        # Evitar duplicados (pero re-encolar ERRORes y PROCESANDO huérfanos)
        existe = ValidacionIA.query.filter_by(reporte_id=rep.id, semana=semana).first()
        forzar = data.get("forzar", False)
        if existe:
            if forzar or existe.estado in ("ERROR", "PROCESANDO", "RECHAZADO"):
                existe.estado = "PENDIENTE"
                existe.motivo = None
                existe.procesado_en = None
                encolados += 1
            continue
        inv = Inventario.query.filter_by(tienda=rep.tienda, producto=rep.producto).first()
        val = ValidacionIA(
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
    semana = request.args.get("semana", semana_actual())
    total      = ValidacionIA.query.filter_by(semana=semana).count()
    pendientes = ValidacionIA.query.filter_by(semana=semana, estado="PENDIENTE").count()
    procesando = ValidacionIA.query.filter_by(semana=semana, estado="PROCESANDO").count()
    aprobados  = ValidacionIA.query.filter_by(semana=semana, estado="APROBADO").count()
    rechazados = ValidacionIA.query.filter_by(semana=semana, estado="RECHAZADO").count()
    revisar    = ValidacionIA.query.filter_by(semana=semana, estado="REVISAR").count()
    errores    = ValidacionIA.query.filter_by(semana=semana, estado="ERROR").count()

    costo_total = db.session.query(db.func.sum(ValidacionIA.costo_usd))\
        .filter_by(semana=semana).scalar() or 0.0

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
    semana = request.args.get("semana", semana_actual())
    estado = request.args.get("estado", "")  # RECHAZADO, REVISAR, etc.

    q = ValidacionIA.query.filter_by(semana=semana)
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

            prompt = PROMPT_IA.format(producto=val.producto)
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
            costo      = (tokens_in * 0.80 + tokens_out * 4.00) / 1_000_000

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
    if data.get("rol") != "supervisor":
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
    if data.get("rol") != "supervisor":
        return jsonify({"error": "No autorizado"}), 403

    pendientes = ValidacionIA.query.filter_by(estado="PENDIENTE").count()
    if pendientes == 0:
        return jsonify({"ok": True, "mensaje": "No hay fotos pendientes", "iniciado": False})

    with _worker_lock:
        if _worker_running:
            return jsonify({"ok": True, "mensaje": "Worker ya en ejecución", "iniciado": False})
        _worker_running = True

    t = threading.Thread(target=_procesar_validaciones, daemon=True)
    t.start()
    return jsonify({"ok": True, "mensaje": f"Procesando {pendientes} fotos en background", "iniciado": True})


@app.route("/api/sistema/stats")
def sistema_stats():
    data = request.args
    if request.args.get("rol") != "supervisor":
        return jsonify({"error": "No autorizado"}), 403

    # Conteo de registros
    n_reportes     = Reporte.query.count()
    n_diferencias  = Diferencia.query.count()
    n_inventario   = Inventario.query.count()
    n_usuarios     = Usuario.query.count()
    n_validaciones = ValidacionIA.query.count()

    # Costo acumulado total en la llave
    costo_total = db.session.query(db.func.sum(ValidacionIA.costo_usd)).scalar() or 0.0
    fotos_total = ValidacionIA.query.filter(ValidacionIA.estado.in_(["APROBADO","RECHAZADO","REVISAR"])).count()

    # Tamaño de la base de datos
    db_size_mb = None
    try:
        if db.engine.dialect.name == "postgresql":
            result = db.session.execute(db.text("SELECT pg_database_size(current_database())")).fetchone()
            db_size_mb = round(result[0] / 1024 / 1024, 2)
        else:
            import os as _os
            db_path = _os.path.join(_os.path.dirname(__file__), "reportes.db")
            db_size_mb = round(_os.path.getsize(db_path) / 1024 / 1024, 2) if _os.path.exists(db_path) else 0
    except Exception:
        db_size_mb = None

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
