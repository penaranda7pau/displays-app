import os
import base64
from datetime import datetime
from flask import Flask, jsonify, request, render_template
from flask_cors import CORS
from flask_sqlalchemy import SQLAlchemy

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
    usuario    = db.Column(db.String(100))
    fecha      = db.Column(db.String(20))

class Inventario(db.Model):
    id       = db.Column(db.Integer, primary_key=True)
    tienda   = db.Column(db.String(200), index=True)
    producto = db.Column(db.String(200))
    cantidad = db.Column(db.Integer, default=0)

with app.app_context():
    db.create_all()

USUARIOS = [
    {"id": 1, "nombre": "Display 1", "usuario": "display1", "password": "1234",     "rol": "display"},
    {"id": 2, "nombre": "Supervisor", "usuario": "admin",    "password": "admin123", "rol": "supervisor"}
]

SYNC_KEY = os.environ.get("SYNC_KEY", "cpfr2024")

FOTOS_DIR = os.path.join(os.path.dirname(__file__), '..', 'fotos')
os.makedirs(FOTOS_DIR, exist_ok=True)

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
    for u in USUARIOS:
        if u["usuario"] == usuario and u["password"] == password:
            return jsonify({"ok": True, "id": u["id"], "nombre": u["nombre"], "rol": u["rol"]})
    return jsonify({"ok": False, "error": "Usuario o contraseña incorrectos"}), 401

@app.route("/api/tiendas")
def tiendas():
    rows = db.session.query(Inventario.tienda).distinct().order_by(Inventario.tienda).all()
    return jsonify([r.tienda for r in rows])

@app.route("/api/productos/<tienda>")
def productos(tienda):
    rows = Inventario.query.filter_by(tienda=tienda).order_by(Inventario.cantidad.desc()).all()
    return jsonify([{"nombre": r.producto, "cantidad": r.cantidad} for r in rows])

@app.route("/api/sync", methods=["POST"])
def sync_inventario():
    data = request.json
    if data.get("key") != SYNC_KEY:
        return jsonify({"ok": False, "error": "No autorizado"}), 403
    inventario = data.get("inventario", {})
    Inventario.query.delete()
    total = 0
    for tienda, productos in inventario.items():
        for p in productos:
            db.session.add(Inventario(tienda=tienda, producto=p["nombre"], cantidad=p["cantidad"]))
            total += 1
    db.session.commit()
    return jsonify({"ok": True, "productos": total})

@app.route("/api/reporte", methods=["POST"])
def guardar_reporte():
    data       = request.json
    tienda     = data.get("tienda", "")
    producto   = data.get("producto", "")
    comentario = data.get("comentario", "")
    foto_b64   = data.get("foto", "")
    usuario    = data.get("usuario", "")
    fecha      = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_foto = ""
    if foto_b64:
        nombre_foto = f"{fecha}_{tienda}_{producto[:20]}.jpg".replace(" ", "_")
        with open(os.path.join(FOTOS_DIR, nombre_foto), "wb") as f:
            f.write(base64.b64decode(foto_b64))
    rep = Reporte(tienda=tienda, producto=producto, comentario=comentario,
                  foto=nombre_foto, usuario=usuario, fecha=fecha)
    db.session.add(rep)
    db.session.commit()
    return jsonify({"ok": True, "id": rep.id})

@app.route("/api/reportes")
def ver_reportes():
    tienda = request.args.get("tienda", "")
    if tienda:
        rows = Reporte.query.filter_by(tienda=tienda).order_by(Reporte.id.desc()).all()
    else:
        rows = Reporte.query.order_by(Reporte.id.desc()).all()
    return jsonify([{"id": r.id, "tienda": r.tienda, "producto": r.producto,
                     "comentario": r.comentario, "foto": r.foto,
                     "usuario": r.usuario, "fecha": r.fecha} for r in rows])

@app.route("/api/reportes/<int:reporte_id>", methods=["DELETE"])
def eliminar_reporte(reporte_id):
    rep = Reporte.query.get_or_404(reporte_id)
    db.session.delete(rep)
    db.session.commit()
    return jsonify({"ok": True})

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
