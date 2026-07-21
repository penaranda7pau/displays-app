import os
import io
import base64
from datetime import datetime
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
    usuario    = db.Column(db.String(100))
    fecha      = db.Column(db.String(20))
    semana     = db.Column(db.String(20), default="")  # semana a la que pertenece

class Inventario(db.Model):
    id       = db.Column(db.Integer, primary_key=True)
    tienda   = db.Column(db.String(200), index=True)
    producto = db.Column(db.String(200))
    cantidad = db.Column(db.Integer, default=0)

class Diferencia(db.Model):
    id         = db.Column(db.Integer, primary_key=True)
    semana     = db.Column(db.String(20), index=True)
    tienda     = db.Column(db.String(200))
    producto   = db.Column(db.String(200))
    cantidad   = db.Column(db.Integer, default=0)
    estado     = db.Column(db.String(50))   # SIN_FOTO / CON_JUSTIFICACION / OK
    comentario = db.Column(db.Text)

with app.app_context():
    db.create_all()

USUARIOS = [
    {"id": 1, "nombre": "Display 1", "usuario": "display1", "password": "1234",     "rol": "display"},
    {"id": 2, "nombre": "Supervisor", "usuario": "admin",    "password": "admin123", "rol": "supervisor"}
]

SYNC_KEY = os.environ.get("SYNC_KEY", "cpfr2024")

FOTOS_DIR = os.path.join(os.path.dirname(__file__), '..', 'fotos')
os.makedirs(FOTOS_DIR, exist_ok=True)

def semana_actual():
    now = datetime.now()
    return now.strftime("%Y-S%V")

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
    for tienda, prods in inventario.items():
        for p in prods:
            db.session.add(Inventario(tienda=tienda, producto=p["nombre"], cantidad=p["cantidad"]))
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
    fecha       = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_foto = ""
    if foto_b64:
        nombre_foto = f"{fecha}_{tienda}_{producto[:20]}.jpg".replace(" ", "_")
        with open(os.path.join(FOTOS_DIR, nombre_foto), "wb") as f:
            f.write(base64.b64decode(foto_b64))
    rep = Reporte(tienda=tienda, producto=producto, comentario=comentario,
                  foto=nombre_foto, usuario=usuario, fecha=fecha, semana=semana_actual())
    db.session.add(rep)
    db.session.commit()
    return jsonify({"ok": True, "id": rep.id})

@app.route("/api/reportes")
def ver_reportes():
    tienda = request.args.get("tienda", "")
    semana = request.args.get("semana", semana_actual())
    q = Reporte.query.filter_by(semana=semana)
    if tienda:
        q = q.filter_by(tienda=tienda)
    rows = q.order_by(Reporte.id.desc()).all()
    return jsonify([{"id": r.id, "tienda": r.tienda, "producto": r.producto,
                     "comentario": r.comentario, "foto": r.foto,
                     "usuario": r.usuario, "fecha": r.fecha} for r in rows])

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

    # Calcular diferencias
    inventario = Inventario.query.all()
    Diferencia.query.filter_by(semana=semana).delete()
    for inv in inventario:
        if inv.cantidad <= 0:
            continue
        key = (inv.tienda, inv.producto)
        rep = mapa.get(key)
        if rep and rep.foto:
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

    # Hoja 1 — Diferencias actuales
    ws1 = wb.active
    ws1.title = f"Diferencias {semana}"
    ws1.append(["Tienda", "Producto", "Stock sistema", "Estado", "Comentario", "Reincidente"])
    estilo_header(ws1, [35, 40, 14, 18, 40, 12])

    difs = Diferencia.query.filter_by(semana=semana).filter(Diferencia.estado != "OK").order_by(Diferencia.tienda, Diferencia.producto).all()
    for d in difs:
        reinc = "SI" if (d.tienda, d.producto) in difs_anteriores else "No"
        estado_txt = "Sin foto" if d.estado == "SIN_FOTO" else "Con justificación"
        row = [d.tienda, d.producto, d.cantidad, estado_txt, d.comentario or "", reinc]
        ws1.append(row)
        r = ws1.max_row
        fill = fill_rojo if d.estado == "SIN_FOTO" else fill_amarillo
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
            estado_txt = "Sin foto" if d.estado == "SIN_FOTO" else "Con justificación"
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

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    nombre = f"diferencias_{semana}.xlsx"
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/api/reportes/<int:reporte_id>", methods=["DELETE"])
def eliminar_reporte(reporte_id):
    rep = Reporte.query.get_or_404(reporte_id)
    db.session.delete(rep)
    db.session.commit()
    return jsonify({"ok": True})

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
