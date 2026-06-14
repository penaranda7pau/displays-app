import os
import base64
import threading
import requests
import tkinter as tk
from tkinter import filedialog
from kivy.app import App
from kivy.uix.screenmanager import ScreenManager, Screen, SlideTransition
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.core.window import Window
from kivy.graphics import Color, RoundedRectangle, Rectangle
from kivy.metrics import dp
from kivy.clock import Clock

API = "http://127.0.0.1:5000"
USUARIO_ACTIVO = {}

FONDO    = (0.08, 0.09, 0.12, 1)
TARJETA  = (0.12, 0.14, 0.18, 1)
VERDE    = (0.0,  0.75, 0.55, 1)
TEXTO    = (0.95, 0.95, 0.95, 1)
SUBTEXTO = (0.55, 0.60, 0.65, 1)
ROJO     = (0.95, 0.30, 0.35, 1)
AMARILLO = (0.95, 0.70, 0.10, 1)
AZUL     = (0.20, 0.50, 0.90, 1)

def fondo(widget, color=FONDO):
    with widget.canvas.before:
        Color(*color)
        r = Rectangle(pos=widget.pos, size=widget.size)
    widget.bind(pos=lambda *_: setattr(r, 'pos', widget.pos),
                size=lambda *_: setattr(r, 'size', widget.size))

def boton(texto, color=VERDE, alto=dp(48)):
    b = Button(text=texto, size_hint=(1, None), height=alto,
               background_normal='', background_color=color,
               color=(0.03, 0.1, 0.08, 1) if color == VERDE else TEXTO,
               bold=True, font_size=dp(15))
    return b

def campo(hint, password=False):
    return TextInput(hint_text=hint, password=password,
                     size_hint=(1, None), height=dp(46),
                     background_color=(0.15, 0.17, 0.22, 1),
                     foreground_color=TEXTO, font_size=dp(15),
                     multiline=False, padding=[dp(12), dp(12)])

class LoginScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        root = BoxLayout(orientation='vertical', padding=dp(40), spacing=dp(16))
        fondo(root)
        root.add_widget(Label(size_hint=(1, 1)))
        root.add_widget(Label(text='[b]Inventario Displays[/b]', markup=True,
                              font_size=dp(26), color=VERDE,
                              size_hint=(1, None), height=dp(50), halign='center'))
        root.add_widget(Label(size_hint=(1, None), height=dp(20)))
        self.inp_usuario = campo("Usuario")
        self.inp_password = campo("Contraseña", password=True)
        self.lbl_error = Label(text='', color=ROJO, size_hint=(1, None),
                               height=dp(24), font_size=dp(13), halign='center')
        btn = boton("Ingresar")
        btn.bind(on_release=self._login)
        root.add_widget(self.inp_usuario)
        root.add_widget(self.inp_password)
        root.add_widget(self.lbl_error)
        root.add_widget(btn)
        root.add_widget(Label(size_hint=(1, 1)))
        self.add_widget(root)

    def _login(self, *_):
        u = self.inp_usuario.text.strip()
        p = self.inp_password.text.strip()
        if not u or not p:
            self.lbl_error.text = "Completá usuario y contraseña"
            return
        self.lbl_error.text = "Verificando..."
        def worker():
            try:
                r = requests.post(f"{API}/api/login",
                                  json={"usuario": u, "password": p}, timeout=5)
                data = r.json()
                Clock.schedule_once(lambda _: self._resultado(data), 0)
            except:
                Clock.schedule_once(lambda _: setattr(self.lbl_error, 'text',
                                    "Sin conexión con el servidor"), 0)
        threading.Thread(target=worker, daemon=True).start()

    def _resultado(self, data):
        if data.get("ok"):
            USUARIO_ACTIVO.update(data)
            self.lbl_error.text = ""
            self.manager.transition = SlideTransition(direction='left')
            self.manager.current = 'tiendas'
        else:
            self.lbl_error.text = data.get("error", "Error al ingresar")

class TiendasScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.root_layout = BoxLayout(orientation='vertical', padding=dp(16), spacing=dp(12))
        fondo(self.root_layout)

        hdr = BoxLayout(size_hint=(1, None), height=dp(50), spacing=dp(8))
        hdr.add_widget(Label(text='[b]Seleccionar Tienda[/b]', markup=True,
                             font_size=dp(20), color=TEXTO, halign='left'))
        self.btn_reportes = boton("Ver reportes", color=AZUL, alto=dp(40))
        self.btn_reportes.size_hint = (None, None)
        self.btn_reportes.width = dp(140)
        self.btn_reportes.bind(on_release=lambda *_: self._ir_reportes())
        hdr.add_widget(self.btn_reportes)
        self.root_layout.add_widget(hdr)

        self.buscar = campo("Buscar tienda...")
        self.buscar.bind(text=self._filtrar)
        self.root_layout.add_widget(self.buscar)
        self.status = Label(text='Cargando...', color=SUBTEXTO,
                            size_hint=(1, None), height=dp(24), font_size=dp(13))
        self.root_layout.add_widget(self.status)
        sv = ScrollView(size_hint=(1, 1))
        self.lista = BoxLayout(orientation='vertical', size_hint=(1, None), spacing=dp(6))
        self.lista.bind(minimum_height=self.lista.setter('height'))
        sv.add_widget(self.lista)
        self.root_layout.add_widget(sv)
        self.add_widget(self.root_layout)
        self.tiendas = []

    def on_enter(self):
        # Mostrar botón solo si es supervisor
        self.btn_reportes.opacity = 1 if USUARIO_ACTIVO.get("rol") == "supervisor" else 0
        self.btn_reportes.disabled = USUARIO_ACTIVO.get("rol") != "supervisor"

        self.tiendas = []
        self.lista.clear_widgets()
        self.status.text = "Cargando tiendas..."
        def worker():
            try:
                r = requests.get(f"{API}/api/tiendas", timeout=8)
                data = r.json()
                Clock.schedule_once(lambda _: self._mostrar(data), 0)
            except:
                Clock.schedule_once(lambda _: setattr(self.status, 'text', "Sin conexión"), 0)
        threading.Thread(target=worker, daemon=True).start()

    def _mostrar(self, tiendas):
        self.tiendas = tiendas
        self.status.text = f"{len(tiendas)} tiendas"
        self._filtrar(None, self.buscar.text)

    def _filtrar(self, _, texto):
        self.lista.clear_widgets()
        filtro = texto.lower() if texto else ""
        mostradas = [t for t in self.tiendas if filtro in t.lower()] if filtro else self.tiendas
        for tienda in mostradas:
            btn = Button(text=tienda, size_hint=(1, None), height=dp(52),
                         background_normal='', background_color=TARJETA,
                         color=TEXTO, font_size=dp(14), halign='left',
                         text_size=(Window.width - dp(32), None))
            btn.bind(on_release=lambda _, t=tienda: self._ir_tienda(t))
            self.lista.add_widget(btn)

    def _ir_tienda(self, tienda):
        self.manager.get_screen('productos').tienda = tienda
        self.manager.transition = SlideTransition(direction='left')
        self.manager.current = 'productos'

    def _ir_reportes(self):
        self.manager.transition = SlideTransition(direction='left')
        self.manager.current = 'reportes'

class ReportesScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.root_layout = BoxLayout(orientation='vertical', padding=dp(12), spacing=dp(8))
        fondo(self.root_layout)

        # Header
        hdr = BoxLayout(size_hint=(1, None), height=dp(50), spacing=dp(8))
        btn_back = boton("←", color=TARJETA, alto=dp(44))
        btn_back.size_hint = (None, None)
        btn_back.width = dp(44)
        btn_back.bind(on_release=lambda *_: self._volver())
        self.lbl_titulo = Label(text='[b]Reportes enviados[/b]', markup=True,
                                font_size=dp(17), color=TEXTO, halign='left')
        btn_refrescar = boton("↻ Actualizar", color=TARJETA, alto=dp(44))
        btn_refrescar.size_hint = (None, None)
        btn_refrescar.width = dp(120)
        btn_refrescar.bind(on_release=lambda *_: self._cargar())
        hdr.add_widget(btn_back)
        hdr.add_widget(self.lbl_titulo)
        hdr.add_widget(btn_refrescar)
        self.root_layout.add_widget(hdr)

        # Resumen
        self.resumen = BoxLayout(size_hint=(1, None), height=dp(36), spacing=dp(8))
        self.lbl_total     = self._chip("0 reportes", TARJETA)
        self.lbl_con_foto  = self._chip("0 con foto", (0.08, 0.30, 0.20, 1))
        self.lbl_sin_stock = self._chip("0 sin stock", (0.35, 0.10, 0.10, 1))
        self.resumen.add_widget(self.lbl_total)
        self.resumen.add_widget(self.lbl_con_foto)
        self.resumen.add_widget(self.lbl_sin_stock)
        self.root_layout.add_widget(self.resumen)

        self.status = Label(text='', color=SUBTEXTO, size_hint=(1, None),
                            height=dp(20), font_size=dp(12))
        self.root_layout.add_widget(self.status)

        # Encabezados tabla
        enc = GridLayout(cols=5, size_hint=(1, None), height=dp(32), spacing=dp(2))
        for txt, sw in [("Tienda", 0.30), ("Producto", 0.28), ("Cant.", 0.08),
                        ("Comentario", 0.20), ("Foto", 0.14)]:
            l = Label(text=f"[b]{txt}[/b]", markup=True, font_size=dp(11),
                      color=VERDE, halign='center', valign='middle', size_hint=(sw, 1))
            l.bind(size=l.setter('text_size'))
            enc.add_widget(l)
        self.root_layout.add_widget(enc)

        sv = ScrollView(size_hint=(1, 1))
        self.lista = BoxLayout(orientation='vertical', size_hint=(1, None), spacing=dp(3))
        self.lista.bind(minimum_height=self.lista.setter('height'))
        sv.add_widget(self.lista)
        self.root_layout.add_widget(sv)

        self.add_widget(self.root_layout)

    def _chip(self, texto, color):
        l = Label(text=texto, size_hint=(1, None), height=dp(30),
                  font_size=dp(12), color=TEXTO, bold=True, halign='center')
        with l.canvas.before:
            Color(*color)
            rect = Rectangle(pos=l.pos, size=l.size)
        l.bind(pos=lambda w, *_: setattr(rect, 'pos', w.pos),
               size=lambda w, *_: setattr(rect, 'size', w.size))
        return l

    def on_enter(self):
        self._cargar()

    def _cargar(self):
        self.lista.clear_widgets()
        self.status.text = "Cargando reportes..."
        def worker():
            try:
                r = requests.get(f"{API}/api/reportes", timeout=8)
                data = r.json()
                Clock.schedule_once(lambda _: self._mostrar(data), 0)
            except:
                Clock.schedule_once(lambda _: setattr(self.status, 'text', "Sin conexión"), 0)
        threading.Thread(target=worker, daemon=True).start()

    def _mostrar(self, reportes):
        self.lista.clear_widgets()
        total = len(reportes)
        con_foto = sum(1 for r in reportes if r.get("foto"))
        sin_stock = sum(1 for r in reportes if r.get("cantidad", 1) == 0)

        self.lbl_total.text     = f"{total} reportes"
        self.lbl_con_foto.text  = f"{con_foto} con foto"
        self.lbl_sin_stock.text = f"{sin_stock} sin stock"
        self.status.text = f"Actualizado · {total} registros"

        if not reportes:
            self.lista.add_widget(Label(text="No hay reportes aún", color=SUBTEXTO,
                                        size_hint=(1, None), height=dp(48), font_size=dp(13)))
            return

        for rep in reportes:
            tienda    = rep.get("tienda", "")
            producto  = rep.get("producto", "")
            comentario = rep.get("comentario", "") or "—"
            tiene_foto = "✓" if rep.get("foto") else "—"
            cantidad  = rep.get("cantidad", "")
            usuario   = rep.get("usuario", "")
            fecha     = rep.get("fecha", "")

            fila = GridLayout(cols=5, size_hint=(1, None), height=dp(52), spacing=dp(2))
            with fila.canvas.before:
                Color(*TARJETA)
                rect = Rectangle(pos=fila.pos, size=fila.size)
            fila.bind(pos=lambda w, *_: setattr(rect, 'pos', w.pos),
                      size=lambda w, *_: setattr(rect, 'size', w.size))

            def lbl(txt, sw, color=TEXTO, fs=dp(10)):
                l = Label(text=str(txt), font_size=fs, color=color,
                          halign='center', valign='middle', size_hint=(sw, 1))
                l.bind(size=l.setter('text_size'))
                return l

            fila.add_widget(lbl(tienda,    0.30, fs=dp(9)))
            fila.add_widget(lbl(producto,  0.28, fs=dp(9)))
            fila.add_widget(lbl(str(cantidad) if cantidad != "" else "—", 0.08))
            fila.add_widget(lbl(comentario, 0.20, color=SUBTEXTO, fs=dp(9)))
            color_foto = VERDE if tiene_foto == "✓" else SUBTEXTO
            fila.add_widget(lbl(tiene_foto, 0.14, color=color_foto, fs=dp(14)))

            self.lista.add_widget(fila)

    def _volver(self):
        self.manager.transition = SlideTransition(direction='right')
        self.manager.current = 'tiendas'

class ProductosScreen(Screen):
    tienda = ""

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.root_layout = BoxLayout(orientation='vertical', padding=dp(12), spacing=dp(8))
        fondo(self.root_layout)

        hdr = BoxLayout(size_hint=(1, None), height=dp(50), spacing=dp(8))
        btn_back = boton("←", color=TARJETA, alto=dp(44))
        btn_back.size_hint = (None, None)
        btn_back.width = dp(44)
        btn_back.bind(on_release=lambda *_: self._volver())
        self.lbl_tienda = Label(text='', font_size=dp(15), color=TEXTO, bold=True, halign='left')
        hdr.add_widget(btn_back)
        hdr.add_widget(self.lbl_tienda)
        self.root_layout.add_widget(hdr)

        self.status = Label(text='', color=SUBTEXTO, size_hint=(1, None),
                            height=dp(22), font_size=dp(12))
        self.root_layout.add_widget(self.status)

        enc = GridLayout(cols=4, size_hint=(1, None), height=dp(36), spacing=dp(4))
        for txt, sw in [("Producto", 0.45), ("Cant.", 0.1), ("Foto", 0.2), ("Comentario", 0.25)]:
            l = Label(text=f"[b]{txt}[/b]", markup=True, font_size=dp(12),
                      color=VERDE, halign='center', valign='middle',
                      size_hint=(sw, 1))
            l.bind(size=l.setter('text_size'))
            enc.add_widget(l)
        self.root_layout.add_widget(enc)

        sv = ScrollView(size_hint=(1, 1))
        self.lista = BoxLayout(orientation='vertical', size_hint=(1, None), spacing=dp(4))
        self.lista.bind(minimum_height=self.lista.setter('height'))
        sv.add_widget(self.lista)
        self.root_layout.add_widget(sv)

        btn_enviar_todo = boton("✓  Enviar todos los reportes", alto=dp(46))
        btn_enviar_todo.bind(on_release=self._enviar_todos)
        self.root_layout.add_widget(btn_enviar_todo)

        self.lbl_status = Label(text='', color=SUBTEXTO, size_hint=(1, None),
                                height=dp(22), font_size=dp(12), halign='center')
        self.root_layout.add_widget(self.lbl_status)

        self.add_widget(self.root_layout)
        self.productos = []
        self.filas = []

    def on_enter(self):
        self.lbl_tienda.text = self.tienda
        self.lista.clear_widgets()
        self.filas = []
        self.productos = []
        self.status.text = "Cargando productos..."
        self.lbl_status.text = ""
        def worker():
            try:
                r = requests.get(f"{API}/api/productos/{self.tienda}", timeout=8)
                data = r.json()
                Clock.schedule_once(lambda _: self._mostrar(data), 0)
            except:
                Clock.schedule_once(lambda _: setattr(self.status, 'text', "Sin conexión"), 0)
        threading.Thread(target=worker, daemon=True).start()

    def _mostrar(self, productos):
        self.productos = productos
        self.status.text = f"{len(productos)} productos"
        self.filas = []
        for p in productos:
            nombre   = p["nombre"]
            cantidad = p["cantidad"]
            fila_data = {"nombre": nombre, "cantidad": cantidad, "foto_b64": "", "comentario": ""}

            fila = GridLayout(cols=4, size_hint=(1, None), height=dp(56), spacing=dp(4))
            with fila.canvas.before:
                Color(*TARJETA)
                rect = Rectangle(pos=fila.pos, size=fila.size)
            fila.bind(pos=lambda w, *_: setattr(rect, 'pos', w.pos),
                      size=lambda w, *_: setattr(rect, 'size', w.size))

            lbl_nombre = Label(text=nombre, font_size=dp(11), color=TEXTO,
                               halign='left', valign='middle', size_hint=(0.45, 1))
            lbl_nombre.bind(size=lbl_nombre.setter('text_size'))

            color_cant = ROJO if cantidad == 0 else TEXTO
            lbl_cant = Label(text=str(cantidad), font_size=dp(13), color=color_cant,
                             halign='center', valign='middle', size_hint=(0.1, 1), bold=True)
            lbl_cant.bind(size=lbl_cant.setter('text_size'))

            btn_foto = Button(text="📷", font_size=dp(18), size_hint=(0.2, 1),
                              background_normal='', background_color=(0.18, 0.22, 0.28, 1))
            btn_foto.bind(on_release=lambda _, fd=fila_data, bf=btn_foto: self._seleccionar_foto(fd, bf))

            inp_com = TextInput(hint_text="Comentario...", font_size=dp(10),
                                size_hint=(0.25, 1),
                                background_color=(0.15, 0.17, 0.22, 1),
                                foreground_color=TEXTO, multiline=False,
                                padding=[dp(6), dp(6)])
            inp_com.bind(text=lambda _, v, fd=fila_data: fd.update({"comentario": v}))

            fila.add_widget(lbl_nombre)
            fila.add_widget(lbl_cant)
            fila.add_widget(btn_foto)
            fila.add_widget(inp_com)

            self.lista.add_widget(fila)
            self.filas.append(fila_data)

    def _seleccionar_foto(self, fila_data, btn):
        root = tk.Tk()
        root.withdraw()
        ruta = filedialog.askopenfilename(
            title="Seleccionar foto",
            filetypes=[("Imágenes", "*.jpg *.jpeg *.png")])
        root.destroy()
        if ruta:
            with open(ruta, "rb") as f:
                fila_data["foto_b64"] = base64.b64encode(f.read()).decode()
            btn.text = "✓"
            btn.background_color = VERDE

    def _enviar_todos(self, *_):
        self.lbl_status.text = "Enviando..."
        def worker():
            enviados = 0
            for fd in self.filas:
                payload = dict(
                    tienda=self.tienda,
                    producto=fd["nombre"],
                    comentario=fd["comentario"],
                    foto=fd["foto_b64"],
                    usuario=USUARIO_ACTIVO.get("nombre", "")
                )
                try:
                    r = requests.post(f"{API}/api/reporte", json=payload, timeout=10)
                    if r.json().get("ok"):
                        enviados += 1
                except:
                    pass
            Clock.schedule_once(lambda _: self._resultado_envio(enviados), 0)
        threading.Thread(target=worker, daemon=True).start()

    def _resultado_envio(self, enviados):
        total = len(self.filas)
        self.lbl_status.text = f"✓ {enviados}/{total} reportes enviados"
        self.lbl_status.color = VERDE

    def _volver(self):
        self.manager.transition = SlideTransition(direction='right')
        self.manager.current = 'tiendas'

class DisplaysApp(App):
    def build(self):
        Window.clearcolor = FONDO
        sm = ScreenManager()
        sm.add_widget(LoginScreen(name='login'))
        sm.add_widget(TiendasScreen(name='tiendas'))
        sm.add_widget(ProductosScreen(name='productos'))
        sm.add_widget(ReportesScreen(name='reportes'))
        return sm

if __name__ == '__main__':
    DisplaysApp().run()
