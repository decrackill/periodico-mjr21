"""
Generador de Noticias — MJR-21
Lee un Word (.docx) con negrillas/cursivas/párrafos/tablas/imágenes y genera
la carpeta nuevos/{slug}/ lista, luego corre build.py.
Pestañas: Crear noticia | Editar noticia | Eliminar / deshacer
"""

import platform
import re
import shutil
import subprocess
import threading
import unicodedata
import zipfile
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).parent
NUEVOS = ROOT / "nuevos"
OUTPUT_ARTICULOS = ROOT / "articulos"
IMG_ARTICULOS = ROOT / "img" / "articulos"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


# ══════════════════════════════════════════════════════════════════════════════
#  UTILIDADES GENERALES
# ══════════════════════════════════════════════════════════════════════════════

def slugify(texto):
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    texto = texto.lower()
    texto = re.sub(r"[^a-z0-9\s-]", "", texto)
    texto = re.sub(r"\s+", "-", texto.strip())
    return texto[:60].strip("-")


def validar_fecha(texto):
    """Devuelve True si el texto es una fecha YYYY-MM-DD válida."""
    try:
        datetime.strptime(texto.strip(), "%Y-%m-%d")
        return True
    except ValueError:
        return False


def elegir_archivo(titulo, filtro_zenity, filtros_tk):
    if platform.system() == "Linux" and shutil.which("zenity"):
        try:
            resultado = subprocess.run(
                ["zenity", "--file-selection", "--title", titulo, "--file-filter", filtro_zenity],
                capture_output=True, text=True,
            )
            ruta = resultado.stdout.strip()
            return ruta if ruta else None
        except Exception:
            pass
    return filedialog.askopenfilename(title=titulo, filetypes=filtros_tk) or None


# ══════════════════════════════════════════════════════════════════════════════
#  CONVERSIÓN DOCX → HTML
# ══════════════════════════════════════════════════════════════════════════════

def runs_a_html(runs):
    partes = []
    for run in runs:
        texto = run.text
        if not texto:
            continue
        t = texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        if run.bold:
            t = f"<strong>{t}</strong>"
        if run.italic:
            t = f"<em>{t}</em>"
        if run.underline:
            t = f"<u>{t}</u>"
        partes.append(t)
    return "".join(partes).strip()


def parrafo_a_html(parrafo):
    return runs_a_html(parrafo.runs)


def tabla_a_filas(tabla, filas_pendientes):
    for i, fila in enumerate(tabla.rows):
        celdas = []
        for celda in fila.cells:
            contenido = " ".join(
                runs_a_html(p.runs) for p in celda.paragraphs if p.text.strip()
            )
            es_header = (i == 0 and not filas_pendientes)
            tc = "th" if es_header else "td"
            celdas.append(f"<{tc}>{contenido}</{tc}>")
        filas_pendientes.append(f"<tr>{''.join(celdas)}</tr>")


def extraer_imagenes_docx(ruta_docx, carpeta_destino):
    carpeta_img = carpeta_destino / "img_inline"
    carpeta_img.mkdir(parents=True, exist_ok=True)
    doc = Document(str(ruta_docx))
    rels = doc.part.rels
    mapa = {}
    with zipfile.ZipFile(str(ruta_docx)) as z:
        for rid, rel in rels.items():
            if "image" not in rel.reltype:
                continue
            zip_path = f"word/{rel.target_ref}"
            if zip_path not in z.namelist():
                continue
            ext = Path(rel.target_ref).suffix.lower() or ".png"
            nombre = f"{rid}{ext}"
            destino = carpeta_img / nombre
            with z.open(zip_path) as src, open(destino, "wb") as dst:
                dst.write(src.read())
            mapa[rid] = destino
    return mapa


def imagen_inline_html(parrafo_elem, mapa_rid, slug):
    blips = parrafo_elem.findall('.//' + qn('a:blip'))
    if not blips:
        return None
    rid = blips[0].get(qn('r:embed'))
    if not rid or rid not in mapa_rid:
        return None
    ruta_img = mapa_rid[rid]
    ruta_publica = f"/img/articulos/{slug}/img_inline/{ruta_img.name}"
    return f'<figure class="article-inline-img"><img src="{ruta_publica}" alt="Imagen del artículo" loading="lazy"></figure>'


def docx_a_cuerpo(ruta_docx, carpeta_slug, slug):
    doc = Document(str(ruta_docx))
    parrafo_map = {p._element: p for p in doc.paragraphs}
    tabla_map   = {t._element: t for t in doc.tables}

    mapa_rid = extraer_imagenes_docx(ruta_docx, carpeta_slug) if carpeta_slug else {}

    bloques = []
    primer_elemento = True
    filas_pendientes = []

    def volcar_tabla():
        if filas_pendientes:
            bloques.append(
                '<div class="table-wrap"><table>' + "".join(filas_pendientes) + '</table></div>'
            )
            filas_pendientes.clear()

    for child in doc.element.body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            img_html = imagen_inline_html(child, mapa_rid, slug)
            if img_html:
                volcar_tabla()
                bloques.append(img_html)
                primer_elemento = False
                continue

            parrafo = parrafo_map.get(child)
            if parrafo is None or not parrafo.text.strip():
                primer_elemento = False
                continue

            if primer_elemento:
                primer_elemento = False
                if parrafo.style.name.startswith("Heading") or parrafo.style.name == "Title":
                    continue

            volcar_tabla()
            html = parrafo_a_html(parrafo)
            if html:
                estilo = parrafo.style.name
                if estilo.startswith("Heading 1"):
                    bloques.append(f"<h2>{html}</h2>")
                elif estilo.startswith("Heading"):
                    bloques.append(f"<h3>{html}</h3>")
                else:
                    bloques.append(f"<p>{html}</p>")
            primer_elemento = False

        elif tag == "tbl":
            tabla = tabla_map.get(child)
            if tabla is not None:
                tabla_a_filas(tabla, filas_pendientes)
            primer_elemento = False
        else:
            primer_elemento = False

    volcar_tabla()
    return "\n\n".join(bloques)


def primer_titulo(ruta_docx):
    doc = Document(str(ruta_docx))
    for p in doc.paragraphs:
        if p.text.strip():
            if p.style.name.startswith("Heading") or p.style.name == "Title":
                return p.text.strip()
            break
    return ""


# ══════════════════════════════════════════════════════════════════════════════
#  UTILIDADES DE ARTÍCULOS
# ══════════════════════════════════════════════════════════════════════════════

def listar_articulos_existentes():
    resultado = []
    if not NUEVOS.exists():
        return resultado
    for carpeta in sorted(NUEVOS.iterdir()):
        if not carpeta.is_dir() or carpeta.name.startswith("_"):
            continue
        archivo_txt = carpeta / "articulo.txt"
        titulo = carpeta.name
        if archivo_txt.exists():
            for linea in archivo_txt.read_text(encoding="utf-8").splitlines():
                if linea.upper().startswith("TITULO:"):
                    titulo = linea.split(":", 1)[1].strip() or carpeta.name
                    break
        resultado.append((carpeta.name, titulo))
    return resultado


def leer_metadata_articulo(slug):
    """Lee el articulo.txt de un slug y devuelve un dict con los campos."""
    archivo = NUEVOS / slug / "articulo.txt"
    if not archivo.exists():
        return {}
    texto = archivo.read_text(encoding="utf-8")
    metadata_texto, _, _ = texto.partition("---")
    meta = {}
    for linea in metadata_texto.strip().splitlines():
        if ":" in linea:
            clave, _, valor = linea.partition(":")
            meta[clave.strip().lower()] = valor.strip()
    return meta


def eliminar_articulo(slug):
    for carpeta in [NUEVOS / slug, OUTPUT_ARTICULOS / slug, IMG_ARTICULOS / slug]:
        if carpeta.exists():
            shutil.rmtree(carpeta)


# ══════════════════════════════════════════════════════════════════════════════
#  VENTANA DE PREVISUALIZACIÓN
# ══════════════════════════════════════════════════════════════════════════════

PREVIEW_CSS = """
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family: Georgia, serif; background:#f4efe2; color:#2b2620;
       line-height:1.65; font-size:16px; padding: 2rem; }
h1 { font-size:1.8rem; margin-bottom:.5rem; font-family: sans-serif; }
h2 { font-size:1.3rem; margin:1.5rem 0 .4rem; font-family: sans-serif; }
h3 { font-size:1.1rem; margin:1.2rem 0 .3rem; font-family: sans-serif; }
p  { margin-bottom:1rem; }
.meta { color:#888; font-size:.85rem; margin-bottom:1.5rem; font-family:sans-serif; }
.tag { display:inline-block; background:#c0030e22; color:#8e0209;
       font-size:.75rem; padding:.2rem .7rem; border-radius:20px;
       font-family:sans-serif; margin-bottom:.8rem; }
img { max-width:100%; border-radius:8px; margin:1rem auto; display:block; }
figure { margin:1.5rem 0; text-align:center; }
table { width:100%; border-collapse:collapse; margin:1.2rem 0; font-size:.9rem; }
th { background:#0c4988; color:#fff; padding:.5rem .8rem; text-align:left; font-family:sans-serif; }
td { padding:.5rem .8rem; border-top:1px solid #ddd; }
tr:nth-child(even) td { background:rgba(0,0,0,.03); }
.table-wrap { overflow-x:auto; }
blockquote { border-left:4px solid #fab902; padding:.5rem 1rem;
             margin:1rem 0; color:#555; font-style:italic; }
</style>
"""


class VentanaPreview(ctk.CTkToplevel):
    def __init__(self, master, titulo, categoria, autor, fecha, resumen, cuerpo_html):
        super().__init__(master)
        self.title("Previsualización del artículo")
        self.geometry("820x700")
        self.resizable(True, True)
        self.resultado = False  # True = confirmar, False = cancelar

        # Header
        header = ctk.CTkFrame(self, fg_color="#0C4988", corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(
            header, text="👁  Previsualización — revisa antes de publicar",
            font=("Arial", 13, "bold"), text_color="white"
        ).pack(side="left", padx=16, pady=10)

        # Área de previsualización (HTML renderizado con tk.Text + tags)
        self._construir_vista(titulo, categoria, autor, fecha, resumen, cuerpo_html)

        # Botones
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", padx=20, pady=12)
        ctk.CTkButton(
            bar, text="✅ Confirmar y publicar", command=self._confirmar,
            fg_color="#0B7439", hover_color="#075226", font=("Arial", 13, "bold"), height=40
        ).pack(side="left", expand=True, fill="x", padx=(0, 8))
        ctk.CTkButton(
            bar, text="✏️ Volver a editar", command=self._cancelar,
            fg_color="#8E0209", hover_color="#5e0106", font=("Arial", 13), height=40
        ).pack(side="left", expand=True, fill="x")

        self.update_idletasks()
        self.after(0, self.grab_set)
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

    def _construir_vista(self, titulo, categoria, autor, fecha, resumen, cuerpo_html):
        import tkinter as tk

        frame = ctk.CTkFrame(self, fg_color="#f4efe2", corner_radius=0)
        frame.pack(fill="both", expand=True, padx=0, pady=0)

        sb = tk.Scrollbar(frame)
        sb.pack(side="right", fill="y")

        txt = tk.Text(
            frame, wrap="word", yscrollcommand=sb.set,
            bg="#f4efe2", fg="#2b2620", font=("Georgia", 13),
            relief="flat", padx=28, pady=20, cursor="arrow",
            state="normal"
        )
        txt.pack(fill="both", expand=True)
        sb.config(command=txt.yview)

        # Tags de estilo
        txt.tag_config("tag_pill", foreground="#8e0209", font=("Arial", 10), spacing1=4)
        txt.tag_config("h1",  font=("Arial", 22, "bold"), spacing1=10, spacing3=4)
        txt.tag_config("h2",  font=("Arial", 16, "bold"), spacing1=14, spacing3=3)
        txt.tag_config("h3",  font=("Arial", 14, "bold"), spacing1=10, spacing3=2)
        txt.tag_config("meta", foreground="#888888", font=("Arial", 10), spacing3=10)
        txt.tag_config("p",   font=("Georgia", 13), spacing3=8)
        txt.tag_config("resumen", font=("Georgia", 12, "italic"), foreground="#555555", spacing3=12)
        txt.tag_config("sep", font=("Arial", 6), spacing1=6, spacing3=6)

        # Contenido
        txt.insert("end", f"{categoria}\n", "tag_pill")
        txt.insert("end", f"{titulo}\n", "h1")
        txt.insert("end", f"{autor}  ·  {fecha}\n", "meta")
        if resumen:
            txt.insert("end", f"{resumen}\n", "resumen")
        txt.insert("end", "─" * 60 + "\n", "sep")

        # Parsear cuerpo HTML básico para mostrarlo legible
        self._renderizar_html(txt, cuerpo_html)

        txt.config(state="disabled")

    def _renderizar_html(self, txt, html):
        """Convierte el HTML del cuerpo a texto con tags de estilo en tk.Text."""
        import re as _re

        # Eliminar figure/img (no se pueden mostrar fácilmente en tk.Text)
        html = _re.sub(r'<figure[^>]*>.*?</figure>', '[📷 Imagen del Word]', html, flags=_re.DOTALL)

        # Tabla simplificada
        def reemplazar_tabla(m):
            contenido = m.group(0)
            celdas = _re.findall(r'<t[hd][^>]*>(.*?)</t[hd]>', contenido, _re.DOTALL)
            celdas_limpias = [_re.sub(r'<[^>]+>', '', c).strip() for c in celdas]
            return " | ".join(celdas_limpias)

        html = _re.sub(r'<div class="table-wrap">.*?</div>', reemplazar_tabla, html, flags=_re.DOTALL)

        # Procesar bloque a bloque
        bloques = html.split("\n\n")
        for bloque in bloques:
            bloque = bloque.strip()
            if not bloque:
                continue
            if bloque.startswith("<h2>"):
                texto = _re.sub(r'<[^>]+>', '', bloque)
                txt.insert("end", texto + "\n", "h2")
            elif bloque.startswith("<h3>"):
                texto = _re.sub(r'<[^>]+>', '', bloque)
                txt.insert("end", texto + "\n", "h3")
            else:
                texto = _re.sub(r'<[^>]+>', '', bloque)
                if texto:
                    txt.insert("end", texto + "\n", "p")

    def _confirmar(self):
        self.resultado = True
        self.grab_release()
        self.destroy()

    def _cancelar(self):
        self.resultado = False
        self.grab_release()
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  VENTANA DE LOG (build en tiempo real)
# ══════════════════════════════════════════════════════════════════════════════

class VentanaLog(ctk.CTkToplevel):
    def __init__(self, master, titulo="Construyendo el sitio..."):
        super().__init__(master)
        self.title(titulo)
        self.geometry("620x380")
        self.resizable(True, True)

        ctk.CTkLabel(self, text=titulo, font=("Arial", 13, "bold")).pack(pady=(14, 6), padx=16)

        self.barra = ctk.CTkProgressBar(self, mode="indeterminate", width=560)
        self.barra.pack(pady=(0, 10))
        self.barra.start()

        import tkinter as tk
        frame = ctk.CTkFrame(self, fg_color="#1a1a1a", corner_radius=8)
        frame.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        self._txt = tk.Text(
            frame, bg="#1a1a1a", fg="#00ff88", font=("Courier", 11),
            relief="flat", padx=10, pady=10, state="disabled", wrap="word"
        )
        sb = tk.Scrollbar(frame, command=self._txt.yview)
        self._txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._txt.pack(fill="both", expand=True)

        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", lambda: None)  # no cerrar manualmente

    def agregar(self, linea):
        self._txt.configure(state="normal")
        self._txt.insert("end", linea)
        self._txt.see("end")
        self._txt.configure(state="disabled")

    def finalizar(self, exito):
        self.barra.stop()
        self.barra.configure(mode="determinate")
        self.barra.set(1.0)
        if exito:
            self.barra.configure(progress_color="#0B7439")
        else:
            self.barra.configure(progress_color="#8E0209")
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        ctk.CTkButton(self, text="Cerrar", command=self.destroy, height=34).pack(pady=(0, 10))


# ══════════════════════════════════════════════════════════════════════════════
#  FUNCIÓN: correr build.py con log en tiempo real
# ══════════════════════════════════════════════════════════════════════════════

def correr_build_con_log(ventana_log, callback_fin):
    """Corre build.py en un hilo separado, manda output línea a línea al log."""
    def _hilo():
        proc = subprocess.Popen(
            ["python3", "build.py"],
            cwd=ROOT,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        )
        for linea in proc.stdout:
            ventana_log.after(0, ventana_log.agregar, linea)
        proc.wait()
        exito = proc.returncode == 0
        ventana_log.after(0, ventana_log.finalizar, exito)
        ventana_log.after(0, callback_fin, exito, proc.returncode)

    threading.Thread(target=_hilo, daemon=True).start()


def correr_build_silencioso():
    """Corre build.py sin ventana, devuelve CompletedProcess."""
    return subprocess.run(
        ["python3", "build.py"], cwd=ROOT, capture_output=True, text=True
    )


# ══════════════════════════════════════════════════════════════════════════════
#  APP PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Generador de Noticias — MJR-21")
        self.geometry("660x820")
        self.ultimo_slug = None

        ctk.CTkLabel(self, text="Generador de Noticias", font=("Arial", 20, "bold")).pack(pady=(20, 4))
        ctk.CTkLabel(self, text="MJR-21 — Departamento de Periodismo", text_color="gray").pack(pady=(0, 14))

        self.tabs = ctk.CTkTabview(self, width=600, height=700)
        self.tabs.pack(padx=20, pady=10, fill="both", expand=True)
        self.tabs.add("Crear noticia")
        self.tabs.add("Editar noticia")
        self.tabs.add("Eliminar / deshacer")

        self._construir_tab_crear(self.tabs.tab("Crear noticia"))
        self._construir_tab_editar(self.tabs.tab("Editar noticia"))
        self._construir_tab_eliminar(self.tabs.tab("Eliminar / deshacer"))

    # ── helpers compartidos ───────────────────────────────────────────────────

    def _campo(self, padre, placeholder, pady=5):
        entry = ctk.CTkEntry(padre, placeholder_text=placeholder, height=38)
        entry.pack(pady=pady, fill="x", padx=20)
        return entry

    def _seccion(self, padre, texto):
        ctk.CTkLabel(padre, text=texto, text_color="gray", font=("Arial", 11)).pack(
            anchor="w", padx=22, pady=(8, 0)
        )

    # ══════════════════════════════════════════════════════════════════════════
    #  PESTAÑA: CREAR NOTICIA
    # ══════════════════════════════════════════════════════════════════════════

    def _construir_tab_crear(self, padre):
        self.c_ruta_docx   = None
        self.c_ruta_imagen = None

        self.c_btn_docx = ctk.CTkButton(padre, text="📄 Elegir archivo Word (.docx)", command=self._c_elegir_docx)
        self.c_btn_docx.pack(pady=(14, 4), fill="x", padx=20)
        self.c_lbl_docx = ctk.CTkLabel(padre, text="Ningún archivo elegido", text_color="gray")
        self.c_lbl_docx.pack()

        self.c_btn_imagen = ctk.CTkButton(padre, text="🖼️ Elegir foto de portada", command=self._c_elegir_imagen)
        self.c_btn_imagen.pack(pady=(8, 4), fill="x", padx=20)
        self.c_lbl_imagen = ctk.CTkLabel(padre, text="Ninguna foto elegida", text_color="gray")
        self.c_lbl_imagen.pack()

        self.c_entry_titulo    = self._campo(padre, "Título")
        self.c_entry_categoria = self._campo(padre, "Categoría (ej. Comunicados, Análisis...)")
        self.c_entry_autor     = self._campo(padre, "Autor")
        self.c_entry_fecha     = self._campo(padre, "Fecha (YYYY-MM-DD)")
        self.c_entry_fecha.insert(0, date.today().isoformat())
        self.c_entry_resumen   = self._campo(padre, "Resumen (1-2 líneas)")

        self.c_lbl_slug = ctk.CTkLabel(padre, text="", text_color="gray", font=("Arial", 11))
        self.c_lbl_slug.pack(pady=(4, 2))
        self.c_lbl_fecha_error = ctk.CTkLabel(padre, text="", text_color="#CB030E", font=("Arial", 11))
        self.c_lbl_fecha_error.pack()

        self.c_entry_titulo.bind("<KeyRelease>", self._c_actualizar_slug)
        self.c_entry_fecha.bind("<KeyRelease>", self._c_validar_fecha)

        self.c_btn_previsualizar = ctk.CTkButton(
            padre, text="👁  Previsualizar antes de publicar",
            command=self._c_previsualizar,
            height=42, font=("Arial", 13, "bold"),
            fg_color="#0C4988", hover_color="#082F58"
        )
        self.c_btn_previsualizar.pack(pady=(14, 4), fill="x", padx=20)

        self.c_lbl_estado = ctk.CTkLabel(padre, text="", text_color="green", wraplength=540)
        self.c_lbl_estado.pack(pady=6)

    def _c_elegir_docx(self):
        ruta = elegir_archivo("Elegir noticia en Word", "Documentos Word | *.docx", [("Documentos Word", "*.docx")])
        if not ruta:
            return
        self.c_ruta_docx = Path(ruta)
        self.c_lbl_docx.configure(text=self.c_ruta_docx.name)
        titulo_sugerido = primer_titulo(self.c_ruta_docx)
        if titulo_sugerido and not self.c_entry_titulo.get():
            self.c_entry_titulo.insert(0, titulo_sugerido)
        self._c_actualizar_slug()

    def _c_elegir_imagen(self):
        ruta = elegir_archivo("Elegir foto de portada", "Imágenes | *.jpg *.jpeg *.png *.webp", [("Imágenes", "*.jpg *.jpeg *.png *.webp")])
        if not ruta:
            return
        self.c_ruta_imagen = Path(ruta)
        self.c_lbl_imagen.configure(text=self.c_ruta_imagen.name)

    def _c_actualizar_slug(self, *_):
        fecha  = self.c_entry_fecha.get().strip() or date.today().isoformat()
        titulo = self.c_entry_titulo.get().strip()
        slug   = f"{fecha}-{slugify(titulo)}" if titulo else fecha
        self.c_lbl_slug.configure(text=f"Carpeta: nuevos/{slug}/")

    def _c_validar_fecha(self, *_):
        texto = self.c_entry_fecha.get().strip()
        if texto and not validar_fecha(texto):
            self.c_lbl_fecha_error.configure(text="⚠ Formato inválido. Usa YYYY-MM-DD (ej. 2026-06-28)")
        else:
            self.c_lbl_fecha_error.configure(text="")
        self._c_actualizar_slug()

    def _c_previsualizar(self):
        """Valida, genera el HTML en memoria y abre la ventana de preview."""
        if not self.c_ruta_docx:
            messagebox.showerror("Falta el Word", "Elige primero el archivo .docx.")
            return
        titulo    = self.c_entry_titulo.get().strip()
        categoria = self.c_entry_categoria.get().strip()
        autor     = self.c_entry_autor.get().strip()
        fecha     = self.c_entry_fecha.get().strip()
        resumen   = self.c_entry_resumen.get().strip()

        if not titulo or not categoria or not autor or not fecha:
            messagebox.showerror("Faltan datos", "Completa título, categoría, autor y fecha.")
            return
        if not validar_fecha(fecha):
            messagebox.showerror("Fecha inválida", "Usa el formato YYYY-MM-DD (ej. 2026-06-28).")
            return

        # Generar HTML sin carpeta real (solo para preview, sin extraer imágenes)
        cuerpo_html = docx_a_cuerpo(self.c_ruta_docx, None, slugify(titulo))

        preview = VentanaPreview(self, titulo, categoria, autor, fecha, resumen, cuerpo_html)
        self.wait_window(preview)

        if preview.resultado:
            self._c_generar(titulo, categoria, autor, fecha, resumen)

    def _c_generar(self, titulo, categoria, autor, fecha, resumen):
        slug    = f"{fecha}-{slugify(titulo)}"
        carpeta = NUEVOS / slug

        if carpeta.exists():
            if not messagebox.askyesno("Ya existe", f"Ya existe '{slug}'. ¿Sobrescribir?"):
                return

        carpeta.mkdir(parents=True, exist_ok=True)
        cuerpo_html = docx_a_cuerpo(self.c_ruta_docx, carpeta, slug)

        contenido = (
            f"TITULO: {titulo}\n"
            f"CATEGORIA: {categoria}\n"
            f"AUTOR: {autor}\n"
            f"FECHA: {fecha}\n"
            f"RESUMEN: {resumen}\n"
            f"---\n"
            f"{cuerpo_html}\n"
        )
        (carpeta / "articulo.txt").write_text(contenido, encoding="utf-8")

        if self.c_ruta_imagen:
            shutil.copy(self.c_ruta_imagen, carpeta / f"foto{self.c_ruta_imagen.suffix.lower()}")

        self.ultimo_slug = slug
        self.c_btn_previsualizar.configure(state="disabled", text="⏳ Construyendo...")
        self.c_lbl_estado.configure(text="")

        log = VentanaLog(self, "Construyendo el sitio...")

        def al_terminar(exito, codigo):
            self.c_btn_previsualizar.configure(state="normal", text="👁  Previsualizar antes de publicar")
            if exito:
                self.c_lbl_estado.configure(
                    text=f"✅ '{slug}' publicada. Recarga el navegador.", text_color="green"
                )
                self._e_refrescar_lista()
                self._d_refrescar_lista()
            else:
                self.c_lbl_estado.configure(
                    text=f"⚠️ build.py terminó con error (código {codigo}).", text_color="red"
                )

        correr_build_con_log(log, al_terminar)

    # ══════════════════════════════════════════════════════════════════════════
    #  PESTAÑA: EDITAR NOTICIA
    # ══════════════════════════════════════════════════════════════════════════

    def _construir_tab_editar(self, padre):
        self.e_slug_actual  = None
        self.e_ruta_docx    = None
        self.e_ruta_imagen  = None

        ctk.CTkLabel(padre, text="Elige una noticia para editar sus datos o reemplazar archivos.",
                     text_color="gray", wraplength=520).pack(pady=(14, 8))

        self.e_menu = ctk.CTkOptionMenu(padre, values=["(sin noticias)"], command=self._e_cargar)
        self.e_menu.pack(fill="x", padx=20, pady=(0, 4))
        ctk.CTkButton(padre, text="🔄 Actualizar lista", command=self._e_refrescar_lista, height=30).pack(pady=(0, 10))

        # Reemplazar Word (opcional)
        self.e_btn_docx = ctk.CTkButton(padre, text="📄 Reemplazar Word (opcional)", command=self._e_elegir_docx)
        self.e_btn_docx.pack(fill="x", padx=20, pady=(0, 4))
        self.e_lbl_docx = ctk.CTkLabel(padre, text="Sin cambio — se conserva el contenido actual", text_color="gray", font=("Arial", 11))
        self.e_lbl_docx.pack()

        # Reemplazar foto (opcional)
        self.e_btn_imagen = ctk.CTkButton(padre, text="🖼️ Reemplazar foto de portada (opcional)", command=self._e_elegir_imagen)
        self.e_btn_imagen.pack(fill="x", padx=20, pady=(6, 4))
        self.e_lbl_imagen = ctk.CTkLabel(padre, text="Sin cambio — se conserva la foto actual", text_color="gray", font=("Arial", 11))
        self.e_lbl_imagen.pack()

        self.e_entry_titulo    = self._campo(padre, "Título")
        self.e_entry_categoria = self._campo(padre, "Categoría")
        self.e_entry_autor     = self._campo(padre, "Autor")
        self.e_entry_fecha     = self._campo(padre, "Fecha (YYYY-MM-DD)")
        self.e_entry_resumen   = self._campo(padre, "Resumen")

        self.e_lbl_fecha_error = ctk.CTkLabel(padre, text="", text_color="#CB030E", font=("Arial", 11))
        self.e_lbl_fecha_error.pack()
        self.e_entry_fecha.bind("<KeyRelease>", self._e_validar_fecha)

        self.e_btn_guardar = ctk.CTkButton(
            padre, text="💾 Guardar cambios y reconstruir",
            command=self._e_guardar,
            height=42, font=("Arial", 13, "bold"),
            fg_color="#0B7439", hover_color="#075226"
        )
        self.e_btn_guardar.pack(pady=14, fill="x", padx=20)

        self.e_lbl_estado = ctk.CTkLabel(padre, text="", text_color="green", wraplength=540)
        self.e_lbl_estado.pack(pady=4)

        self._e_refrescar_lista()

    def _e_refrescar_lista(self):
        self._e_mapa = {f"{t}  ({s})": s for s, t in listar_articulos_existentes()}
        opciones = list(self._e_mapa.keys()) or ["(sin noticias)"]
        self.e_menu.configure(values=opciones)
        self.e_menu.set(opciones[0])
        if opciones[0] != "(sin noticias)":
            self._e_cargar(opciones[0])

    def _e_cargar(self, seleccion):
        slug = self._e_mapa.get(seleccion)
        if not slug:
            return
        self.e_slug_actual = slug
        self.e_ruta_docx   = None
        self.e_ruta_imagen = None
        self.e_lbl_docx.configure(text="Sin cambio — se conserva el contenido actual")
        self.e_lbl_imagen.configure(text="Sin cambio — se conserva la foto actual")

        meta = leer_metadata_articulo(slug)
        for entry, clave in [
            (self.e_entry_titulo,    "titulo"),
            (self.e_entry_categoria, "categoria"),
            (self.e_entry_autor,     "autor"),
            (self.e_entry_fecha,     "fecha"),
            (self.e_entry_resumen,   "resumen"),
        ]:
            entry.delete(0, "end")
            entry.insert(0, meta.get(clave, ""))

        self.e_lbl_estado.configure(text="")
        self.e_lbl_fecha_error.configure(text="")

    def _e_elegir_docx(self):
        ruta = elegir_archivo("Elegir nuevo Word", "Documentos Word | *.docx", [("Documentos Word", "*.docx")])
        if not ruta:
            return
        self.e_ruta_docx = Path(ruta)
        self.e_lbl_docx.configure(text=f"Nuevo Word: {self.e_ruta_docx.name}")

    def _e_elegir_imagen(self):
        ruta = elegir_archivo("Elegir nueva foto", "Imágenes | *.jpg *.jpeg *.png *.webp", [("Imágenes", "*.jpg *.jpeg *.png *.webp")])
        if not ruta:
            return
        self.e_ruta_imagen = Path(ruta)
        self.e_lbl_imagen.configure(text=f"Nueva foto: {self.e_ruta_imagen.name}")

    def _e_validar_fecha(self, *_):
        texto = self.e_entry_fecha.get().strip()
        if texto and not validar_fecha(texto):
            self.e_lbl_fecha_error.configure(text="⚠ Formato inválido. Usa YYYY-MM-DD")
        else:
            self.e_lbl_fecha_error.configure(text="")

    def _e_guardar(self):
        if not self.e_slug_actual:
            messagebox.showinfo("Nada seleccionado", "Elige primero una noticia de la lista.")
            return

        titulo    = self.e_entry_titulo.get().strip()
        categoria = self.e_entry_categoria.get().strip()
        autor     = self.e_entry_autor.get().strip()
        fecha     = self.e_entry_fecha.get().strip()
        resumen   = self.e_entry_resumen.get().strip()

        if not titulo or not categoria or not autor or not fecha:
            messagebox.showerror("Faltan datos", "Completa todos los campos.")
            return
        if not validar_fecha(fecha):
            messagebox.showerror("Fecha inválida", "Usa el formato YYYY-MM-DD.")
            return

        carpeta = NUEVOS / self.e_slug_actual

        # Regenerar cuerpo si se eligió nuevo Word
        if self.e_ruta_docx:
            # Limpiar imágenes inline anteriores
            inline_vieja = carpeta / "img_inline"
            if inline_vieja.exists():
                shutil.rmtree(inline_vieja)
            cuerpo_html = docx_a_cuerpo(self.e_ruta_docx, carpeta, self.e_slug_actual)
        else:
            # Conservar el cuerpo actual del articulo.txt
            txt_actual = (carpeta / "articulo.txt").read_text(encoding="utf-8")
            _, _, cuerpo_html = txt_actual.partition("---")
            cuerpo_html = cuerpo_html.strip()

        contenido = (
            f"TITULO: {titulo}\n"
            f"CATEGORIA: {categoria}\n"
            f"AUTOR: {autor}\n"
            f"FECHA: {fecha}\n"
            f"RESUMEN: {resumen}\n"
            f"---\n"
            f"{cuerpo_html}\n"
        )
        (carpeta / "articulo.txt").write_text(contenido, encoding="utf-8")

        # Reemplazar foto de portada si se eligió una nueva
        if self.e_ruta_imagen:
            for vieja in carpeta.glob("foto.*"):
                vieja.unlink()
            shutil.copy(self.e_ruta_imagen, carpeta / f"foto{self.e_ruta_imagen.suffix.lower()}")

        self.e_btn_guardar.configure(state="disabled", text="⏳ Construyendo...")

        log = VentanaLog(self, "Aplicando cambios...")

        def al_terminar(exito, codigo):
            self.e_btn_guardar.configure(state="normal", text="💾 Guardar cambios y reconstruir")
            if exito:
                self.e_lbl_estado.configure(
                    text=f"✅ '{self.e_slug_actual}' actualizada. Recarga el navegador.", text_color="green"
                )
                self._e_refrescar_lista()
                self._d_refrescar_lista()
            else:
                self.e_lbl_estado.configure(
                    text=f"⚠️ build.py falló (código {codigo}).", text_color="red"
                )

        correr_build_con_log(log, al_terminar)

    # ══════════════════════════════════════════════════════════════════════════
    #  PESTAÑA: ELIMINAR / DESHACER
    # ══════════════════════════════════════════════════════════════════════════

    def _construir_tab_eliminar(self, padre):
        ctk.CTkLabel(
            padre, text="Elimina una noticia existente, o deshaz la última que creaste.",
            text_color="gray", wraplength=520
        ).pack(pady=(14, 10))

        self.d_lbl_ultimo = ctk.CTkLabel(padre, text="Última noticia creada en esta sesión: ninguna", wraplength=520)
        self.d_lbl_ultimo.pack(pady=(0, 6))

        self.d_btn_deshacer = ctk.CTkButton(
            padre, text="↩️ Deshacer la última noticia creada",
            command=self._d_deshacer, fg_color="#B98C02", hover_color="#8a6800"
        )
        self.d_btn_deshacer.pack(pady=10, fill="x", padx=20)

        ctk.CTkLabel(padre, text="— o elige cualquier noticia para eliminarla —", text_color="gray").pack(pady=(16, 6))

        self.d_menu = ctk.CTkOptionMenu(padre, values=["(sin noticias)"])
        self.d_menu.pack(pady=6, fill="x", padx=20)
        ctk.CTkButton(padre, text="🔄 Actualizar lista", command=self._d_refrescar_lista, height=30).pack(pady=4)

        self.d_btn_eliminar = ctk.CTkButton(
            padre, text="🗑️ Eliminar la noticia seleccionada",
            command=self._d_eliminar, fg_color="#8E0209", hover_color="#5e0106"
        )
        self.d_btn_eliminar.pack(pady=14, fill="x", padx=20)

        self.d_lbl_estado = ctk.CTkLabel(padre, text="", text_color="green", wraplength=520)
        self.d_lbl_estado.pack(pady=6)

        self._d_refrescar_lista()

    def _d_refrescar_lista(self):
        self._d_mapa = {f"{t}  ({s})": s for s, t in listar_articulos_existentes()}
        opciones = list(self._d_mapa.keys()) or ["(sin noticias)"]
        self.d_menu.configure(values=opciones)
        self.d_menu.set(opciones[0])
        if self.ultimo_slug:
            self.d_lbl_ultimo.configure(text=f"Última noticia creada: {self.ultimo_slug}")
        else:
            self.d_lbl_ultimo.configure(text="Última noticia creada en esta sesión: ninguna")

    def _d_deshacer(self):
        if not self.ultimo_slug:
            messagebox.showinfo("Nada que deshacer", "No has creado ninguna noticia en esta sesión.")
            return
        if not messagebox.askyesno("Confirmar", f"¿Eliminar '{self.ultimo_slug}'?"):
            return
        eliminar_articulo(self.ultimo_slug)
        slug = self.ultimo_slug
        self.ultimo_slug = None
        self._d_ejecutar_build_eliminar(slug)

    def _d_eliminar(self):
        seleccion = self.d_menu.get()
        slug = self._d_mapa.get(seleccion)
        if not slug:
            messagebox.showinfo("Nada seleccionado", "No hay ninguna noticia para eliminar.")
            return
        if not messagebox.askyesno("Confirmar eliminación", f"¿Eliminar permanentemente '{seleccion}'?\n\nEsto no se puede deshacer."):
            return
        eliminar_articulo(slug)
        if self.ultimo_slug == slug:
            self.ultimo_slug = None
        self._d_ejecutar_build_eliminar(slug)

    def _d_ejecutar_build_eliminar(self, slug):
        self.d_btn_eliminar.configure(state="disabled")
        log = VentanaLog(self, "Reconstruyendo tras eliminación...")

        def al_terminar(exito, codigo):
            self.d_btn_eliminar.configure(state="normal")
            if exito:
                self.d_lbl_estado.configure(
                    text=f"✅ '{slug}' eliminada y sitio reconstruido. Recarga el navegador.", text_color="green"
                )
            else:
                self.d_lbl_estado.configure(
                    text=f"⚠️ Se borró la carpeta pero build.py falló (código {codigo}).", text_color="red"
                )
            self._d_refrescar_lista()
            self._e_refrescar_lista()

        correr_build_con_log(log, al_terminar)


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    NUEVOS.mkdir(parents=True, exist_ok=True)
    app = App()
    app.mainloop()