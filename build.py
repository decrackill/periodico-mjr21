import json
import hashlib
import shutil
from pathlib import Path
from datetime import datetime
from PIL import Image

ROOT = Path(__file__).parent
NUEVOS = ROOT / "nuevos"
DATA_FILE = ROOT / "data" / "articulos.json"
HASHES_FILE = ROOT / "data" / "hashes.json"
TEMPLATES = ROOT / "templates"
OUTPUT_ARTICULOS = ROOT / "articulos"
IMG_ARTICULOS = ROOT / "img" / "articulos"
PAGINA_DIR = ROOT / "pagina"
ANCHO_MAX = 1600
CALIDAD_WEBP = 80
EXTENSIONES_IMG = [".jpg", ".jpeg", ".png", ".webp"]
ARTICULOS_POR_PAGINA = 12


# ── Hashes (build incremental) ─────────────────────────────────────────────

def hash_carpeta(carpeta: Path) -> str:
    """MD5 del texto + metadatos de imagen fuente. Cambia solo si el contenido cambia."""
    h = hashlib.md5()
    txt = carpeta / "articulo.txt"
    if txt.exists():
        h.update(txt.read_bytes())
    for ext in EXTENSIONES_IMG:
        imgs = list(carpeta.glob(f"*{ext}"))
        if imgs:
            stat = imgs[0].stat()
            h.update(f"{imgs[0].name}:{stat.st_size}:{stat.st_mtime_ns}".encode())
            break
    return h.hexdigest()


def cargar_hashes() -> dict:
    if HASHES_FILE.exists():
        return json.loads(HASHES_FILE.read_text(encoding="utf-8"))
    return {}


def guardar_hashes(hashes: dict):
    HASHES_FILE.parent.mkdir(parents=True, exist_ok=True)
    HASHES_FILE.write_text(json.dumps(hashes, ensure_ascii=False, indent=2), encoding="utf-8")


# ── Parseo ─────────────────────────────────────────────────────────────────

def parsear_articulo(carpeta: Path) -> dict:
    texto = (carpeta / "articulo.txt").read_text(encoding="utf-8")
    metadata_texto, _, cuerpo_texto = texto.partition("---")

    metadata = {}
    for linea in metadata_texto.strip().splitlines():
        if ":" in linea:
            clave, _, valor = linea.partition(":")
            metadata[clave.strip().lower()] = valor.strip()

    cuerpo_html = cuerpo_texto.strip()

    imagen_origen = None
    for ext in EXTENSIONES_IMG:
        encontradas = list(carpeta.glob(f"*{ext}"))
        if encontradas:
            imagen_origen = encontradas[0]
            break

    return {
        "slug": carpeta.name,
        "titulo": metadata.get("titulo", ""),
        "categoria": metadata.get("categoria", ""),
        "autor": metadata.get("autor", ""),
        "fecha": metadata.get("fecha", ""),
        "resumen": metadata.get("resumen", ""),
        "cuerpo": cuerpo_html,
        "_imagen_origen": imagen_origen,
    }


def cargar_articulos() -> list:
    NUEVOS.mkdir(parents=True, exist_ok=True)
    articulos = []
    for carpeta in sorted(NUEVOS.iterdir()):
        if not carpeta.is_dir() or carpeta.name.startswith("_"):
            continue
        if not (carpeta / "articulo.txt").exists():
            print(f"  ADVERTENCIA: '{carpeta.name}' no tiene articulo.txt, se ignora")
            continue
        articulos.append(parsear_articulo(carpeta))
    articulos.sort(key=lambda a: a["fecha"], reverse=True)
    return articulos


# ── Imágenes ───────────────────────────────────────────────────────────────

def procesar_imagen(slug: str, origen) -> str:
    carpeta_destino = IMG_ARTICULOS / slug
    destino = carpeta_destino / "portada.webp"

    if origen is None:
        if destino.exists():
            return f"/img/articulos/{slug}/portada.webp"
        print(f"  ADVERTENCIA: sin foto para '{slug}'")
        return "/img/placeholder.webp"

    carpeta_destino.mkdir(parents=True, exist_ok=True)
    with Image.open(origen) as img:
        img = img.convert("RGB")
        if img.width > ANCHO_MAX:
            ratio = ANCHO_MAX / img.width
            img = img.resize((ANCHO_MAX, int(img.height * ratio)))
        img.save(destino, "WEBP", quality=CALIDAD_WEBP)

    print(f"  OK: imagen de '{slug}' procesada")
    return f"/img/articulos/{slug}/portada.webp"


def copiar_imagenes_inline(slug: str):
    """Copia nuevos/{slug}/img_inline/ → img/articulos/{slug}/img_inline/ si existe."""
    origen = NUEVOS / slug / "img_inline"
    if not origen.exists():
        return
    destino = IMG_ARTICULOS / slug / "img_inline"
    if destino.exists():
        shutil.rmtree(destino)
    shutil.copytree(origen, destino)
    print(f"  OK: imágenes inline de '{slug}' copiadas ({len(list(destino.iterdir()))} archivo/s)")


# ── Limpieza ───────────────────────────────────────────────────────────────

def limpiar_huerfanos(articulos: list, total_paginas: int):
    slugs_validos = {a["slug"] for a in articulos}

    for carpeta in OUTPUT_ARTICULOS.glob("*"):
        if carpeta.is_dir() and carpeta.name not in slugs_validos:
            shutil.rmtree(carpeta)
            print(f"  Eliminada carpeta huérfana: articulos/{carpeta.name}")

    for carpeta in IMG_ARTICULOS.glob("*"):
        if carpeta.is_dir() and carpeta.name not in slugs_validos:
            shutil.rmtree(carpeta)
            print(f"  Eliminada carpeta huérfana: img/articulos/{carpeta.name}")

    # Eliminar páginas de paginación que ya no existen
    if PAGINA_DIR.exists():
        for carpeta in PAGINA_DIR.iterdir():
            if carpeta.is_dir():
                try:
                    num = int(carpeta.name)
                    if num > total_paginas:
                        shutil.rmtree(carpeta)
                        print(f"  Eliminada página huérfana: pagina/{carpeta.name}/")
                except ValueError:
                    pass


# ── HTML helpers ───────────────────────────────────────────────────────────

def tarjeta_html(art: dict) -> str:
    return f"""
    <a href="/articulos/{art['slug']}/" class="news-card">
      <img src="{art['imagen']}" alt="{art['titulo']}" loading="lazy">
      <div class="news-card-body">
        <span class="tag-pill">{art['categoria']}</span>
        <h3>{art['titulo']}</h3>
        <p>{art['resumen']}</p>
        <span class="news-card-meta">{art['autor']} · {art['fecha']}</span>
      </div>
    </a>
    """


def destacado_html(art: dict) -> str:
    return f"""
    <a href="/articulos/{art['slug']}/" class="featured-card">
      <img src="{art['imagen']}" alt="{art['titulo']}">
      <div class="featured-body">
        <span class="tag-pill">{art['categoria']}</span>
        <h2>{art['titulo']}</h2>
        <p>{art['resumen']}</p>
        <span class="news-card-meta">{art['autor']} · {art['fecha']}</span>
      </div>
    </a>
    """


def relacionados_html(actual: dict, todos: list) -> str:
    relacionados = [
        a for a in todos
        if a["categoria"] == actual["categoria"] and a["slug"] != actual["slug"]
    ][:3]
    if not relacionados:
        return ""
    tarjetas = "\n".join(tarjeta_html(a) for a in relacionados)
    return f'<section class="related"><h3>Más en {actual["categoria"]}</h3><div class="news-grid">{tarjetas}</div></section>'


def paginacion_html(pagina_actual: int, total_paginas: int) -> str:
    if total_paginas <= 1:
        return ""
    partes = []
    if pagina_actual > 1:
        href = "/" if pagina_actual == 2 else f"/pagina/{pagina_actual - 1}/"
        partes.append(f'<a href="{href}" class="page-btn">← Anterior</a>')
    partes.append(f'<span class="page-info">Página {pagina_actual} de {total_paginas}</span>')
    if pagina_actual < total_paginas:
        partes.append(f'<a href="/pagina/{pagina_actual + 1}/" class="page-btn">Siguiente →</a>')
    return f'<nav class="pagination">{"".join(partes)}</nav>'


# ── Generadores ────────────────────────────────────────────────────────────

def generar_pagina_home(articulos_pagina: list, pagina: int, total_paginas: int):
    """Genera una página del home (index.html o pagina/N/index.html)."""
    plantilla = (TEMPLATES / "home.html").read_text(encoding="utf-8")
    anio = str(datetime.now().year)

    if pagina == 1:
        destacado = destacado_html(articulos_pagina[0]) if articulos_pagina else ""
        tarjetas = "\n".join(tarjeta_html(a) for a in articulos_pagina[1:])
        destino = ROOT / "index.html"
    else:
        destacado = ""
        tarjetas = "\n".join(tarjeta_html(a) for a in articulos_pagina)
        carpeta = PAGINA_DIR / str(pagina)
        carpeta.mkdir(parents=True, exist_ok=True)
        destino = carpeta / "index.html"

    html = plantilla.replace("{{DESTACADO}}", destacado)
    html = html.replace("{{TARJETAS}}", tarjetas)
    html = html.replace("{{PAGINACION}}", paginacion_html(pagina, total_paginas))
    html = html.replace("{{ANIO}}", anio)
    destino.write_text(html, encoding="utf-8")


def generar_home(articulos: list):
    """Divide artículos en páginas y genera todos los index."""
    total = len(articulos)
    total_paginas = max(1, -(-total // ARTICULOS_POR_PAGINA))  # ceil division

    for num_pagina in range(1, total_paginas + 1):
        inicio = (num_pagina - 1) * ARTICULOS_POR_PAGINA
        fin = inicio + ARTICULOS_POR_PAGINA
        generar_pagina_home(articulos[inicio:fin], num_pagina, total_paginas)

    print(f"  Home: {total_paginas} página(s) generada(s) ({ARTICULOS_POR_PAGINA} artículos/página)")
    return total_paginas


def generar_articulo(articulo: dict, todos: list):
    plantilla = (TEMPLATES / "articulo.html").read_text(encoding="utf-8")
    html = plantilla
    for campo in ["titulo", "categoria", "autor", "fecha", "imagen", "cuerpo"]:
        html = html.replace("{{" + campo.upper() + "}}", articulo[campo])
    html = html.replace("{{RELACIONADOS}}", relacionados_html(articulo, todos))
    html = html.replace("{{ANIO}}", str(datetime.now().year))
    carpeta = OUTPUT_ARTICULOS / articulo["slug"]
    carpeta.mkdir(parents=True, exist_ok=True)
    (carpeta / "index.html").write_text(html, encoding="utf-8")


def guardar_json(articulos: list):
    datos = [{k: v for k, v in a.items() if not k.startswith("_")} for a in articulos]
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    DATA_FILE.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")


# ── Main ───────────────────────────────────────────────────────────────────

def main():
    articulos = cargar_articulos()
    hashes_previos = cargar_hashes()
    hashes_nuevos = {}
    imagenes_procesadas = 0
    imagenes_saltadas = 0

    print("Procesando imágenes...")
    for art in articulos:
        slug = art["slug"]
        carpeta = NUEVOS / slug
        hash_actual = hash_carpeta(carpeta)
        hashes_nuevos[slug] = hash_actual

        portada_existe = (IMG_ARTICULOS / slug / "portada.webp").exists()
        sin_cambios = (hashes_previos.get(slug) == hash_actual) and portada_existe

        if sin_cambios:
            # Skip imagen — usar ruta ya existente
            art["imagen"] = f"/img/articulos/{slug}/portada.webp"
            imagenes_saltadas += 1
        else:
            art["imagen"] = procesar_imagen(slug, art.pop("_imagen_origen"))
            copiar_imagenes_inline(slug)
            if "_imagen_origen" in art:
                art.pop("_imagen_origen", None)
            imagenes_procesadas += 1

        # Asegurar que _imagen_origen no quede en el dict
        art.pop("_imagen_origen", None)

    if imagenes_saltadas:
        print(f"  Saltadas: {imagenes_saltadas} imagen(es) sin cambios")

    total_paginas = max(1, -(-len(articulos) // ARTICULOS_POR_PAGINA))
    limpiar_huerfanos(articulos, total_paginas)
    guardar_json(articulos)

    print("Generando HTML...")
    total_paginas = generar_home(articulos)
    for art in articulos:
        generar_articulo(art, articulos)

    guardar_hashes(hashes_nuevos)
    print(f"Listo: {len(articulos)} artículo(s) · {imagenes_procesadas} imagen(es) procesada(s) · {total_paginas} página(s) de home.")


if __name__ == "__main__":
    main()