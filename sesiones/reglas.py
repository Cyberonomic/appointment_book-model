"""Catálogos y reglas del modelo."""

from __future__ import annotations

import re
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl.styles.colors import COLOR_INDEX

VERDE_ASISTE = "6AA84F"
VERDE_ASISTE_ALT = "00FF00"
VERDE_ASISTE_TEMA = "34A853"

_THEME_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
_ORDEN_TEMA = (
    "lt1",
    "dk1",
    "lt2",
    "dk2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
)

# Variantes de ejemplo para pruebas de normalización. En producción, usa --variants (archivo local).
VARIANTES_SEDE: dict[str, str] = {
    "SITE_A_TYPO1": "SITE_A",
    "SITE_A_TYPO2": "SITE_A",
    "SITE_B_EXTENDED": "SITE_B",
    "SITE_B_ALIAS": "SITE_B",
}

# Sedes reales solo en el catálogo local (--sites); no se versionan en el repo.
SEDES_BASE: tuple[str, ...] = ()

# Short generic phrases shipped with the repo. Add institution-specific lines via --excluded-phrases.
TEXTOS_EXCLUIDOS_BASE: tuple[str, ...] = (
    "meeting",
    "break",
    "total",
    "lunch",
    "workshop",
)

_textos_excluidos: tuple[str, ...] = TEXTOS_EXCLUIDOS_BASE

ESTADOS = (
    "vacio",
    "asiste",
    "reagendo",
    "inasistencia",
    "egreso",
    "baja",
    "ingreso",
    "no_sesion",
    "sin_clasificar",
)

# Relleno semántico (últimos 6 hex del RGB resuelto).
COLOR_A_ESTADO: dict[str, str] = {
    VERDE_ASISTE: "asiste",
    VERDE_ASISTE_ALT: "asiste",
    VERDE_ASISTE_TEMA: "asiste",
    "93C47D": "asiste",
    "38761D": "asiste",
    "FFFF00": "reagendo",
    "FF0000": "inasistencia",
    "1155CC": "egreso",
    "0000FF": "egreso",
    "ED7D31": "baja",
    "E69138": "baja",
    "9900FF": "baja",
    "FF6D01": "baja",
    "FFC000": "ingreso",
    "FF00FF": "no_sesion",
}

# Flags adicionales por color (p. ej. aproximación respecto a la leyenda).
COLOR_FLAGS: dict[str, tuple[str, ...]] = {
    "FF6D01": ("color_aproximado",),
}

# Fondos de semana/día/rejilla: no implican estado; solo palabras clave en el texto.
COLORES_FONDO_CLAROS: frozenset[str] = frozenset(
    {
        "D9EAD3",
        "FFF2CC",
        "F4CCCC",
        "CFE2F3",
        "E6B8AF",
        "FCE5CD",
        "D9D2E9",
        "FFFFFF",
        "F3F3F3",
        "D8D8D8",
        "D9D9D9",
        "CCCCCC",
        "B7B7B7",
        "B6D7A8",
        "D7F1FD",
        "C9DAF8",
        "9FC5E8",
        "D0E0E3",
        "EAD1DC",
        "FFE599",
        "FFD966",
        "F9CB9C",
        "EA9999",
        "DD7E6B",
        "B4A7D6",
        "C2B795",
        "E06666",
    }
)

# Catálogo documentado (escaneo + reglas) para comparar con carteleras nuevas.
COLORES_CATALOGO: frozenset[str] = frozenset(COLOR_A_ESTADO) | COLORES_FONDO_CLAROS

_RE_NOMBRE = re.compile(r"\b[A-ZÁÉÍÓÚÑ][A-Za-záéíóúñ]+")
_PALABRAS_REAGENDO_SOSPECHOSAS = (
    "cons",
    "seguimiento",
    "capacitacion",
    "junta",
    "taller",
    "reunion",
    "revision",
    "supervision",
    "comida",
    "descanso",
    "totales",
    "vacaciones",
    "festivo",
)

# Optional letter prefix, optional IU/IF/IC marker, then digits. Override with --folio-pattern.
FOLIO_PATTERN_DEFAULT = (
    r"\b("
    r"[A-Z]{0,8}(?:IU|IF|IC)|"
    r"IU|IF"
    r")\s*-?\s*(\d{2,6})\b"
)


def configurar_folio_re(patron: str | None = None) -> None:
    """Compile the folio extractor regex (CLI / config override)."""
    global FOLIO_RE
    FOLIO_RE = re.compile(patron or FOLIO_PATTERN_DEFAULT, re.IGNORECASE)


configurar_folio_re()


@dataclass(frozen=True)
class ClasificacionSlot:
    estado: str
    motivo: str = ""
    flags: tuple[str, ...] = field(default_factory=tuple)


class ErrorDeModelo(Exception):
    """El insumo no se puede convertir sin ambigüedad."""

    def __init__(self, mensajes: str | list[str]):
        if isinstance(mensajes, str):
            mensajes = [mensajes]
        self.mensajes = list(mensajes)
        super().__init__("\n".join(self.mensajes))


def plegar(texto: str) -> str:
    plano = unicodedata.normalize("NFD", str(texto).replace("\n", " "))
    sin_acento = "".join(c for c in plano if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", sin_acento).upper().strip()


def _compacto(texto: str) -> str:
    return re.sub(r"\s+", "", plegar(texto))


def normalizar_sede(texto: str | None, extra: dict[str, str] | None = None) -> str | None:
    """Devuelve la sede canónica o None si no hay texto."""
    if texto is None or not str(texto).strip():
        return None
    clave = plegar(texto)
    mapa = {plegar(k): v for k, v in VARIANTES_SEDE.items()}
    if extra:
        mapa.update({plegar(k): v for k, v in extra.items()})
    for destino in {plegar(v) for v in mapa.values()}:
        mapa.setdefault(destino, destino)
    if clave in mapa:
        return mapa[clave]
    for origen, destino in sorted(mapa.items(), key=lambda par: len(par[0]), reverse=True):
        if origen and origen in clave:
            return destino
    return clave


def partir_encabezado(
    encabezado: str | None,
    sedes: tuple[str, ...],
    extra: dict[str, str] | None = None,
) -> tuple[str, str | None]:
    """Separa terapeuta y sede. La sede más larga gana."""
    texto = (encabezado or "").strip()
    if not texto:
        return "", None
    for sede in sedes:
        coincidencia = re.search(re.escape(sede) + r"\s*$", texto, flags=re.IGNORECASE)
        if coincidencia is None:
            continue
        return texto[: coincidencia.start()].strip(), normalizar_sede(sede, extra)
    return texto, None


def cargar_textos_excluidos(ruta: Path | None) -> None:
    """Append phrases from a local file (one per line) to the built-in generic list."""
    global _textos_excluidos
    extra = tuple(leer_lineas(ruta)) if ruta is not None and ruta.is_file() else ()
    _textos_excluidos = TEXTOS_EXCLUIDOS_BASE + extra


def reiniciar_textos_excluidos() -> None:
    global _textos_excluidos
    _textos_excluidos = TEXTOS_EXCLUIDOS_BASE


def es_texto_excluido(texto: str | None) -> bool:
    if texto is None or not str(texto).strip():
        return False
    plegado = plegar(texto)
    compacto = _compacto(texto)
    if "VACACIONES" in compacto:
        return True
    return any(plegar(frase) in plegado for frase in _textos_excluidos)


def suffix_hex(color_hex: str | None) -> str | None:
    if not color_hex or not isinstance(color_hex, str):
        return None
    limpio = color_hex.upper().strip()
    if len(limpio) < 6:
        return limpio
    return limpio[-6:]


def _aplicar_tinta(hex6: str, tint: float) -> str:
    if not tint:
        return hex6.upper()
    r = int(hex6[0:2], 16)
    g = int(hex6[2:4], 16)
    b = int(hex6[4:6], 16)
    if tint < 0:
        factor = 1 + tint
        r, g, b = r * factor, g * factor, b * factor
    else:
        r = r + (255 - r) * tint
        g = g + (255 - g) * tint
        b = b + (255 - b) * tint
    clamp = lambda x: max(0, min(255, int(x + 0.5)))
    return f"{clamp(r):02X}{clamp(g):02X}{clamp(b):02X}"


def _paleta_tema(workbook) -> tuple[str | None, ...]:
    cache = getattr(workbook, "_sesiones_paleta_tema", None)
    if cache is not None:
        return cache
    theme_bytes = getattr(workbook, "loaded_theme", None)
    palette: list[str | None] = [None] * len(_ORDEN_TEMA)
    if theme_bytes:
        try:
            root = ET.fromstring(theme_bytes)
            scheme = root.find(".//a:clrScheme", _THEME_NS)
            if scheme is not None:
                for indice, nombre in enumerate(_ORDEN_TEMA):
                    nodo = scheme.find(f"a:{nombre}", _THEME_NS)
                    if nodo is None:
                        continue
                    srgb = nodo.find("a:srgbClr", _THEME_NS)
                    if srgb is not None and srgb.get("val"):
                        palette[indice] = srgb.get("val").upper()
                        continue
                    sys = nodo.find("a:sysClr", _THEME_NS)
                    if sys is not None and sys.get("lastClr"):
                        palette[indice] = sys.get("lastClr").upper()
        except ET.ParseError:
            pass
    workbook._sesiones_paleta_tema = tuple(palette)
    return workbook._sesiones_paleta_tema


def _rgb_desde_indexado(indice: int) -> str | None:
    if indice is None or indice < 0 or indice >= len(COLOR_INDEX):
        return None
    crudo = COLOR_INDEX[indice]
    if not isinstance(crudo, str) or len(crudo) < 6:
        return None
    return crudo.upper()


def resolver_color_celda(celda, workbook) -> tuple[str | None, bool]:
    """Devuelve (ARGB/RGB en mayúsculas, resuelto). resuelto=False → color_desconocido."""
    try:
        relleno = celda.fill
        if relleno is None or relleno.fill_type is None or relleno.fill_type == "none":
            return None, True
        fg = relleno.fgColor
        if fg is None:
            return None, True
        if fg.type == "rgb" and isinstance(fg.rgb, str):
            return fg.rgb.upper(), True
        if fg.type == "theme" and workbook is not None and fg.theme is not None:
            palette = _paleta_tema(workbook)
            if fg.theme < 0 or fg.theme >= len(palette):
                return None, False
            base = palette[fg.theme]
            if not base:
                return None, False
            hex6 = _aplicar_tinta(base, float(fg.tint or 0))
            return f"FF{hex6}", True
        if fg.type == "indexed" and fg.indexed is not None:
            rgb = _rgb_desde_indexado(int(fg.indexed))
            return (rgb, True) if rgb else (None, False)
    except (TypeError, AttributeError, ValueError):
        return None, False
    return None, False


def tipo_relleno(celda) -> str:
    """none | rgb | theme | indexed | other"""
    try:
        relleno = celda.fill
        if relleno is None or relleno.fill_type is None or relleno.fill_type == "none":
            return "none"
        fg = relleno.fgColor
        if fg is None:
            return "other"
        if fg.type == "rgb" and isinstance(fg.rgb, str):
            return "rgb"
        if fg.type == "theme":
            return "theme"
        if fg.type == "indexed":
            return "indexed"
        return "other"
    except (TypeError, AttributeError, ValueError):
        return "other"


def extraer_folio(texto: str | None) -> str | None:
    if texto is None:
        return None
    crudo = str(texto).strip()
    if not crudo:
        return None
    coincidencia = FOLIO_RE.search(crudo)
    if coincidencia is None:
        return None
    return f"{coincidencia.group(1).upper()}{coincidencia.group(2)}"


def extraer_iu_if(folio: str | None) -> str | None:
    if not folio:
        return None
    letras = re.sub(r"\d+", "", folio).upper()
    if "IU" in letras:
        return "IU"
    if "IF" in letras:
        return "IF"
    return None


def color_rgb(celda, workbook=None) -> str | None:
    rgb, resuelto = resolver_color_celda(celda, workbook)
    return rgb if resuelto else None


def _estado_por_texto_clave(texto: str) -> str | None:
    p = plegar(texto)
    if "(ALTA)" in p:
        return "egreso"
    if "(BAJA)" in p:
        return "baja"
    if "INASISTENCIA" in p or "NO ASISTE" in p:
        return "inasistencia"
    if "REAGEND" in p:
        return "reagendo"
    return None


def revisar_texto_reagendo(texto: str | None) -> bool:
    if texto is None or not str(texto).strip():
        return True
    p = plegar(texto)
    if any(palabra in p for palabra in _PALABRAS_REAGENDO_SOSPECHOSAS):
        return True
    if _RE_NOMBRE.search(str(texto)):
        return False
    return True


def _anexar_flags_texto(crudo: str, flags: list[str]) -> None:
    if "APOYO EN CURSO" in plegar(crudo) and "revisar_texto" not in flags:
        flags.append("revisar_texto")


def _anexar_flags_color(sufijo: str | None, flags: list[str]) -> None:
    if not sufijo:
        return
    for nombre in COLOR_FLAGS.get(sufijo, ()):
        if nombre not in flags:
            flags.append(nombre)


def _estado_por_sufijo(sufijo: str, layout: str | None) -> str | None:
    if sufijo == "FF9900":
        if layout == "horizontal":
            return "ingreso"
        if layout == "vertical":
            return "baja"
        return None
    return COLOR_A_ESTADO.get(sufijo)


def clasificar_slot(
    texto: str | None,
    color_hex: str | None,
    *,
    resuelto: bool = True,
    layout: str | None = None,
) -> ClasificacionSlot:
    """Clasifica un slot según texto y relleno (pasos 1–4 del modelo)."""
    crudo = "" if texto is None else str(texto).strip()
    if not crudo:
        return ClasificacionSlot("vacio")

    if es_texto_excluido(crudo):
        return ClasificacionSlot("no_sesion", motivo="texto_excluido")

    if not resuelto:
        return ClasificacionSlot("sin_clasificar", "color_desconocido")

    sufijo = suffix_hex(color_hex)
    flags: list[str] = []

    if sufijo in COLORES_FONDO_CLAROS:
        sugerido = _estado_por_texto_clave(crudo)
        if sugerido:
            if sugerido == "reagendo" and revisar_texto_reagendo(crudo):
                flags.append("revisar_texto")
            flags.append("estado_por_texto")
            _anexar_flags_texto(crudo, flags)
            return ClasificacionSlot(sugerido, "", tuple(flags))
        _anexar_flags_texto(crudo, flags)
        return ClasificacionSlot("sin_clasificar", "texto_sin_color", tuple(flags))

    if not sufijo:
        _anexar_flags_texto(crudo, flags)
        return ClasificacionSlot("sin_clasificar", "texto_sin_color", tuple(flags))

    estado = _estado_por_sufijo(sufijo, layout) if sufijo else None
    if estado:
        sugerido = _estado_por_texto_clave(crudo)
        if sugerido and sugerido != estado:
            flags.append("conflicto_color_texto")
        if estado == "reagendo" and revisar_texto_reagendo(crudo):
            flags.append("revisar_texto")
        _anexar_flags_color(sufijo, flags)
        _anexar_flags_texto(crudo, flags)
        return ClasificacionSlot(estado, "", tuple(flags))

    _anexar_flags_texto(crudo, flags)
    return ClasificacionSlot("sin_clasificar", "color_desconocido", tuple(flags))


def clasificar_celda(celda, workbook=None, layout: str | None = None) -> ClasificacionSlot:
    rgb, resuelto = resolver_color_celda(celda, workbook)
    return clasificar_slot(celda.value, rgb, resuelto=resuelto, layout=layout)


def es_asiste(celda) -> bool:
    """Regla C: texto no vacío y relleno que termina en 6AA84F."""
    if celda.value is None or str(celda.value).strip() == "":
        return False
    rgb = color_rgb(celda)
    return bool(rgb) and rgb.upper().endswith(VERDE_ASISTE)


def es_sesion(celda) -> bool:
    return es_asiste(celda) and not es_texto_excluido(celda.value)


def leer_lineas(ruta: Path) -> list[str]:
    if not ruta.exists():
        return []
    return [ln.strip() for ln in ruta.read_text(encoding="utf-8").splitlines() if ln.strip() and not ln.strip().startswith("#")]


def cargar_sedes(ruta: Path, ruta_variantes: Path | None = None) -> tuple[tuple[str, ...], dict[str, str]]:
    if not ruta.is_file():
        raise ErrorDeModelo(
            f"No está el catálogo de sedes: {ruta}. Una sede por línea. Ese archivo no se versiona."
        )
    vistos: dict[str, str] = {}
    errores: list[str] = []
    for numero, sede in enumerate(leer_lineas(ruta), start=1):
        clave = plegar(sede)
        if clave in vistos:
            errores.append(f"{ruta.name}: sede repetida ({sede})")
            continue
        vistos[clave] = sede.strip()
    if not vistos:
        errores.append(f"{ruta} no tiene sedes.")
    for sede in SEDES_BASE:
        vistos.setdefault(plegar(sede), sede)

    extra: dict[str, str] = {}
    if ruta_variantes is not None:
        if not ruta_variantes.is_file():
            raise ErrorDeModelo(
                f"No está el archivo de variantes: {ruta_variantes}. "
                "Cada línea es ORIGEN=DESTINO. Si no hay variantes, omite --variants."
            )
        for linea in leer_lineas(ruta_variantes):
            origen, separador, destino = linea.partition("=")
            origen, destino = origen.strip(), destino.strip()
            if not separador or not origen or not destino:
                errores.append(f"{ruta_variantes.name}: usa ORIGEN=DESTINO")
                continue
            if plegar(origen) not in vistos:
                errores.append(f"{ruta_variantes.name}: {origen} no está en el catálogo")
                continue
            extra[origen] = destino

    if errores:
        raise ErrorDeModelo(errores)
    ordenadas = tuple(sorted(vistos.values(), key=len, reverse=True))
    return ordenadas, extra


def codigo(mapa: dict[str, str], clave: str, prefijo: str, ancho: int) -> str:
    if clave not in mapa:
        mapa[clave] = f"{prefijo}-{len(mapa) + 1:0{ancho}d}"
    return mapa[clave]
