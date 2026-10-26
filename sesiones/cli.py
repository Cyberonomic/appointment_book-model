"""Línea de comandos del modelo de agendas y expedientes."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from sesiones.agenda import (
    COLUMNAS_CERRADAS,
    COLUMNAS_PUBLICAS,
    COLUMNAS_SLOTS_CERRADAS,
    COLUMNAS_SLOTS_INTERNAS,
    COLUMNAS_SLOTS_PUBLICAS,
    ErrorDeSesiones,
    escribir_clasificacion_xlsx,
    escribir_csv,
    filas_revision,
    procesar,
    procesar_clasificacion,
)
from sesiones.bloques import mes
from sesiones.expedientes import (
    DEFAULT_HOJA_2025,
    DEFAULT_HOJA_2026,
    cargar_libro,
    leer_expediente_2025,
    leer_expediente_2026,
)
from sesiones.exportar import (
    COLUMNAS_EVENTOS,
    COLUMNAS_PERSONAS,
    exportar_eventos_2025_publicos,
    exportar_personas_2026_publicas,
    escribir_publicos,
)
from sesiones.reglas import ErrorDeModelo, cargar_textos_excluidos, configurar_folio_re


def _configurar_reglas(args: argparse.Namespace) -> None:
    if getattr(args, "folio_pattern", None):
        configurar_folio_re(args.folio_pattern)
    if getattr(args, "excluded_phrases", None):
        cargar_textos_excluidos(args.excluded_phrases)


def _parser_reglas(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--folio-pattern",
        help="Regex with two groups (prefix, digits) for case identifiers in cell text",
    )
    parser.add_argument(
        "--excluded-phrases",
        type=Path,
        help="Local file: one internal-work phrase per line (appended to built-in list)",
    )


def _parser_agenda(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--workbook", required=True, type=Path, help="Excel de la cartelera")
    parser.add_argument("--sites", required=True, type=Path, help="Catálogo de sedes, una por línea")
    parser.add_argument("--variants", type=Path, help="Variantes ORIGEN=DESTINO")
    parser.add_argument("--month", required=True, help="oct2025, nov2025, dic2025, ene2026, feb2026, ...")
    parser.add_argument("--output", required=True, type=Path, help="CSV de salida")
    parser.add_argument(
        "--closed",
        action="store_true",
        help="Escribe el CSV de octubre/noviembre ya publicado (sin folio ni layout)",
    )
    parser.add_argument(
        "--keep-internal",
        action="store_true",
        help="No aplica el filtro de texto interno (solo para reproducir los CSV cerrados)",
    )
    parser.add_argument(
        "--clasificacion",
        type=Path,
        help="Excel interno con todos los slots clasificados (incluye nombres y folios)",
    )
    parser.add_argument(
        "--revision",
        type=Path,
        help="Excel interno sin_clasificar y filas con flags (default: revision.xlsx junto a --clasificacion)",
    )
    parser.add_argument(
        "--clasificacion-publica",
        type=Path,
        help="CSV anonimizado de slots clasificados (mismas reglas de alias que la salida pública de sesiones)",
    )
    parser.add_argument(
        "--incluir-vacios",
        action="store_true",
        help="Incluye celdas vacías en la clasificación",
    )
    _parser_reglas(parser)


def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Extrae sesiones de agenda o eventos de expediente. Una fila de agenda no es una persona."
    )
    sub = parser.add_subparsers(dest="comando")

    agenda = sub.add_parser("agenda", help="Cartelera semanal → sesiones efectuadas")
    _parser_agenda(agenda)

    ev = sub.add_parser("eventos-2025", help="Formulario 2025 → eventos públicos")
    ev.add_argument("--workbook", required=True, type=Path)
    ev.add_argument("--output", required=True, type=Path)
    ev.add_argument("--sheet", default=DEFAULT_HOJA_2025, help=f"Worksheet name (default: {DEFAULT_HOJA_2025})")
    _parser_reglas(ev)

    per = sub.add_parser("personas-2026", help="Base 2026 → personas públicas")
    per.add_argument("--workbook", required=True, type=Path)
    per.add_argument("--output", required=True, type=Path)
    per.add_argument("--sheet", default=DEFAULT_HOJA_2026, help=f"Worksheet name (default: {DEFAULT_HOJA_2026})")
    _parser_reglas(per)
    return parser


def _agenda(args: argparse.Namespace) -> int:
    _configurar_reglas(args)
    layout, bloques = mes(args.month)
    resultado = procesar(
        args.workbook,
        args.sites,
        bloques,
        layout,
        args.variants,
        excluir_interno=not args.keep_internal,
        cerradas=args.closed,
    )
    columnas = COLUMNAS_CERRADAS if args.closed else COLUMNAS_PUBLICAS
    escribir_csv(args.output, resultado.filas, columnas)
    print(resultado.informe)
    print(f"escrito {args.output}")

    if args.clasificacion:
        clasificado = procesar_clasificacion(
            args.workbook,
            args.sites,
            bloques,
            layout,
            args.variants,
            incluir_vacios=args.incluir_vacios,
            cerradas=args.closed,
        )
        escribir_clasificacion_xlsx(
            args.clasificacion,
            clasificado.filas_internas,
            COLUMNAS_SLOTS_INTERNAS,
        )
        ruta_revision = args.revision or (args.clasificacion.parent / "revision.xlsx")
        escribir_clasificacion_xlsx(
            ruta_revision,
            filas_revision(clasificado.filas_internas, clasificado.filas_hora_invalida),
            COLUMNAS_SLOTS_INTERNAS,
        )
        print(clasificado.informe_clasificacion())
        print(f"escrito {args.clasificacion}")
        print(f"escrito {ruta_revision}")
        if args.clasificacion_publica:
            cols_pub = COLUMNAS_SLOTS_CERRADAS if args.closed else COLUMNAS_SLOTS_PUBLICAS
            escribir_csv(args.clasificacion_publica, clasificado.filas_publicas, cols_pub)
            print(f"escrito {args.clasificacion_publica}")

    return 0


def _eventos(args: argparse.Namespace) -> int:
    _configurar_reglas(args)
    libro = cargar_libro(args.workbook)
    try:
        filas = exportar_eventos_2025_publicos(leer_expediente_2025(libro, hoja=args.sheet))
    finally:
        libro.close()
    escribir_publicos(args.output, filas, COLUMNAS_EVENTOS)
    print(f"{len(filas)} eventos")
    print(f"escrito {args.output}")
    return 0


def _personas(args: argparse.Namespace) -> int:
    _configurar_reglas(args)
    libro = cargar_libro(args.workbook)
    try:
        filas = exportar_personas_2026_publicas(leer_expediente_2026(libro, hoja=args.sheet))
    finally:
        libro.close()
    escribir_publicos(args.output, filas, COLUMNAS_PERSONAS)
    print(f"{len(filas)} personas")
    print(f"escrito {args.output}")
    return 0


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0].startswith("--"):
        argv = ["agenda", *argv]
    args = construir_parser().parse_args(argv)
    if not args.comando:
        construir_parser().print_help()
        return 2
    try:
        if args.comando == "agenda":
            return _agenda(args)
        if args.comando == "eventos-2025":
            return _eventos(args)
        if args.comando == "personas-2026":
            return _personas(args)
    except (ErrorDeSesiones, ErrorDeModelo, KeyError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    return 2
