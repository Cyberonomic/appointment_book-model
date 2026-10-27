"""Salidas públicas y el cruce interno por ID."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from sesiones.agenda import COLUMNAS_PUBLICAS, SesionAgenda, escribir_csv, exportar_sesiones_publicas
from sesiones.expedientes import Evento2025, Persona2026
from sesiones.reglas import extraer_iu_if


COLUMNAS_EVENTOS = ["evento_id", "fecha", "folio", "tipo", "estatus", "sexo", "rango_edad", "sustancia"]
COLUMNAS_PERSONAS = [
    "folio",
    "iu_if",
    "edad",
    "genero",
    "sustancia",
    "total_sesiones",
    "estatus_tratamiento",
    "motivo_egreso",
    *[f"sesiones_{m:02d}" for m in range(1, 13)],
]


@dataclass(frozen=True)
class PersonaUnificada:
    folio: str
    origen: str
    terapeuta_ultimo: str | None
    sede: str | None
    genero: str | None
    edad_o_rango: str | None
    sustancia: str | None
    iu_if: str | None
    total_sesiones: int
    desenlace: str | None


def exportar_eventos_2025_publicos(eventos: list[Evento2025]) -> list[dict[str, str | int]]:
    filas = []
    for evento in eventos:
        filas.append(
            {
                "evento_id": evento.evento_id,
                "fecha": evento.fecha.isoformat() if evento.fecha else "",
                "folio": evento.folio or "",
                "tipo": evento.tipo or "",
                "estatus": evento.estatus or "",
                "sexo": evento.sexo or "",
                "rango_edad": evento.rango_edad or "",
                "sustancia": evento.sustancia or "",
            }
        )
    return filas


def exportar_personas_2026_publicas(personas: list[Persona2026]) -> list[dict[str, str | int]]:
    filas = []
    for persona in personas:
        fila: dict[str, str | int] = {
            "folio": persona.folio,
            "iu_if": persona.iu_if or "",
            "edad": persona.edad or "",
            "genero": persona.genero or "",
            "sustancia": persona.sustancia_impacto or "",
            "total_sesiones": persona.total_sesiones,
            "estatus_tratamiento": persona.estatus_tratamiento or "",
            "motivo_egreso": persona.motivo_egreso or "",
        }
        for mes in range(1, 13):
            fila[f"sesiones_{mes:02d}"] = persona.sesiones_mes.get(mes, 0)
        filas.append(fila)
    return filas


def unificar_personas(eventos: list[Evento2025], personas: list[Persona2026]) -> list[PersonaUnificada]:
    """Vista interna."""
    por_folio: dict[str, list[Evento2025]] = defaultdict(list)
    for evento in eventos:
        if evento.folio:
            por_folio[evento.folio].append(evento)

    vistos = set()
    unificadas: list[PersonaUnificada] = []
    for persona in personas:
        evs = por_folio.get(persona.folio, [])
        origen = "ambos" if evs else "2026"
        ultimo = evs[-1] if evs else None
        sesiones_2025 = sum(1 for e in evs if e.es_sesion)
        unificadas.append(
            PersonaUnificada(
                folio=persona.folio,
                origen=origen,
                terapeuta_ultimo=persona.terapeuta or (ultimo.terapeuta if ultimo else None),
                sede=persona.sede or (ultimo.sede if ultimo else None),
                genero=persona.genero or (ultimo.sexo if ultimo else None),
                edad_o_rango=persona.edad or (ultimo.rango_edad if ultimo else None),
                sustancia=persona.sustancia_impacto or (ultimo.sustancia if ultimo else None),
                iu_if=persona.iu_if or extraer_iu_if(persona.folio),
                total_sesiones=persona.total_sesiones + sesiones_2025,
                desenlace=persona.motivo_egreso or persona.estatus_tratamiento,
            )
        )
        vistos.add(persona.folio)

    for folio, evs in por_folio.items():
        if folio in vistos:
            continue
        ultimo = evs[-1]
        unificadas.append(
            PersonaUnificada(
                folio=folio,
                origen="2025",
                terapeuta_ultimo=ultimo.terapeuta,
                sede=ultimo.sede,
                genero=ultimo.sexo,
                edad_o_rango=ultimo.rango_edad,
                sustancia=ultimo.sustancia,
                iu_if=extraer_iu_if(folio),
                total_sesiones=sum(1 for e in evs if e.es_sesion),
                desenlace=next(
                    (e.estatus for e in reversed(evs) if e.estatus and e.estatus != "sesión"),
                    None,
                ),
            )
        )
    return unificadas


def cruzar_por_folio(
    sesiones: list[SesionAgenda],
    personas: list[PersonaUnificada],
) -> list[tuple[SesionAgenda, PersonaUnificada | None]]:
    """Solo une si el ID de la sesión existe. No inventa."""
    indice = {p.folio: p for p in personas}
    cruce = []
    for sesion in sesiones:
        persona = indice.get(sesion.folio) if sesion.folio else None
        cruce.append((sesion, persona))
    return cruce


def escribir_publicos(
    ruta: Path,
    filas: list[dict[str, str | int]],
    columnas: list[str],
) -> None:
    escribir_csv(ruta, filas, columnas)


def guardar_sesiones(ruta: Path, sesiones: list[SesionAgenda], cerradas: bool = False) -> None:
    filas = exportar_sesiones_publicas(sesiones, cerradas=cerradas)
    columnas = list(filas[0].keys()) if filas else COLUMNAS_PUBLICAS
    escribir_csv(ruta, filas, columnas)
