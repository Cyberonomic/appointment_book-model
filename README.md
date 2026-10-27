# Session

Appointment-grid extractor. Turns colour-coded weekly appointment books (Excel) into a clean table of sessions that actually occurred. A row is a session slot. The only join key to a case file is an explicit identifier; if it is missing, the session is counted but not linked. The extractor does not guess layout: day blocks are declared. Fill colour plus non-empty text marks an effective point; other type-work text is excluded.

Developed and tested on synthetic data only. 

## Module

1. `sesiones.reglas` — site normalization, excluded phrases, identifier parsing (`--folio-pattern`, `--excluded-phrases`).
2. `sesiones.agenda` — vertical and horizontal grid extractors.
3. `sesiones.expedientes` — 2025 and 2026 rows (`--sheet` for worksheet name).
4. `sesiones.exportar` — public CSV helpers and join-by-identifier.

Month block coordinates live in `sesiones/bloques.py` as one **example** layout family; declare your own blocks for other grids.

## Appointment grid.

**Attended session rule:** non-empty cell text and a colour fill (default suffix `6AA84F`). Internal-work phrases are skipped (built-in: meeting, break, total, lunch, workshop; extend via `--excluded-phrases`). Two names in one cell still count as one session.

**Vertical** months stack days; **horizontal** months place days side by side. Sites come from a local catalog file (`--sites`). A colour-effective slot without a header is skipped.

Public CSV columns: `contacto_id`, `fecha`, `terapeuta_alias`, `sede_alias`, `hora`, `folio`, `iu_if`, `layout` (aliases only; no cell text).

```bash
python -m sesiones agenda \
  --workbook ./datos/cartelera_ejemplo.xlsx \
  --sites sedes_local.txt \
  --month oct2025 \
  --output ./salida/sesiones_publicas.csv
```

Optional: `--folio-pattern`, `--excluded-phrases textos_excluidos_local.txt`, `--closed`, `--keep-internal`.

Month keys: `oct2025`, `nov2025`, `dic2025`, `ene2026`, `feb2026`, `mar2026`, `abr2026`, `may2026`, and others defined in `bloques.py`.

## Case files (expedientes)

Column indices match fixed 2025 / 2026 export layouts; adjust indices in code if your file differs.

```bash
python -m sesiones eventos-2025 \
  --workbook ./datos/expediente_2025_ejemplo.xlsx \
  --sheet Sheet1 \
  --output ./salida/eventos_2025.csv

python -m sesiones personas-2026 \
  --workbook ./datos/expediente_2026_ejemplo.xlsx \
  --sheet Data \
  --output ./salida/personas_2026.csv
```

Public outputs omit names and phone numbers. Join to sessions only when the session row carries the same identifier (`cruzar_por_folio`).

## Out of scope

Does not treat repeated names on the grid as one person. Does not merge 2025 and 2026 tables. Does not infer layout or coordinator totals.

## Tests

```bash
python -m unittest
```

All tests use in-memory synthetic workbooks.
