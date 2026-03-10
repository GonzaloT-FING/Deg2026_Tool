"""Polarization curve (.DTA) -> Excel (.xlsx) exporter for Gamry chronopotentiometry.

What this first version does:
  - Finds all .DTA files whose filename starts with 'Curva_Polarizacion_'
  - Separates Asc and Dsc files
  - Sorts each curve by the LAST number after '#'
  - Reconstructs the full ascending and descending polarization curves
  - Exports ONE .xlsx per curve with three sheets:
        1) Metadata  -> Campo / Valor / Unidad
        2) Asc       -> headers row, units row, then numeric data
        3) Dsc       -> headers row, units row, then numeric data

This version focuses on importing, gathering, parsing, concatenating, and exporting.
No further processing/plotting is implemented yet.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from collections import defaultdict
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series


# ---------------------------------------------------------------------------
# Export labels
# ---------------------------------------------------------------------------
META_ROWS_ORDER = [
    "Técnica",
    "Fecha",
    "Hora",
    "Duración del paso",
    "Rango I",
    "Paso I",
    "Tiempo de muestreo",
    "Área",
]

DATA_EXPORT = [
    ("Pt", "Pt", ""),
    ("T", "time", "s"),
    ("Vf", "Voltaje", "V"),
    ("Im", "Corriente", "A"),
    ("Sig", "Sig", "V"),
    ("Ach", "Ach", "V"),
    ("Temp", "Temperatura", "ºC"),
]

FILE_RE = re.compile(
    r"^Curva_Polarizacion_(?P<direction>Asc|Dsc)_(?P<description>.+?)_#(?P<curve_id>\d+)_#(?P<file_index>\d+)\.DTA$",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Data containers
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class PolarizationFile:
    path: Path
    direction: str
    description: str
    curve_id: int
    file_index: int


@dataclass
class ParsedDTA:
    meta_values: dict[str, str]
    meta_units: dict[str, str]
    header: list[str]
    units: list[str]
    rows: list[list[str]]


@dataclass
class CurveBundle:
    description: str
    curve_id: int
    asc_files: list[PolarizationFile]
    dsc_files: list[PolarizationFile]


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def to_float(val: str) -> float | None:
    """Convert Gamry-style numbers (decimal comma) to float."""
    s = val.strip()
    if not s:
        return None
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _drop_leading_blank(parts: list[str]) -> list[str]:
    if parts and parts[0] == "":
        return parts[1:]
    return parts


def _extract_parenthesized_unit(text: str) -> str:
    matches = re.findall(r"\(([^()]*)\)", text)
    if matches:
        return matches[-1].strip()
    return ""


def _extract_meta_unit(key: str, description: str) -> str:
    unit = _extract_parenthesized_unit(description)
    if unit:
        return unit

    fallback_units = {
        "TITLE": "",
        "DATE": "",
        "TIME": "",
        "IPRESTEP": "A",
        "ISTEP1": "A",
        "ISTEP2": "A",
        "TPRESTEP": "s",
        "TSTEP1": "s",
        "TSTEP2": "s",
        "SAMPLETIME": "s",
        "AREA": "cm^2",
    }
    return fallback_units.get(key, "")


def _fmt_range_value(i_min: float | None, i_max: float | None) -> str:
    if i_min is None or i_max is None:
        return ""
    return f"{i_min:g} a {i_max:g}"


def _column_index(parsed: ParsedDTA, column_name: str) -> int | None:
    try:
        return parsed.header.index(column_name)
    except ValueError:
        return None


def _required_float(row: list[str], idx: int | None) -> float:
    if idx is None or idx >= len(row):
        raise ValueError("Falta una columna requerida en la tabla CURVE.")
    num = to_float(row[idx])
    if num is None:
        raise ValueError(f"No se pudo convertir a número: {row[idx]!r}")
    return num


def _step_delta_from_file(item: PolarizationFile) -> float | None:
    parsed = parse_gamry_dta(item.path)
    i1 = to_float(parsed.meta_values.get("ISTEP1", ""))
    i2 = to_float(parsed.meta_values.get("ISTEP2", ""))
    if i1 is None or i2 is None:
        return None
    return abs(i2 - i1)


def infer_current_tolerance(files: list[PolarizationFile]) -> float:
    """Infer a tolerance to identify current plateaus from measured current."""
    for item in files:
        step_delta = _step_delta_from_file(item)
        if step_delta is not None and step_delta > 0:
            return max(min(step_delta * 0.1, 1e-3), 1e-5)
    return 1e-5


# ---------------------------------------------------------------------------
# File discovery and grouping
# ---------------------------------------------------------------------------
def _parse_filename(path: Path) -> PolarizationFile | None:
    match = FILE_RE.match(path.name)
    if not match:
        return None

    return PolarizationFile(
        path=path,
        direction=match.group("direction").title(),
        description=match.group("description"),
        curve_id=int(match.group("curve_id")),
        file_index=int(match.group("file_index")),
    )


def discover_curve_bundles(input_dir: Path) -> list[CurveBundle]:
    grouped: dict[tuple[str, int], dict[str, list[PolarizationFile]]] = defaultdict(
        lambda: {"Asc": [], "Dsc": []}
    )

    for path in sorted(input_dir.glob("*.DTA")):
        info = _parse_filename(path)
        if info is None:
            continue
        grouped[(info.description, info.curve_id)][info.direction].append(info)

    bundles: list[CurveBundle] = []
    for (description, curve_id), by_dir in sorted(grouped.items(), key=lambda x: (x[0][0], x[0][1])):
        asc_files = sorted(by_dir["Asc"], key=lambda item: item.file_index)
        dsc_files = sorted(by_dir["Dsc"], key=lambda item: item.file_index)
        bundles.append(
            CurveBundle(
                description=description,
                curve_id=curve_id,
                asc_files=asc_files,
                dsc_files=dsc_files,
            )
        )
    return bundles


# ---------------------------------------------------------------------------
# Parsing one Gamry CHRONOP file
# ---------------------------------------------------------------------------
def parse_gamry_dta(path: Path) -> ParsedDTA:
    text = path.read_text(encoding="latin-1", errors="replace")
    lines = text.splitlines()

    meta_values: dict[str, str] = {}
    meta_units: dict[str, str] = {}
    header: list[str] | None = None
    units: list[str] = []
    rows: list[list[str]] = []

    table_started = False

    for line in lines:
        if not table_started:
            if line.startswith("CURVE") and "TABLE" in line:
                table_started = True
                continue

            if not line.strip():
                continue

            parts = line.split("\t")
            if len(parts) >= 3 and parts[0].strip():
                key = parts[0].strip()
                value = parts[2].strip()
                description = " ".join(p.strip() for p in parts[3:] if p.strip())
                meta_values[key] = value
                meta_units[key] = _extract_meta_unit(key, description)
            continue

        if not line.strip():
            continue

        parts = _drop_leading_blank([p.strip() for p in line.rstrip("\r\n").split("\t")])
        if not parts:
            continue

        first = parts[0]
        if header is None:
            if first == "Pt":
                header = parts
            continue

        if not units:
            if first == "#":
                units = parts
            continue

        if re.fullmatch(r"-?\d+", first):
            rows.append(parts)

    if header is None:
        raise ValueError(f"No data header found in {path.name} (expected 'Pt ...')")

    return ParsedDTA(
        meta_values=meta_values,
        meta_units=meta_units,
        header=header,
        units=units,
        rows=rows,
    )


# ---------------------------------------------------------------------------
# Build metadata
# ---------------------------------------------------------------------------
def _collect_current_extremes(files: list[PolarizationFile]) -> tuple[float | None, float | None]:
    currents: list[float] = []
    for item in files:
        parsed = parse_gamry_dta(item.path)
        for key in ("ISTEP1", "ISTEP2"):
            val = to_float(parsed.meta_values.get(key, ""))
            if val is not None:
                currents.append(val)
    if not currents:
        return None, None
    return min(currents), max(currents)


def build_metadata(bundle: CurveBundle) -> list[tuple[str, object, str]]:
    reference_files = bundle.asc_files or bundle.dsc_files
    if not reference_files:
        raise ValueError("No se encontraron archivos Asc ni Dsc para exportar.")

    first_parsed = parse_gamry_dta(reference_files[0].path)

    step_duration = to_float(first_parsed.meta_values.get("TSTEP1", ""))
    sample_time = to_float(first_parsed.meta_values.get("SAMPLETIME", ""))
    area = to_float(first_parsed.meta_values.get("AREA", ""))

    i1 = to_float(first_parsed.meta_values.get("ISTEP1", ""))
    i2 = to_float(first_parsed.meta_values.get("ISTEP2", ""))
    delta_i = abs(i2 - i1) if i1 is not None and i2 is not None else None

    i_min, i_max = _collect_current_extremes(reference_files)

    metadata_map: dict[str, tuple[object, str]] = {
        "Técnica": (first_parsed.meta_values.get("TITLE", ""), ""),
        "Fecha": (first_parsed.meta_values.get("DATE", ""), ""),
        "Hora": (first_parsed.meta_values.get("TIME", ""), ""),
        "Duración del paso": (step_duration if step_duration is not None else "", "s"),
        "Rango I": (_fmt_range_value(i_min, i_max), "A"),
        "Paso I": (delta_i if delta_i is not None else "", "A"),
        "Tiempo de muestreo": (sample_time if sample_time is not None else "", "s"),
        "Área": (area if area is not None else "", "cm^2"),
    }

    return [(field, *metadata_map[field]) for field in META_ROWS_ORDER]


# ---------------------------------------------------------------------------
# Concatenate Asc / Dsc data
# ---------------------------------------------------------------------------
def _extract_local_rows(parsed: ParsedDTA) -> list[dict[str, float]]:
    idx_map = {source: _column_index(parsed, source) for source, _, _ in DATA_EXPORT}

    missing = [source for source, idx in idx_map.items() if idx is None]
    if missing:
        raise ValueError(
            "Faltan columnas requeridas en la tabla CURVE: " + ", ".join(missing)
        )

    out: list[dict[str, float]] = []
    for raw_row in parsed.rows:
        record: dict[str, float] = {}
        for source, export_name, _unit in DATA_EXPORT:
            record[export_name] = _required_float(raw_row, idx_map[source])
        out.append(record)
    return out


def concatenate_curve_data(files: list[PolarizationFile]) -> list[dict[str, float]]:
    all_rows: list[dict[str, float]] = []
    time_offset = 0.0
    global_pt = 0

    for item in files:
        parsed = parse_gamry_dta(item.path)
        local_rows = _extract_local_rows(parsed)
        if not local_rows:
            continue

        first_local_time = local_rows[0]["time"]

        for local in local_rows:
            row = {
                "Pt": float(global_pt),
                "time": local["time"] - first_local_time + time_offset,
                "Voltaje": local["Voltaje"],
                "Corriente": local["Corriente"],
                "Sig": local["Sig"],
                "Ach": local["Ach"],
                "Temperatura": local["Temperatura"],
            }
            all_rows.append(row)
            global_pt += 1

        sample_time = to_float(parsed.meta_values.get("SAMPLETIME", ""))
        if sample_time is None:
            # Fallback: infer it from the first two time values if possible.
            if len(local_rows) >= 2:
                sample_time = local_rows[1]["time"] - local_rows[0]["time"]
            else:
                sample_time = 0.0

        time_offset = all_rows[-1]["time"] + sample_time

    return all_rows


# ---------------------------------------------------------------------------
# Stable-point helper
# ---------------------------------------------------------------------------
def find_last_point_of_each_step(
    rows: list[dict[str, float]],
    current_tolerance: float,
) -> list[dict[str, float]]:
    """Return the last point of each current plateau."""
    if not rows:
        return []

    stable_rows: list[dict[str, float]] = []
    plateau_start = 0
    plateau_current = rows[0]["Corriente"]
    step_number = 1

    for idx in range(1, len(rows)):
        current = rows[idx]["Corriente"]
        if abs(current - plateau_current) > current_tolerance:
            last_row = dict(rows[idx - 1])
            last_row["Step"] = float(step_number)
            stable_rows.append(last_row)

            plateau_start = idx
            plateau_current = rows[plateau_start]["Corriente"]
            step_number += 1
        else:
            span = idx - plateau_start + 1
            plateau_current = ((plateau_current * (span - 1)) + current) / span

    last_row = dict(rows[-1])
    last_row["Step"] = float(step_number)
    stable_rows.append(last_row)
    return stable_rows


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------
def _write_metadata_sheet(ws, metadata_rows: list[tuple[str, object, str]]) -> None:
    ws.title = "Metadata"
    ws["A1"] = "Campo"
    ws["B1"] = "Valor"
    ws["C1"] = "Unidad"
    for ref in ("A1", "B1", "C1"):
        ws[ref].font = Font(bold=True)

    for row_idx, (field, value, unit) in enumerate(metadata_rows, start=2):
        ws.cell(row=row_idx, column=1, value=field)
        ws.cell(row=row_idx, column=2, value=value)
        ws.cell(row=row_idx, column=3, value=unit)

    ws.freeze_panes = "A2"


def _write_data_sheet(ws, rows: list[dict[str, float]], include_step: bool = False) -> None:
    headers = [label for _src, label, _unit in DATA_EXPORT]
    units = [unit for _src, _label, unit in DATA_EXPORT]

    if include_step:
        headers = ["Step"] + headers
        units = [""] + units

    for col_num, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = Font(bold=True)

    for col_num, unit in enumerate(units, start=1):
        ws.cell(row=2, column=col_num, value=unit)

    for row_num, record in enumerate(rows, start=3):
        for col_num, header in enumerate(headers, start=1):
            value = record.get(header, "")
            if header in {"Pt", "Step"} and value != "":
                value = int(value)
            ws.cell(row=row_num, column=col_num, value=value)

    ws.freeze_panes = "A3"


def _add_last_point_chart(ws, title: str, include_step: bool) -> None:
    if ws.max_row < 4:
        return

    current_col = 5 if include_step else 4
    voltage_col = 4 if include_step else 3

    chart = ScatterChart()
    chart.title = title
    chart.style = 2
    chart.x_axis.title = "Corriente (A)"
    chart.y_axis.title = "Voltaje (V)"
    chart.height = 8
    chart.width = 14

    xvalues = Reference(ws, min_col=current_col, min_row=3, max_row=ws.max_row)
    yvalues = Reference(ws, min_col=voltage_col, min_row=3, max_row=ws.max_row)
    series = Series(yvalues, xvalues, title="Puntos estabilizados")
    chart.series.append(series)

    ws.add_chart(chart, "J2")


def _auto_format_sheet(ws) -> None:
    for col_num in range(1, ws.max_column + 1):
        max_len = 0
        for row_num in range(1, min(ws.max_row, 100) + 1):
            value = ws.cell(row=row_num, column=col_num).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_num)].width = min(max(10, max_len + 2), 45)

    for row_cells in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row_cells:
            cell.alignment = Alignment(vertical="top")


def export_curve_bundle(bundle: CurveBundle, out_path: Path) -> None:
    metadata_rows = build_metadata(bundle)
    asc_rows = concatenate_curve_data(bundle.asc_files)
    dsc_rows = concatenate_curve_data(bundle.dsc_files)

    asc_last_rows = find_last_point_of_each_step(
        asc_rows,
        infer_current_tolerance(bundle.asc_files),
    )
    dsc_last_rows = find_last_point_of_each_step(
        dsc_rows,
        infer_current_tolerance(bundle.dsc_files),
    )

    wb = Workbook()
    wb.remove(wb.active)

    ws_meta = wb.create_sheet("Metadata")
    _write_metadata_sheet(ws_meta, metadata_rows)

    ws_asc = wb.create_sheet("Asc")
    _write_data_sheet(ws_asc, asc_rows)

    ws_dsc = wb.create_sheet("Dsc")
    _write_data_sheet(ws_dsc, dsc_rows)

    ws_asc_last = wb.create_sheet("Asc_last")
    _write_data_sheet(ws_asc_last, asc_last_rows, include_step=True)
    _add_last_point_chart(ws_asc_last, "Asc - puntos estabilizados", include_step=True)

    ws_dsc_last = wb.create_sheet("Dsc_last")
    _write_data_sheet(ws_dsc_last, dsc_last_rows, include_step=True)
    _add_last_point_chart(ws_dsc_last, "Dsc - puntos estabilizados", include_step=True)

    for ws in (ws_meta, ws_asc, ws_dsc, ws_asc_last, ws_dsc_last):
        _auto_format_sheet(ws)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


# ---------------------------------------------------------------------------
# Public entry point (GUI-compatible)
# ---------------------------------------------------------------------------
def export_folder(
    input_dir: Path,
    output_dir: Path,
    selected_options: list[str] | None = None,
) -> list[Path]:
    """Export one .xlsx per detected polarization-curve bundle.

    The selected_options argument is intentionally accepted so gui.py can call
    this function in the same style it already uses for EIS.
    This first version does not use the options yet.
    """
    del selected_options  # intentionally unused for now

    input_dir = Path(input_dir)
    output_dir = Path(output_dir)

    bundles = discover_curve_bundles(input_dir)
    exported_files: list[Path] = []

    for bundle in bundles:
        out_name = f"Curva_Polarizacion_{bundle.description}_#{bundle.curve_id}.xlsx"
        out_path = output_dir / out_name
        export_curve_bundle(bundle, out_path)
        exported_files.append(out_path)

    return exported_files


if __name__ == "__main__":
    repo_dir = Path(__file__).resolve().parents[1]
    input_dir = repo_dir / "data"
    output_dir = repo_dir / "outputs"

    exported = export_folder(input_dir, output_dir)
    if exported:
        print("Archivos exportados:")
        for path in exported:
            print(f" - {path}")
    else:
        print("No se encontraron archivos de polarización para exportar.")
