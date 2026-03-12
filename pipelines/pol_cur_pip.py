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

from math import floor, ceil

import tkinter as tk
from tkinter import ttk, messagebox
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from matplotlib.ticker import LinearLocator, MaxNLocator, FormatStrFormatter, StrMethodFormatter
from math import floor, ceil

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


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

MARKER_OPTIONS = ["none", "^", "v", "o", "s", "d", "x", "+"]
LINESTYLE_OPTIONS = ["none", "-", "--", ":", "-."]

PC_PLOT_COLORS = {
    "asc_voltage": "#06a8c2",
    "dsc_voltage": "#2b3d8c",
    "asc_temperature": "#cf9a32",
    "dsc_temperature": "#ab3030",
}


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
def _format_limit_value(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:.6g}"


def _padded_limits(values: list[float], rel_pad: float = 0.05) -> tuple[float | None, float | None]:
    if not values:
        return None, None

    vmin = min(values)
    vmax = max(values)

    if vmin == vmax:
        pad = max(abs(vmin) * rel_pad, 1e-6)
    else:
        pad = (vmax - vmin) * rel_pad

    return vmin - pad, vmax + pad


def compute_default_v_vs_i_limits(bundle: CurveBundle) -> dict[str, str]:
    curve_data = build_curve_bundle_data(bundle)

    asc_rows = (
        select_fractional_point_per_step(curve_data["asc_rows"], curve_data["asc_tol"], 1.0)
        if curve_data["asc_rows"] else []
    )
    dsc_rows = (
        select_fractional_point_per_step(curve_data["dsc_rows"], curve_data["dsc_tol"], 1.0)
        if curve_data["dsc_rows"] else []
    )

    rows = asc_rows + dsc_rows
    if not rows:
        return {
            "x_min": "",
            "x_max": "",
            "v_min": "",
            "v_max": "",
        }

    x_min, x_max = _padded_limits([r["Corriente"] for r in rows])
    v_min, v_max = _padded_limits([r["Voltaje"] for r in rows])

    return {
        "x_min": _format_limit_value(x_min),
        "x_max": _format_limit_value(x_max),
        "v_min": _format_limit_value(v_min),
        "v_max": _format_limit_value(v_max),
    }

def draw_v_vs_i_on_figure(
    fig: Figure,
    bundle: CurveBundle,
    show_asc: bool,
    show_dsc: bool,
    show_voltage: bool,
    show_temperature: bool,
    point_fraction: float,
    asc_marker: str,
    dsc_marker: str,
    voltage_linestyle: str,
    temperature_linestyle: str,
    tick_count: int = 6,
    x_min: float | None = None,
    x_max: float | None = None,
    v_min: float | None = None,
    v_max: float | None = None,
    plot_title: str = "",
    title_fontsize: float = 14,
    tick_fontsize: float = 10,
    label_fontsize: float = 11,
    legend_fontsize: float = 10,
    marker_size: float = 6,
    hollow_markers: bool = False,
) -> bool:
    fig.clear()

    if not (show_asc or show_dsc):
        return False
    if not (show_voltage or show_temperature):
        return False

    tick_count = max(2, int(tick_count))

    curve_data = build_curve_bundle_data(bundle)

    asc_rows = (
        select_fractional_point_per_step(
            curve_data["asc_rows"],
            curve_data["asc_tol"],
            point_fraction,
        )
        if show_asc and curve_data["asc_rows"]
        else []
    )

    dsc_rows = (
        select_fractional_point_per_step(
            curve_data["dsc_rows"],
            curve_data["dsc_tol"],
            point_fraction,
        )
        if show_dsc and curve_data["dsc_rows"]
        else []
    )

    if not asc_rows and not dsc_rows:
        return False

    ax_main = fig.add_subplot(111)
    ax_temp = None

    if show_voltage and show_temperature:
        ax_temp = ax_main.twinx()

    asc_marker_mpl = _mpl_marker(asc_marker)
    dsc_marker_mpl = _mpl_marker(dsc_marker)
    voltage_ls_mpl = _mpl_linestyle(voltage_linestyle)
    temp_ls_mpl = _mpl_linestyle(temperature_linestyle)

    def _series_visible(marker_value: str, line_value: str) -> bool:
        return not (marker_value == "none" and line_value == "none")

    def _line_kwargs(color: str, mpl_marker: str, mpl_linestyle: str) -> dict:
        kwargs = {
            "color": color,
            "marker": mpl_marker,
            "linestyle": mpl_linestyle,
            "linewidth": 1.5,
            "markersize": marker_size,
        }
        

        if mpl_marker != "None":
            kwargs["markeredgecolor"] = color
            kwargs["markeredgewidth"] = 1.2

            # x and + are already line-only markers, so facecolor does not matter much
            if hollow_markers and mpl_marker not in {"x", "+"}:
                kwargs["markerfacecolor"] = "none"
            else:
                kwargs["markerfacecolor"] = "none"

        return kwargs

    # Voltage on main axis
    if show_voltage:
        if asc_rows and _series_visible(asc_marker, voltage_linestyle):
            ax_main.plot(
                [r["Corriente"] for r in asc_rows],
                [r["Voltaje"] for r in asc_rows],
                label="Asc V",
                **_line_kwargs(
                    PC_PLOT_COLORS["asc_voltage"],
                    asc_marker_mpl,
                    voltage_ls_mpl,
                ),
            )
        if dsc_rows and _series_visible(dsc_marker, voltage_linestyle):
            ax_main.plot(
                [r["Corriente"] for r in dsc_rows],
                [r["Voltaje"] for r in dsc_rows],
                label="Dsc V",
                **_line_kwargs(
                    PC_PLOT_COLORS["dsc_voltage"],
                    dsc_marker_mpl,
                    voltage_ls_mpl,
                ),
            )
        ax_main.set_ylabel("Voltaje (V)", fontsize=label_fontsize)

    # Temperature on second axis if needed
    if show_temperature:
        target_ax = ax_temp if ax_temp is not None else ax_main

        if asc_rows and _series_visible(asc_marker, temperature_linestyle):
            target_ax.plot(
                [r["Corriente"] for r in asc_rows],
                [r["Temperatura"] for r in asc_rows],
                label="Asc T",
                **_line_kwargs(
                    PC_PLOT_COLORS["asc_temperature"],
                    asc_marker_mpl,
                    temp_ls_mpl,
                ),
            )
        if dsc_rows and _series_visible(dsc_marker, temperature_linestyle):
            target_ax.plot(
                [r["Corriente"] for r in dsc_rows],
                [r["Temperatura"] for r in dsc_rows],
                label="Dsc T",
                **_line_kwargs(
                    PC_PLOT_COLORS["dsc_temperature"],
                    dsc_marker_mpl,
                    temp_ls_mpl,
                ),
            )
        target_ax.set_ylabel("Temperatura (°C)", fontsize=label_fontsize)

    handles, labels = ax_main.get_legend_handles_labels()
    if ax_temp is not None:
        h2, l2 = ax_temp.get_legend_handles_labels()
        handles += h2
        labels += l2

    if not handles:
        fig.clear()
        return False

    default_title = (
        f"V vs I - {bundle.description} #{bundle.curve_id} "
        f"(punto step = {point_fraction:.2f})"
    )
    final_title = plot_title.strip() if plot_title.strip() else default_title

    ax_main.set_xlabel("Corriente (A)", fontsize=label_fontsize)
    ax_main.set_title(final_title, fontsize=title_fontsize)
    ax_main.grid(True)

    # Tick font size
    ax_main.tick_params(axis="both", labelsize=tick_fontsize)
    ax_main.xaxis.set_major_locator(MaxNLocator(nbins=tick_count))

    if x_min is not None or x_max is not None:
        ax_main.set_xlim(left=x_min, right=x_max)

    # Voltage axis
    if show_voltage:
        if v_min is not None or v_max is not None:
            ax_main.set_ylim(bottom=v_min, top=v_max)
            ax_main.yaxis.set_major_locator(LinearLocator(tick_count))
        else:
            ax_main.yaxis.set_major_locator(MaxNLocator(nbins=tick_count))
        ax_main.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))

    # Temperature axis
    if ax_temp is not None:
        ax_temp.tick_params(axis="y", labelsize=tick_fontsize)

        temp_lines = ax_temp.get_lines()
        temp_values = []
        for line in temp_lines:
            temp_values.extend(line.get_ydata())

        apply_temperature_axis_scaling(
            ax_temp=ax_temp,
            temp_values=temp_values,
            tick_count=tick_count,
        )

    ax_main.legend(handles, labels, fontsize=legend_fontsize)
    fig.tight_layout()
    return True

def _mpl_marker(value: str) -> str:
    return "None" if value == "none" else value


def _mpl_linestyle(value: str) -> str:
    return "None" if value == "none" else value

def compute_autofit_v_vs_i_limits(
    bundle: CurveBundle,
    show_asc: bool,
    show_dsc: bool,
    show_voltage: bool,
    show_temperature: bool,
    point_fraction: float,
) -> dict[str, str]:
    if not (show_asc or show_dsc):
        raise ValueError("Debe seleccionar Asc y/o Dsc para usar Autofit.")

    if not (show_voltage or show_temperature):
        raise ValueError("Debe seleccionar Voltaje y/o Temperatura para usar Autofit.")

    curve_data = build_curve_bundle_data(bundle)

    asc_rows = (
        select_fractional_point_per_step(
            curve_data["asc_rows"],
            curve_data["asc_tol"],
            point_fraction,
        )
        if show_asc and curve_data["asc_rows"]
        else []
    )

    dsc_rows = (
        select_fractional_point_per_step(
            curve_data["dsc_rows"],
            curve_data["dsc_tol"],
            point_fraction,
        )
        if show_dsc and curve_data["dsc_rows"]
        else []
    )

    rows = asc_rows + dsc_rows
    if not rows:
        raise ValueError("No hay datos válidos para ajustar los ejes.")

    out = {
        "x_min": "0",
        "x_max": "",
        "v_min": "",
        "v_max": "",
    }

    currents = [r["Corriente"] for r in rows]
    if currents:
        out["x_max"] = str(int(ceil(max(currents))))

    if show_voltage:
        voltages = [r["Voltaje"] for r in rows]
        if voltages:
            out["v_min"] = str(int(floor(min(voltages))))
            out["v_max"] = str(int(ceil(max(voltages))))

    return out

def apply_temperature_axis_scaling(ax_temp, temp_values: list[float], tick_count: int) -> None:
    if not temp_values:
        return

    t_lo = floor(min(temp_values))
    t_hi = ceil(max(temp_values))

    if t_lo == t_hi:
        t_hi = t_lo + 1

    tick_count = max(2, int(tick_count))

    ax_temp.set_ylim(t_lo, t_hi)
    ax_temp.yaxis.set_major_locator(LinearLocator(tick_count))
    ax_temp.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))

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

def _optional_float(value: str | None) -> float | None:
    if value is None:
        return None
    text = value.strip()
    if not text:
        return None
    return float(text.replace(",", "."))


def build_curve_bundle_data(bundle: CurveBundle) -> dict[str, object]:
    asc_rows = concatenate_curve_data(bundle.asc_files)
    dsc_rows = concatenate_curve_data(bundle.dsc_files)

    asc_tol = infer_current_tolerance(bundle.asc_files) if bundle.asc_files else 1e-5
    dsc_tol = infer_current_tolerance(bundle.dsc_files) if bundle.dsc_files else 1e-5

    return {
        "asc_rows": asc_rows,
        "dsc_rows": dsc_rows,
        "asc_tol": asc_tol,
        "dsc_tol": dsc_tol,
    }


def split_rows_into_steps(
    rows: list[dict[str, float]],
    current_tolerance: float,
) -> list[list[dict[str, float]]]:
    """Split concatenated rows into current plateaus (steps)."""
    if not rows:
        return []

    steps: list[list[dict[str, float]]] = []
    start = 0
    plateau_current = rows[0]["Corriente"]

    for idx in range(1, len(rows)):
        current = rows[idx]["Corriente"]

        if abs(current - plateau_current) > current_tolerance:
            steps.append(rows[start:idx])
            start = idx
            plateau_current = rows[start]["Corriente"]
        else:
            span = idx - start + 1
            plateau_current = ((plateau_current * (span - 1)) + current) / span

    steps.append(rows[start:])
    return steps


def pick_fractional_point_from_step(
    step_rows: list[dict[str, float]],
    fraction: float,
) -> dict[str, float]:
    """Pick one point inside a step:
    0.0 = first point, 1.0 = last point, values in between = proportional index.
    """
    if not step_rows:
        raise ValueError("El step está vacío.")

    f = max(0.0, min(1.0, float(fraction)))

    if len(step_rows) == 1:
        return dict(step_rows[0])

    idx = int(round(f * (len(step_rows) - 1)))
    idx = max(0, min(len(step_rows) - 1, idx))
    return dict(step_rows[idx])


def select_fractional_point_per_step(
    rows: list[dict[str, float]],
    current_tolerance: float,
    fraction: float,
) -> list[dict[str, float]]:
    """Return one representative point per step according to a fractional position."""
    steps = split_rows_into_steps(rows, current_tolerance)
    selected_rows: list[dict[str, float]] = []

    for step_number, step_rows in enumerate(steps, start=1):
        selected = pick_fractional_point_from_step(step_rows, fraction)
        selected["Step"] = float(step_number)
        selected_rows.append(selected)

    return selected_rows


def find_last_point_of_each_step(
    rows: list[dict[str, float]],
    current_tolerance: float,
) -> list[dict[str, float]]:
    """Compatibility helper for export sheets: last point = fraction 1.0."""
    return select_fractional_point_per_step(rows, current_tolerance, 1.0)


def open_v_vs_i_window(input_dir: Path) -> None:
    bundles = discover_curve_bundles(Path(input_dir))
    if not bundles:
        raise ValueError("No se encontraron curvas de polarización válidas.")

    bundle = bundles[0]

    default_limits = compute_default_v_vs_i_limits(bundle)

    win = tk.Toplevel()
    win.title(f"PC - V vs I - {bundle.description} #{bundle.curve_id}")
    win.geometry("1200x700")

    controls_frame = ttk.Frame(win, padding=10)
    controls_frame.pack(side="left", fill="y")

    plot_outer = ttk.Frame(win, padding=10)
    plot_outer.pack(side="right", fill="both", expand=True)

    toolbar_frame = ttk.Frame(plot_outer)
    toolbar_frame.pack(side="top", fill="x")

    canvas_frame = ttk.Frame(plot_outer)
    canvas_frame.pack(side="top", fill="both", expand=True)

    fig = Figure(figsize=(9, 5.5), dpi=100)

    canvas = FigureCanvasTkAgg(fig, master=canvas_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)

    toolbar = NavigationToolbar2Tk(canvas, toolbar_frame, pack_toolbar=False)
    toolbar.update()
    toolbar.pack(side="left", fill="x")

    status_var = tk.StringVar(value="Listo.")

    asc_marker_var = tk.StringVar(value="^")
    dsc_marker_var = tk.StringVar(value="v")
    voltage_line_var = tk.StringVar(value="-")
    temperature_line_var = tk.StringVar(value="--")

    asc_var = tk.BooleanVar(value=True)
    dsc_var = tk.BooleanVar(value=True)
    voltage_var = tk.BooleanVar(value=True)
    temperature_var = tk.BooleanVar(value=False)
    point_fraction_var = tk.DoubleVar(value=1.0)

    x_min_var = tk.StringVar(value=default_limits["x_min"])
    x_max_var = tk.StringVar(value=default_limits["x_max"])
    v_min_var = tk.StringVar(value=default_limits["v_min"])
    v_max_var = tk.StringVar(value=default_limits["v_max"])

    tick_count_var = tk.IntVar(value=6)

    plot_title_var = tk.StringVar(value="")
    title_fontsize_var = tk.StringVar(value="14")
    tick_fontsize_var = tk.StringVar(value="10")
    label_fontsize_var = tk.StringVar(value="11")
    legend_fontsize_var = tk.StringVar(value="10")
    marker_size_var = tk.StringVar(value="6")
    hollow_markers_var = tk.BooleanVar(value=False)

    initial_state = {
        "asc": True,
        "dsc": True,
        "voltage": True,
        "temperature": False,
        "fraction": 1.0,
        "x_min": default_limits["x_min"],
        "x_max": default_limits["x_max"],
        "v_min": default_limits["v_min"],
        "v_max": default_limits["v_max"],
        "asc_marker": "^",
        "dsc_marker": "v",
        "voltage_line": "-",
        "temperature_line": "--",
        "tick_count": 6,
        "plot_title": "",
        "title_fontsize": "14",
        "tick_fontsize": "10",
        "label_fontsize": "11",
        "legend_fontsize": "10",
        "marker_size": "6",
        "hollow_markers": False,
    }

    def _schedule_plot(*_args):
        if suspend_events["value"]:
            return
        if plot_job["id"] is not None:
            win.after_cancel(plot_job["id"])
        plot_job["id"] = win.after(20, _plot)

    ttk.Label(
        controls_frame,
        text=f"Curva detectada:\n{bundle.description} #{bundle.curve_id}",
        justify="left",
    ).pack(anchor="w", pady=(0, 10))

    series_box = ttk.LabelFrame(controls_frame, text="Series")
    series_box.pack(fill="x", pady=5)

    style_box = ttk.LabelFrame(controls_frame, text="Estilo")
    style_box.pack(fill="x", pady=5)

    ttk.Label(style_box, text="Asc marker").grid(row=0, column=0, sticky="w", padx=8, pady=3)
    asc_marker_combo = ttk.Combobox(
        style_box,
        textvariable=asc_marker_var,
        values=MARKER_OPTIONS,
        state="readonly",
        width=10,
    )
    asc_marker_combo.grid(row=0, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(style_box, text="Dsc marker").grid(row=1, column=0, sticky="w", padx=8, pady=3)
    dsc_marker_combo = ttk.Combobox(
        style_box,
        textvariable=dsc_marker_var,
        values=MARKER_OPTIONS,
        state="readonly",
        width=10,
    )
    dsc_marker_combo.grid(row=1, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(style_box, text="Voltaje line").grid(row=2, column=0, sticky="w", padx=8, pady=3)
    voltage_line_combo = ttk.Combobox(
        style_box,
        textvariable=voltage_line_var,
        values=LINESTYLE_OPTIONS,
        state="readonly",
        width=10,
    )
    voltage_line_combo.grid(row=2, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(style_box, text="Temperatura line").grid(row=3, column=0, sticky="w", padx=8, pady=3)
    temperature_line_combo = ttk.Combobox(
        style_box,
        textvariable=temperature_line_var,
        values=LINESTYLE_OPTIONS,
        state="readonly",
        width=10,
    )

    ttk.Label(style_box, text="Ticks").grid(row=4, column=0, sticky="w", padx=8, pady=3)

    tick_spin = tk.Spinbox(
        style_box,
        from_=2,
        to=10,
        textvariable=tick_count_var,
        width=8,
    )
    tick_spin.grid(row=4, column=1, sticky="w", padx=8, pady=3)
    temperature_line_combo.grid(row=3, column=1, sticky="w", padx=8, pady=3)

    point_box = ttk.LabelFrame(controls_frame, text="Punto dentro de cada step")
    point_box.pack(fill="x", pady=5)

    limits_box = ttk.LabelFrame(controls_frame, text="Límites de ejes")
    limits_box.pack(fill="x", pady=5)

    plot_job = {"id": None}
    suspend_events = {"value": False}

    text_box = ttk.LabelFrame(controls_frame, text="Texto / tamaños")
    text_box.pack(fill="x", pady=5)

    ttk.Label(text_box, text="Título").grid(row=0, column=0, sticky="w", padx=8, pady=3)
    title_entry = ttk.Entry(text_box, textvariable=plot_title_var, width=28)
    title_entry.grid(row=0, column=1, sticky="we", padx=8, pady=3)

    ttk.Label(text_box, text="Title size").grid(row=1, column=0, sticky="w", padx=8, pady=3)
    title_size_entry = ttk.Entry(text_box, textvariable=title_fontsize_var, width=10)
    title_size_entry.grid(row=1, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Tick size").grid(row=2, column=0, sticky="w", padx=8, pady=3)
    tick_size_entry = ttk.Entry(text_box, textvariable=tick_fontsize_var, width=10)
    tick_size_entry.grid(row=2, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Label size").grid(row=3, column=0, sticky="w", padx=8, pady=3)
    label_size_entry = ttk.Entry(text_box, textvariable=label_fontsize_var, width=10)
    label_size_entry.grid(row=3, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Legend size").grid(row=4, column=0, sticky="w", padx=8, pady=3)
    legend_size_entry = ttk.Entry(text_box, textvariable=legend_fontsize_var, width=10)
    legend_size_entry.grid(row=4, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Marker size").grid(row=5, column=0, sticky="w", padx=8, pady=3)
    marker_size_entry = ttk.Entry(text_box, textvariable=marker_size_var, width=10)
    marker_size_entry.grid(row=5, column=1, sticky="w", padx=8, pady=3)

    ttk.Checkbutton(
        text_box,
        text="Hollow markers",
        variable=hollow_markers_var,
        command=_schedule_plot,
    ).grid(row=6, column=0, columnspan=2, sticky="w", padx=8, pady=4)

    def _positive_float(text: str, name: str) -> float:
        value = text.strip().replace(",", ".")
        if not value:
            raise ValueError(f"{name} no puede estar vacío.")
        num = float(value)
        if num <= 0:
            raise ValueError(f"{name} debe ser mayor que 0.")
        return num

    def _collect_limits():
        return dict(
            x_min=_optional_float(x_min_var.get()),
            x_max=_optional_float(x_max_var.get()),
            v_min=_optional_float(v_min_var.get()),
            v_max=_optional_float(v_max_var.get()),
        )
    
    def _plot():
        plot_job["id"] = None

        try:
            has_plot = draw_v_vs_i_on_figure(
                fig=fig,
                bundle=bundle,
                show_asc=asc_var.get(),
                show_dsc=dsc_var.get(),
                asc_marker=asc_marker_var.get(),
                dsc_marker=dsc_marker_var.get(),
                voltage_linestyle=voltage_line_var.get(),
                temperature_linestyle=temperature_line_var.get(),
                show_voltage=voltage_var.get(),
                show_temperature=temperature_var.get(),
                point_fraction=point_fraction_var.get(),
                tick_count=tick_count_var.get(),
                plot_title=plot_title_var.get(),
                title_fontsize=_positive_float(title_fontsize_var.get(), "Title size"),
                tick_fontsize=_positive_float(tick_fontsize_var.get(), "Tick size"),
                label_fontsize=_positive_float(label_fontsize_var.get(), "Label size"),
                legend_fontsize=_positive_float(legend_fontsize_var.get(), "Legend size"),
                marker_size=_positive_float(marker_size_var.get(), "Marker size"),
                hollow_markers=hollow_markers_var.get(),
                **_collect_limits(),
            )

        except ValueError as exc:
            status_var.set(f"Error: {exc}")
            return

        if not has_plot:
            fig.clear()
            canvas.draw_idle()
            status_var.set("No se muestra gráfico: seleccione al menos una dirección y una magnitud.")
            return

        canvas.draw_idle()
        status_var.set("Gráfico actualizado.")

    def _update_point_label(_event=None):
        point_value_label.config(text=f"{point_fraction_var.get():.2f}")

    def _on_scale_move(_value=None):
        _update_point_label()
        _plot()

    def _on_scale_release(_event=None):
        _schedule_plot()

    def _autofit():
        try:
            fitted = compute_autofit_v_vs_i_limits(
                bundle=bundle,
                show_asc=asc_var.get(),
                show_dsc=dsc_var.get(),
                show_voltage=voltage_var.get(),
                show_temperature=temperature_var.get(),
                point_fraction=point_fraction_var.get(),
            )
        except ValueError as exc:
            status_var.set(f"Error: {exc}")
            return

        suspend_events["value"] = True
        try:
            x_min_var.set(fitted["x_min"])
            x_max_var.set(fitted["x_max"])

            if fitted["v_min"] != "" or fitted["v_max"] != "":
                v_min_var.set(fitted["v_min"])
                v_max_var.set(fitted["v_max"])
        finally:
            suspend_events["value"] = False

        _plot()
        status_var.set("Autofit aplicado.")

    def _reset():
        suspend_events["value"] = True
        try:
            asc_var.set(initial_state["asc"])
            dsc_var.set(initial_state["dsc"])
            asc_marker_var.set(initial_state["asc_marker"])
            dsc_marker_var.set(initial_state["dsc_marker"])
            voltage_line_var.set(initial_state["voltage_line"])
            temperature_line_var.set(initial_state["temperature_line"])
            voltage_var.set(initial_state["voltage"])
            temperature_var.set(initial_state["temperature"])
            point_fraction_var.set(initial_state["fraction"])
            tick_count_var.set(initial_state["tick_count"])
            plot_title_var.set(initial_state["plot_title"])
            title_fontsize_var.set(initial_state["title_fontsize"])
            tick_fontsize_var.set(initial_state["tick_fontsize"])
            label_fontsize_var.set(initial_state["label_fontsize"])
            legend_fontsize_var.set(initial_state["legend_fontsize"])
            marker_size_var.set(initial_state["marker_size"])
            hollow_markers_var.set(initial_state["hollow_markers"])

            x_min_var.set(initial_state["x_min"])
            x_max_var.set(initial_state["x_max"])
            v_min_var.set(initial_state["v_min"])
            v_max_var.set(initial_state["v_max"])
            _update_point_label()
        finally:
            suspend_events["value"] = False

        _plot()
        status_var.set("Valores restaurados.")

    asc_marker_combo.bind("<<ComboboxSelected>>", _schedule_plot)
    dsc_marker_combo.bind("<<ComboboxSelected>>", _schedule_plot)
    voltage_line_combo.bind("<<ComboboxSelected>>", _schedule_plot)
    temperature_line_combo.bind("<<ComboboxSelected>>", _schedule_plot)
    tick_spin.bind("<Return>", _schedule_plot)
    tick_spin.bind("<FocusOut>", _schedule_plot)
    tick_spin.config(command=_schedule_plot)

    for widget in (
        title_entry,
        title_size_entry,
        tick_size_entry,
        label_size_entry,
        legend_size_entry,
        marker_size_entry,
    ):
        widget.bind("<Return>", _schedule_plot)
        widget.bind("<KP_Enter>", _schedule_plot)
        widget.bind("<FocusOut>", _schedule_plot)

    ttk.Checkbutton(series_box, text="Asc", variable=asc_var, command=_schedule_plot).pack(
        anchor="w", padx=8, pady=2
    )
    ttk.Checkbutton(series_box, text="Dsc", variable=dsc_var, command=_schedule_plot).pack(
        anchor="w", padx=8, pady=2
    )
    ttk.Checkbutton(series_box, text="Voltaje", variable=voltage_var, command=_schedule_plot).pack(
        anchor="w", padx=8, pady=2
    )
    ttk.Checkbutton(series_box, text="Temperatura", variable=temperature_var, command=_schedule_plot).pack(
        anchor="w", padx=8, pady=2
    )

    point_value_label = ttk.Label(point_box, text="1.00")
    point_value_label.pack(anchor="e", padx=8, pady=(4, 0))
    

    point_scale = ttk.Scale(
        point_box,
        from_=0.0,
        to=1.0,
        orient="horizontal",
        variable=point_fraction_var,
        command=_on_scale_move,
    )
    point_scale.pack(fill="x", padx=8, pady=6)

    ttk.Label(point_box, text="0 = primer punto, 1 = último punto").pack(
        anchor="w", padx=8, pady=(0, 6)
    )

    limit_specs = [
        ("I min", x_min_var),
        ("I max", x_max_var),
        ("V min", v_min_var),
        ("V max", v_max_var),
    ]

    entry_widgets = []
    for row_idx, (label, var) in enumerate(limit_specs):
        ttk.Label(limits_box, text=label).grid(row=row_idx, column=0, sticky="w", padx=8, pady=3)
        entry = ttk.Entry(limits_box, textvariable=var, width=12)
        entry.grid(row=row_idx, column=1, sticky="w", padx=8, pady=3)
        entry.bind("<KP_Enter>", _schedule_plot)
        entry.bind("<FocusOut>", _schedule_plot)
        entry.bind("<Return>", _schedule_plot)
        entry_widgets.append(entry)

    ttk.Label(
        controls_frame,
        textvariable=status_var,
        wraplength=260,
        justify="left",
    ).pack(anchor="w", fill="x", pady=(10, 10))

    buttons_frame = ttk.Frame(controls_frame)
    buttons_frame.pack(fill="x", pady=(5, 0))

    ttk.Button(buttons_frame, text="Reset", command=_reset).pack(side="left", padx=(0, 6))
    ttk.Button(buttons_frame, text="Autofit", command=_autofit).pack(side="left")

    _update_point_label()
    _plot()
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

    ws_dsc_last = wb.create_sheet("Dsc_last")
    _write_data_sheet(ws_dsc_last, dsc_last_rows, include_step=True)

    for ws in (ws_meta, ws_asc, ws_dsc, ws_asc_last, ws_dsc_last):
        _auto_format_sheet(ws)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


# ---------------------------------------------------------------------------
# Public entry point (GUI-compatible)
# ---------------------------------------------------------------------------
def export_folder(input_dir: Path, output_dir: Path) -> list[Path]:
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

def _show_pc_stub(title: str) -> None:
    messagebox.showinfo("PC", f"{title} aún no está implementado.")


def run_pipeline(
    input_dir: Path,
    output_dir: Path,
    selected_options: list[str] | None = None,
) -> list[Path]:
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)

    exported_files = export_folder(input_dir, output_dir)
    if not exported_files:
        return []

    chosen = set(selected_options or [])

    if "V vs I" in chosen:
        open_v_vs_i_window(input_dir)

    if "Series by time" in chosen:
        _show_pc_stub("Series by time")

    if "dV/dI" in chosen:
        _show_pc_stub("dV/dI")

    if "Step Stability" in chosen:
        _show_pc_stub("Step Stability")

    return exported_files



def open_series_by_time_window(input_dir: Path) -> None:
    import tkinter.messagebox as mb
    mb.showinfo("PC", "Series by time aún no está implementado.")


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
