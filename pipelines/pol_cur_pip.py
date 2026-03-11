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

import tkinter as tk
from tkinter import ttk, messagebox
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

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
def build_v_vs_i_figure(
    bundle: CurveBundle,
    show_asc: bool,
    show_dsc: bool,
    show_voltage: bool,
    show_temperature: bool,
    point_fraction: float,
    x_min: float | None = None,
    x_max: float | None = None,
    v_min: float | None = None,
    v_max: float | None = None,
    t_min: float | None = None,
    t_max: float | None = None,
) -> Figure | None:
    if not (show_asc or show_dsc):
        return None
    if not (show_voltage or show_temperature):
        return None

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
        return None

    fig = Figure(figsize=(9, 5.5), dpi=100)
    ax_main = fig.add_subplot(111)
    ax_temp = None

    # If both are requested, use dual axis.
    # If only one is requested, use the main axis only.
    if show_voltage and show_temperature:
        ax_temp = ax_main.twinx()

    # Voltage
    if show_voltage:
        if asc_rows:
            ax_main.plot(
                [r["Corriente"] for r in asc_rows],
                [r["Voltaje"] for r in asc_rows],
                marker="o",
                label="Asc V",
            )
        if dsc_rows:
            ax_main.plot(
                [r["Corriente"] for r in dsc_rows],
                [r["Voltaje"] for r in dsc_rows],
                marker="s",
                label="Dsc V",
            )
        ax_main.set_ylabel("Voltaje (V)")

    # Temperature
    if show_temperature:
        target_ax = ax_temp if ax_temp is not None else ax_main

        if asc_rows:
            target_ax.plot(
                [r["Corriente"] for r in asc_rows],
                [r["Temperatura"] for r in asc_rows],
                marker="o",
                linestyle="--",
                label="Asc T",
            )
        if dsc_rows:
            target_ax.plot(
                [r["Corriente"] for r in dsc_rows],
                [r["Temperatura"] for r in dsc_rows],
                marker="s",
                linestyle="--",
                label="Dsc T",
            )
        target_ax.set_ylabel("Temperatura (°C)")

    ax_main.set_xlabel("Corriente (A)")
    ax_main.set_title(
        f"V vs I - {bundle.description} #{bundle.curve_id} "
        f"(punto step = {point_fraction:.2f})"
    )
    ax_main.grid(True)

    if x_min is not None or x_max is not None:
        ax_main.set_xlim(left=x_min, right=x_max)

    if show_voltage and (v_min is not None or v_max is not None):
        ax_main.set_ylim(bottom=v_min, top=v_max)

    if show_temperature:
        target_ax = ax_temp if ax_temp is not None else ax_main
        if t_min is not None or t_max is not None:
            target_ax.set_ylim(bottom=t_min, top=t_max)

    handles, labels = ax_main.get_legend_handles_labels()
    if ax_temp is not None:
        h2, l2 = ax_temp.get_legend_handles_labels()
        handles += h2
        labels += l2
    if handles:
        ax_main.legend(handles, labels)

    fig.tight_layout()
    return fig

def open_v_vs_i_window(input_dir: Path) -> None:
    bundles = discover_curve_bundles(Path(input_dir))
    if not bundles:
        raise ValueError("No se encontraron curvas de polarización válidas.")

    # According to your workflow, there should be only one valid curve per folder.
    bundle = bundles[0]

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

    status_var = tk.StringVar(value="Configure y presione Plot.")

    asc_var = tk.BooleanVar(value=True)
    dsc_var = tk.BooleanVar(value=True)
    voltage_var = tk.BooleanVar(value=True)
    temperature_var = tk.BooleanVar(value=False)

    point_fraction_var = tk.DoubleVar(value=1.0)

    x_min_var = tk.StringVar(value="")
    x_max_var = tk.StringVar(value="")
    v_min_var = tk.StringVar(value="")
    v_max_var = tk.StringVar(value="")
    t_min_var = tk.StringVar(value="")
    t_max_var = tk.StringVar(value="")

    ttk.Label(
        controls_frame,
        text=f"Curva detectada:\n{bundle.description} #{bundle.curve_id}",
        justify="left",
    ).pack(anchor="w", pady=(0, 10))

    series_box = ttk.LabelFrame(controls_frame, text="Series")
    series_box.pack(fill="x", pady=5)

    ttk.Checkbutton(series_box, text="Asc", variable=asc_var).pack(anchor="w", padx=8, pady=2)
    ttk.Checkbutton(series_box, text="Dsc", variable=dsc_var).pack(anchor="w", padx=8, pady=2)
    ttk.Checkbutton(series_box, text="Voltaje", variable=voltage_var).pack(anchor="w", padx=8, pady=2)
    ttk.Checkbutton(series_box, text="Temperatura", variable=temperature_var).pack(anchor="w", padx=8, pady=2)

    point_box = ttk.LabelFrame(controls_frame, text="Punto dentro de cada step")
    point_box.pack(fill="x", pady=5)

    point_value_label = ttk.Label(point_box, text="1.00")
    point_value_label.pack(anchor="e", padx=8, pady=(4, 0))

    def _update_point_label(_value=None):
        point_value_label.config(text=f"{point_fraction_var.get():.2f}")

    point_scale = ttk.Scale(
        point_box,
        from_=0.0,
        to=1.0,
        orient="horizontal",
        variable=point_fraction_var,
        command=_update_point_label,
    )
    point_scale.pack(fill="x", padx=8, pady=6)

    ttk.Label(
        point_box,
        text="0 = primer punto, 1 = último punto",
    ).pack(anchor="w", padx=8, pady=(0, 6))

    limits_box = ttk.LabelFrame(controls_frame, text="Límites de ejes")
    limits_box.pack(fill="x", pady=5)

    limit_rows = [
        ("I min", x_min_var),
        ("I max", x_max_var),
        ("V min", v_min_var),
        ("V max", v_max_var),
        ("T min", t_min_var),
        ("T max", t_max_var),
    ]

    for row_idx, (label, var) in enumerate(limit_rows):
        ttk.Label(limits_box, text=label).grid(row=row_idx, column=0, sticky="w", padx=8, pady=3)
        ttk.Entry(limits_box, textvariable=var, width=12).grid(
            row=row_idx, column=1, sticky="w", padx=8, pady=3
        )

    ttk.Label(
        controls_frame,
        textvariable=status_var,
        wraplength=260,
        justify="left",
    ).pack(anchor="w", fill="x", pady=(10, 10))

    canvas_ref = {"canvas": None, "toolbar": None}

    def _clear_canvas():
        if canvas_ref["toolbar"] is not None:
            canvas_ref["toolbar"].destroy()
            canvas_ref["toolbar"] = None

        if canvas_ref["canvas"] is not None:
            canvas_ref["canvas"].get_tk_widget().destroy()
            canvas_ref["canvas"] = None

    def _plot():
        try:
            fig = build_v_vs_i_figure(
                bundle=bundle,
                show_asc=asc_var.get(),
                show_dsc=dsc_var.get(),
                show_voltage=voltage_var.get(),
                show_temperature=temperature_var.get(),
                point_fraction=point_fraction_var.get(),
                x_min=_optional_float(x_min_var.get()),
                x_max=_optional_float(x_max_var.get()),
                v_min=_optional_float(v_min_var.get()),
                v_max=_optional_float(v_max_var.get()),
                t_min=_optional_float(t_min_var.get()),
                t_max=_optional_float(t_max_var.get()),
            )
        except ValueError as exc:
            status_var.set(f"Error: {exc}")
            return

        if fig is None:
            _clear_canvas()
            status_var.set("No se muestra gráfico: seleccione al menos una dirección y una magnitud.")
            return

        _clear_canvas()

        canvas = FigureCanvasTkAgg(fig, master=canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(canvas, toolbar_frame, pack_toolbar=False)
        toolbar.update()
        toolbar.pack(side="left", fill="x")

        canvas_ref["canvas"] = canvas
        canvas_ref["toolbar"] = toolbar
        status_var.set("Gráfico actualizado.")

    buttons_frame = ttk.Frame(controls_frame)
    buttons_frame.pack(fill="x", pady=(5, 0))

    ttk.Button(buttons_frame, text="Plot", command=_plot).pack(side="left", padx=(0, 6))
    ttk.Button(buttons_frame, text="Cerrar", command=win.destroy).pack(side="left")

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
