"""Degradation (.DTA) -> Excel (.xlsx) exporter and V vs t viewer."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from math import ceil, floor, log10
import re

import tkinter as tk
from tkinter import ttk, messagebox
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.ticker import LinearLocator, LogFormatterSciNotation, LogLocator, MaxNLocator, StrMethodFormatter

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


META_FIELDS = [
    ("TITLE", "Tecnica"),
    ("DATE", "Fecha"),
    ("TIME", "Hora"),
    ("ISTEP1", "Corriente"),
    ("TSTEP1", "Duracion"),
    ("SAMPLETIME", "Tiempo de muestreo"),
    ("AREA", "Area"),
]

DATA_EXPORT = [
    ("Pt", "Pt"),
    ("T", "time"),
    ("Vf", "Voltaje"),
    ("Im", "Corriente"),
    ("Temp", "Temperatura"),
]

DEG_FILE_RE = re.compile(
    r"^Degradacion_galvanostatica_19A_60C_#(?P<stage>\d+)\.DTA$",
    re.IGNORECASE,
)

LINESTYLE_OPTIONS = ["none", "-", "--", ":", "-."]
SMOOTHING_ALGORITHMS = ["Median filter", "Rolling average"]

DEG_STAGE_COLORS = [
    "#06a8c2",
    "#2b3d8c",
    "#cf9a32",
    "#ab3030",
    "#2f9e44",
    "#7c4dff",
    "#8d6e63",
    "#00897b",
]

DEG_PLOT_COLORS = {
    "temperature": "#cf9a32",
}

SECONDS_PER_MINUTE = 60.0
SECONDS_PER_HOUR = 3600.0
TIME_UNIT_OPTIONS = ["s", "min", "h"]


@dataclass(frozen=True)
class DegFile:
    path: Path
    stage: int


@dataclass
class ParsedDTA:
    meta_values: dict[str, str]
    meta_units: dict[str, str]
    header: list[str]
    units: list[str]
    rows: list[list[str]]


def to_float(val: str) -> float | None:
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
        "ISTEP1": "A",
        "TSTEP1": "s",
        "SAMPLETIME": "s",
        "AREA": "cm^2",
    }
    return fallback_units.get(key, "")


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


def find_deg_files(input_dir: Path) -> list[DegFile]:
    files: list[DegFile] = []
    for path in sorted(Path(input_dir).glob("*.DTA")):
        match = DEG_FILE_RE.match(path.name)
        if not match:
            continue
        files.append(DegFile(path=path, stage=int(match.group("stage"))))
    return sorted(files, key=lambda item: item.stage)


def _column_index(parsed: ParsedDTA, column_name: str) -> int | None:
    try:
        return parsed.header.index(column_name)
    except ValueError:
        return None


def _required_numeric_series(parsed: ParsedDTA, x_name: str, y_name: str) -> tuple[list[float], list[float]]:
    x_idx = _column_index(parsed, x_name)
    y_idx = _column_index(parsed, y_name)
    if x_idx is None or y_idx is None:
        raise ValueError(f"Faltan columnas requeridas: {x_name}, {y_name}")

    x_values: list[float] = []
    y_values: list[float] = []

    for row in parsed.rows:
        if x_idx >= len(row) or y_idx >= len(row):
            continue
        x_val = to_float(row[x_idx])
        y_val = to_float(row[y_idx])
        if x_val is None or y_val is None:
            continue
        x_values.append(x_val)
        y_values.append(y_val)

    if not x_values or not y_values:
        raise ValueError("No se encontraron datos numericos validos para graficar.")

    return x_values, y_values


def _optional_float(value: str | None) -> float | None:
    if value is None:
        return None
    text = value.strip()
    if not text:
        return None
    return float(text.replace(",", "."))


def _positive_float(text: str, name: str) -> float:
    value = text.strip().replace(",", ".")
    if not value:
        raise ValueError(f"{name} no puede estar vacio.")
    num = float(value)
    if num <= 0:
        raise ValueError(f"{name} debe ser mayor que 0.")
    return num


def _time_unit_scale(time_unit: str) -> float:
    if time_unit == "min":
        return SECONDS_PER_MINUTE
    if time_unit == "h":
        return SECONDS_PER_HOUR
    return 1.0


def _time_unit_decimals(time_unit: str) -> int:
    if time_unit == "h":
        return 2
    if time_unit == "min":
        return 1
    return 1


def _dv_dt_scale(time_unit: str) -> float:
    return 1e6 * _time_unit_scale(time_unit)


def _dv_dt_unit_label(time_unit: str) -> str:
    return f"μV/{time_unit}" if time_unit != "s" else "μV/s"


def _mpl_linestyle(value: str) -> str:
    return "None" if value == "none" else value


def _build_scrollable_controls(parent) -> ttk.Frame:
    outer = ttk.Frame(parent, width=320)
    outer.pack(side="left", fill="y")
    outer.pack_propagate(False)

    canvas = tk.Canvas(outer, highlightthickness=0, width=320)
    scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    inner = ttk.Frame(canvas, padding=10)
    window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _update_scrollregion(_event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _sync_width(event):
        canvas.itemconfigure(window_id, width=event.width)

    inner.bind("<Configure>", _update_scrollregion)
    canvas.bind("<Configure>", _sync_width)

    def _on_mousewheel(event):
        if event.delta:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            step = -1 if event.num == 4 else 1
            canvas.yview_scroll(step, "units")

    def _bind_mousewheel(_event=None):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", _on_mousewheel)
        canvas.bind_all("<Button-5>", _on_mousewheel)

    def _unbind_mousewheel(_event=None):
        canvas.unbind_all("<MouseWheel>")
        canvas.unbind_all("<Button-4>")
        canvas.unbind_all("<Button-5>")

    outer.bind("<Enter>", _bind_mousewheel)
    outer.bind("<Leave>", _unbind_mousewheel)

    return inner


def _format_limit_value(value: float | None, decimals: int = 1) -> str:
    if value is None:
        return ""
    return f"{value:.{decimals}f}"


def _round_down_dec(value: float, decimals: int = 1) -> float:
    scale = 10 ** decimals
    return floor(value * scale) / scale


def _round_up_dec(value: float, decimals: int = 1) -> float:
    scale = 10 ** decimals
    return ceil(value * scale) / scale


def _padded_limits(
    values: list[float],
    rel_pad: float = 0.05,
    decimals: int = 1,
) -> tuple[float | None, float | None]:
    if not values:
        return None, None

    vmin = min(values)
    vmax = max(values)

    if vmin == vmax:
        pad = max(abs(vmin) * rel_pad, 1e-6)
    else:
        pad = (vmax - vmin) * rel_pad

    lo = _round_down_dec(vmin - pad, decimals)
    hi = _round_up_dec(vmax + pad, decimals)
    return lo, hi


def _positive_dv_dt_rows(rows: list[dict[str, float]]) -> list[dict[str, float]]:
    return [row for row in rows if row["dVdt"] > 0]


def _scale_compatible_dv_dt_rows(
    rows: list[dict[str, float]],
    logarithmic_y: bool,
) -> list[dict[str, float]]:
    return _positive_dv_dt_rows(rows) if logarithmic_y else rows


def _format_log_limit_value(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:.6g}"


def _format_adaptive_limit_value(value: float | None) -> str:
    if value is None:
        return ""
    return f"{value:.6g}"


def _adaptive_linear_limits(values: list[float]) -> tuple[float | None, float | None]:
    if not values:
        return None, None

    vmin = min(values)
    vmax = max(values)

    if vmin == vmax:
        pad = max(abs(vmin) * 0.05, 1e-9)
    else:
        pad = (vmax - vmin) * 0.05

    lo = vmin - pad
    hi = vmax + pad

    if lo == hi:
        hi = lo + max(abs(lo) * 0.05, 1e-9)

    return lo, hi


def _log_axis_limits(values: list[float]) -> tuple[float | None, float | None]:
    positive_values = [value for value in values if value > 0]
    if not positive_values:
        return None, None

    vmin = min(positive_values)
    vmax = max(positive_values)

    lower = 10 ** floor(log10(vmin))
    upper = 10 ** ceil(log10(vmax))

    if lower == upper:
        upper *= 10.0

    return lower, upper


def _normalize_median_window(window_size: int) -> int:
    size = max(1, int(window_size))
    if size % 2 == 0:
        size += 1
    return size


def _median_filter(values: list[float], window_size: int) -> list[float]:
    if not values:
        return []

    size = _normalize_median_window(window_size)
    if size == 1:
        return list(values)

    radius = size // 2
    filtered: list[float] = []

    for idx in range(len(values)):
        start = max(0, idx - radius)
        end = min(len(values), idx + radius + 1)
        window = sorted(values[start:end])
        filtered.append(window[len(window) // 2])

    return filtered


def _normalize_average_window(window_size: int) -> int:
    return max(1, int(window_size))


def _rolling_average_filter(values: list[float], window_size: int) -> list[float]:
    if not values:
        return []

    size = _normalize_average_window(window_size)
    if size == 1:
        return list(values)

    radius_left = (size - 1) // 2
    radius_right = size // 2
    filtered: list[float] = []

    for idx in range(len(values)):
        start = max(0, idx - radius_left)
        end = min(len(values), idx + radius_right + 1)
        window = values[start:end]
        filtered.append(sum(window) / len(window))

    return filtered


def apply_smoothing(values: list[float], algorithm: str, window_size: int) -> list[float]:
    if algorithm == "Median filter":
        return _median_filter(values, window_size)
    if algorithm == "Rolling average":
        return _rolling_average_filter(values, window_size)
    raise ValueError(f"Algoritmo de suavizado no soportado: {algorithm}")


def apply_x_edge_ticks(ax, x_min: float | None, x_max: float | None, tick_count: int) -> None:
    if x_min is None and x_max is None:
        return
    ax.xaxis.set_major_locator(LinearLocator(max(2, int(tick_count))))
    ax.xaxis.set_major_formatter(StrMethodFormatter("{x:g}"))


def compute_default_v_vs_t_limits(
    parsed_items: list[tuple[DegFile, ParsedDTA]],
    time_unit: str = "s",
) -> dict[str, str]:
    time_values: list[float] = []
    voltage_values: list[float] = []
    temperature_values: list[float] = []

    for _deg_file, parsed in parsed_items:
        t_vals, v_vals = _required_numeric_series(parsed, "T", "Vf")
        time_values.extend(t_vals)
        voltage_values.extend(v_vals)
        try:
            _temp_time, temp_vals = _required_numeric_series(parsed, "T", "Temp")
            temperature_values.extend(temp_vals)
        except ValueError:
            pass

    if not time_values or not voltage_values:
        return {"t_min": "", "t_max": "", "v_min": "", "v_max": "", "temp_min": "", "temp_max": ""}

    time_scale = _time_unit_scale(time_unit)
    time_decimals = _time_unit_decimals(time_unit)
    scaled_time_values = [value / time_scale for value in time_values]

    _t_min, t_max = _padded_limits(scaled_time_values, decimals=time_decimals)
    t_min = 0.0
    v_min, v_max = _padded_limits(voltage_values)
    temp_min, temp_max = _padded_limits(temperature_values) if temperature_values else (None, None)

    return {
        "t_min": _format_limit_value(t_min, time_decimals),
        "t_max": _format_limit_value(t_max, time_decimals),
        "v_min": _format_limit_value(v_min),
        "v_max": _format_limit_value(v_max),
        "temp_min": _format_limit_value(temp_min) if temp_min is not None else "",
        "temp_max": _format_limit_value(temp_max) if temp_max is not None else "",
    }


def compute_autofit_v_vs_t_limits(
    parsed_items: list[tuple[DegFile, ParsedDTA]],
    show_temperature: bool,
    time_unit: str = "s",
) -> dict[str, str]:
    time_values: list[float] = []
    voltage_values: list[float] = []
    temperature_values: list[float] = []

    for _deg_file, parsed in parsed_items:
        t_vals, v_vals = _required_numeric_series(parsed, "T", "Vf")
        time_values.extend(t_vals)
        voltage_values.extend(v_vals)
        if show_temperature:
            try:
                _temp_time, temp_vals = _required_numeric_series(parsed, "T", "Temp")
                temperature_values.extend(temp_vals)
            except ValueError:
                pass

    if not time_values or not voltage_values:
        raise ValueError("No hay datos validos para ajustar los ejes.")

    time_scale = _time_unit_scale(time_unit)
    time_decimals = _time_unit_decimals(time_unit)
    scaled_time_values = [value / time_scale for value in time_values]

    _t_min, t_max = _padded_limits(scaled_time_values, decimals=time_decimals)
    t_min = 0.0
    v_min, v_max = _padded_limits(voltage_values)

    out = {
        "t_min": _format_limit_value(t_min, time_decimals),
        "t_max": _format_limit_value(t_max, time_decimals),
        "v_min": _format_limit_value(v_min),
        "v_max": _format_limit_value(v_max),
        "temp_min": "",
        "temp_max": "",
    }
    if show_temperature and temperature_values:
        temp_min, temp_max = _padded_limits(temperature_values)
        out["temp_min"] = _format_limit_value(temp_min)
        out["temp_max"] = _format_limit_value(temp_max)
    return out


def build_dv_dt_rows(
    parsed: ParsedDTA,
    time_unit: str = "s",
    smoothing_algorithm: str = "Median filter",
    smoothing_window: int = 1,
) -> list[dict[str, float]]:
    time_values, voltage_values = _required_numeric_series(parsed, "T", "Vf")
    time_scale = _time_unit_scale(time_unit)
    dvdt_scale = _dv_dt_scale(time_unit)

    dvdt_time: list[float] = []
    dvdt_values: list[float] = []

    for idx in range(len(time_values) - 1):
        t0 = time_values[idx]
        t1 = time_values[idx + 1]
        v0 = voltage_values[idx]
        v1 = voltage_values[idx + 1]

        dt = t1 - t0
        if dt <= 0:
            continue

        dvdt_time.append(((t0 + t1) / 2.0) / time_scale)
        dvdt_values.append(((v1 - v0) / dt) * dvdt_scale)

    smoothed_dvdt = apply_smoothing(dvdt_values, smoothing_algorithm, smoothing_window)
    dvdt_rows = [
        {"time": time_value, "dVdt": dvdt_value}
        for time_value, dvdt_value in zip(dvdt_time, smoothed_dvdt)
    ]

    if not dvdt_rows:
        raise ValueError("No hay suficientes datos validos para calcular dV/dt.")

    return dvdt_rows


def compute_default_dv_dt_limits(
    parsed_items: list[tuple[DegFile, ParsedDTA]],
    time_unit: str = "s",
    smoothing_algorithm: str = "Median filter",
    smoothing_window: int = 1,
    logarithmic_y: bool = True,
    decimals: int = 3,
) -> dict[str, str]:
    time_values: list[float] = []
    dvdt_values: list[float] = []

    for _deg_file, parsed in parsed_items:
        rows = _scale_compatible_dv_dt_rows(
            build_dv_dt_rows(
                parsed,
                time_unit=time_unit,
                smoothing_algorithm=smoothing_algorithm,
                smoothing_window=smoothing_window,
            ),
            logarithmic_y=logarithmic_y,
        )
        time_values.extend(row["time"] for row in rows)
        dvdt_values.extend(row["dVdt"] for row in rows)

    if not time_values or not dvdt_values:
        return {
            "t_min": "",
            "t_max": "",
            "dvdt_min": "",
            "dvdt_max": "",
        }

    _t_min, t_max = _padded_limits(time_values, decimals=_time_unit_decimals(time_unit))
    t_min = 0.0
    if logarithmic_y:
        dvdt_min, dvdt_max = _log_axis_limits(dvdt_values)
    else:
        dvdt_min, dvdt_max = _adaptive_linear_limits(dvdt_values)

    return {
        "t_min": _format_limit_value(t_min, _time_unit_decimals(time_unit)),
        "t_max": _format_limit_value(t_max, _time_unit_decimals(time_unit)),
        "dvdt_min": _format_log_limit_value(dvdt_min) if logarithmic_y else _format_adaptive_limit_value(dvdt_min),
        "dvdt_max": _format_log_limit_value(dvdt_max) if logarithmic_y else _format_adaptive_limit_value(dvdt_max),
    }


def compute_autofit_dv_dt_limits(
    parsed_items: list[tuple[DegFile, ParsedDTA]],
    time_unit: str = "s",
    smoothing_algorithm: str = "Median filter",
    smoothing_window: int = 1,
    logarithmic_y: bool = True,
    locked_t_min: float | None = None,
    locked_t_max: float | None = None,
    locked_dvdt_min: float | None = None,
    locked_dvdt_max: float | None = None,
    decimals: int = 3,
) -> dict[str, str]:
    rows: list[dict[str, float]] = []

    for _deg_file, parsed in parsed_items:
        rows.extend(
            _scale_compatible_dv_dt_rows(
                build_dv_dt_rows(
                    parsed,
                    time_unit=time_unit,
                    smoothing_algorithm=smoothing_algorithm,
                    smoothing_window=smoothing_window,
                ),
                logarithmic_y=logarithmic_y,
            )
        )

    filtered_rows: list[dict[str, float]] = []
    for row in rows:
        if locked_t_min is not None and row["time"] < locked_t_min:
            continue
        if locked_t_max is not None and row["time"] > locked_t_max:
            continue
        if locked_dvdt_min is not None and row["dVdt"] < locked_dvdt_min:
            continue
        if locked_dvdt_max is not None and row["dVdt"] > locked_dvdt_max:
            continue
        filtered_rows.append(row)

    if not filtered_rows:
        raise ValueError("No hay datos validos dentro de los limites bloqueados para ajustar los ejes.")

    time_values = [row["time"] for row in filtered_rows]
    dvdt_values = [row["dVdt"] for row in filtered_rows]

    _t_min, t_max = _padded_limits(time_values, decimals=_time_unit_decimals(time_unit))
    t_min = 0.0
    if logarithmic_y:
        dvdt_min, dvdt_max = _log_axis_limits(dvdt_values)
    else:
        dvdt_min, dvdt_max = _padded_limits(dvdt_values, decimals=decimals)

    return {
        "t_min": _format_limit_value(t_min, _time_unit_decimals(time_unit)),
        "t_max": _format_limit_value(t_max, _time_unit_decimals(time_unit)),
        "dvdt_min": _format_log_limit_value(dvdt_min) if logarithmic_y else _format_adaptive_limit_value(dvdt_min),
        "dvdt_max": _format_log_limit_value(dvdt_max) if logarithmic_y else _format_adaptive_limit_value(dvdt_max),
    }


def draw_dv_dt_on_figure(
    fig: Figure,
    parsed_items: list[tuple[DegFile, ParsedDTA]],
    dvdt_linestyle: str,
    time_unit: str = "s",
    smoothing_algorithm: str = "Median filter",
    smoothing_window: int = 1,
    logarithmic_y: bool = True,
    x_tick_count: int = 6,
    y_tick_count: int = 6,
    t_min: float | None = None,
    t_max: float | None = None,
    dvdt_min: float | None = None,
    dvdt_max: float | None = None,
    plot_title: str = "",
    title_fontsize: float = 14,
    tick_fontsize: float = 10,
    label_fontsize: float = 11,
    legend_fontsize: float = 10,
    line_width: float = 1.5,
) -> bool:
    fig.clear()

    if not parsed_items:
        return False

    x_tick_count = max(2, int(x_tick_count))
    y_tick_count = max(2, int(y_tick_count))
    mpl_linestyle = _mpl_linestyle(dvdt_linestyle)

    if mpl_linestyle == "None":
        return False

    ax = fig.add_subplot(111)

    def _line_kwargs(color: str) -> dict:
        return {
            "color": color,
            "linestyle": mpl_linestyle,
            "linewidth": line_width,
        }

    for idx, (deg_file, parsed) in enumerate(parsed_items):
        delta_rows = _scale_compatible_dv_dt_rows(
            build_dv_dt_rows(
                parsed,
                time_unit=time_unit,
                smoothing_algorithm=smoothing_algorithm,
                smoothing_window=smoothing_window,
            ),
            logarithmic_y=logarithmic_y,
        )
        if not delta_rows:
            continue
        color = DEG_STAGE_COLORS[idx % len(DEG_STAGE_COLORS)]
        ax.plot(
            [row["time"] for row in delta_rows],
            [row["dVdt"] for row in delta_rows],
            label=f"Stage #{deg_file.stage} dV/dt",
            **_line_kwargs(color),
        )

    handles, labels = ax.get_legend_handles_labels()
    if not handles:
        fig.clear()
        return False

    default_title = "dV/dt vs t - Degradacion galvanostatica"
    final_title = plot_title.strip() if plot_title.strip() else default_title

    ax.set_xlabel(f"Tiempo [{time_unit}]", fontsize=label_fontsize)
    ax.set_ylabel(f"dV/dt [{_dv_dt_unit_label(time_unit)}]", fontsize=label_fontsize)
    ax.set_title(final_title, fontsize=title_fontsize)
    ax.grid(True, alpha=0.25)
    ax.tick_params(axis="both", labelsize=tick_fontsize)
    ax.xaxis.set_major_locator(MaxNLocator(nbins=x_tick_count))
    ax.xaxis.set_major_formatter(StrMethodFormatter("{x:g}"))

    if t_min is not None or t_max is not None:
        ax.set_xlim(left=t_min, right=t_max)
        apply_x_edge_ticks(ax, t_min, t_max, x_tick_count)

    if logarithmic_y:
        if dvdt_min is not None and dvdt_min <= 0:
            raise ValueError("dV/dt min debe ser mayor que 0 para usar escala logaritmica.")
        if dvdt_max is not None and dvdt_max <= 0:
            raise ValueError("dV/dt max debe ser mayor que 0 para usar escala logaritmica.")

        ax.set_yscale("log")
        if dvdt_min is not None or dvdt_max is not None:
            ax.set_ylim(bottom=dvdt_min, top=dvdt_max)

        ax.yaxis.set_major_locator(LogLocator(base=10.0, numticks=max(2, y_tick_count)))
        ax.yaxis.set_minor_locator(LogLocator(base=10.0, subs=tuple(range(2, 10))))
        ax.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
        ax.yaxis.set_minor_formatter(LogFormatterSciNotation(base=10.0, labelOnlyBase=True))
        ax.grid(True, which="minor", axis="y", alpha=0.25)
    else:
        if dvdt_min is not None or dvdt_max is not None:
            ax.set_ylim(bottom=dvdt_min, top=dvdt_max)
            ax.yaxis.set_major_locator(LinearLocator(y_tick_count))
        else:
            ax.yaxis.set_major_locator(MaxNLocator(nbins=y_tick_count))
        ax.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))

    ax.legend(handles, labels, fontsize=legend_fontsize)
    fig.tight_layout()
    return True


def draw_v_vs_t_on_figure(
    fig: Figure,
    parsed_items: list[tuple[DegFile, ParsedDTA]],
    show_temperature: bool,
    voltage_linestyle: str = "-",
    temperature_linestyle: str = "--",
    time_unit: str = "s",
    x_tick_count: int = 6,
    y_tick_count: int = 6,
    plot_title: str = "",
    title_fontsize: float = 14,
    tick_fontsize: float = 10,
    label_fontsize: float = 11,
    legend_fontsize: float = 10,
    line_width: float = 1.5,
    t_min: float | None = None,
    t_max: float | None = None,
    v_min: float | None = None,
    v_max: float | None = None,
    temp_min: float | None = None,
    temp_max: float | None = None,
) -> bool:
    fig.clear()

    if not parsed_items:
        return False

    has_voltage = _mpl_linestyle(voltage_linestyle) != "None"
    has_temperature = show_temperature and _mpl_linestyle(temperature_linestyle) != "None"
    if not has_voltage and not has_temperature:
        return False

    ax = fig.add_subplot(111)
    ax_temp = ax.twinx() if has_temperature else None
    time_scale = _time_unit_scale(time_unit)

    for idx, (deg_file, parsed) in enumerate(parsed_items):
        t_vals, v_vals = _required_numeric_series(parsed, "T", "Vf")
        plot_t_vals = [value / time_scale for value in t_vals]
        stage_color = DEG_STAGE_COLORS[idx % len(DEG_STAGE_COLORS)]
        if has_voltage:
            ax.plot(
                plot_t_vals,
                v_vals,
                color=stage_color,
                linewidth=line_width,
                linestyle=_mpl_linestyle(voltage_linestyle),
                label=f"Stage #{deg_file.stage} V",
            )
        if has_temperature and ax_temp is not None:
            temp_time, temp_vals = _required_numeric_series(parsed, "T", "Temp")
            plot_temp_time = [value / time_scale for value in temp_time]
            ax_temp.plot(
                plot_temp_time,
                temp_vals,
                color=DEG_PLOT_COLORS["temperature"] if len(parsed_items) == 1 else stage_color,
                linewidth=line_width,
                linestyle=_mpl_linestyle(temperature_linestyle),
                label=f"Stage #{deg_file.stage} T",
            )

    default_title = "V vs t - Degradacion galvanostatica"
    final_title = plot_title.strip() if plot_title.strip() else default_title

    ax.set_xlabel(f"Tiempo [{time_unit}]", fontsize=label_fontsize)
    if has_voltage:
        ax.set_ylabel("Voltaje [V]", fontsize=label_fontsize)
    ax.set_title(final_title, fontsize=title_fontsize)
    ax.grid(True, alpha=0.25)
    ax.tick_params(axis="both", labelsize=tick_fontsize)
    ax.xaxis.set_major_locator(MaxNLocator(nbins=max(2, int(x_tick_count))))
    ax.yaxis.set_major_locator(MaxNLocator(nbins=max(2, int(y_tick_count))))
    ax.xaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
    ax.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))

    if t_min is not None or t_max is not None:
        ax.set_xlim(left=t_min, right=t_max)
        apply_x_edge_ticks(ax, t_min, t_max, x_tick_count)
    if has_voltage and (v_min is not None or v_max is not None):
        ax.set_ylim(bottom=v_min, top=v_max)
        ax.yaxis.set_major_locator(LinearLocator(max(2, int(y_tick_count))))
        ax.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))

    handles, labels = ax.get_legend_handles_labels()
    if ax_temp is not None:
        ax_temp.tick_params(axis="y", labelsize=tick_fontsize)
        ax_temp.set_ylabel("Temperatura [°C]", fontsize=label_fontsize)
        if temp_min is not None or temp_max is not None:
            ax_temp.set_ylim(bottom=temp_min, top=temp_max)
            ax_temp.yaxis.set_major_locator(LinearLocator(max(2, int(y_tick_count))))
        else:
            ax_temp.yaxis.set_major_locator(MaxNLocator(nbins=max(2, int(y_tick_count))))
        ax_temp.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
        h2, l2 = ax_temp.get_legend_handles_labels()
        handles += h2
        labels += l2

    if has_voltage:
        ax.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))

    if handles:
        ax.legend(handles, labels, fontsize=legend_fontsize)
    fig.tight_layout()
    return True


def open_v_vs_t_window(input_dir: Path) -> None:
    deg_files = find_deg_files(Path(input_dir))
    if not deg_files:
        raise ValueError("No se encontraron archivos de degradacion validos.")

    parsed_items = [(deg_file, parse_gamry_dta(deg_file.path)) for deg_file in deg_files]
    default_limits = compute_default_v_vs_t_limits(parsed_items, time_unit="s")

    win = tk.Toplevel()
    win.title("Deg - V vs t")
    win.geometry("1200x720")

    controls_frame = _build_scrollable_controls(win)

    plot_outer = ttk.Frame(win, padding=10)
    plot_outer.pack(side="right", fill="both", expand=True)

    toolbar_frame = ttk.Frame(plot_outer)
    toolbar_frame.pack(side="top", fill="x")

    canvas_frame = ttk.Frame(plot_outer)
    canvas_frame.pack(side="top", fill="both", expand=True)

    fig = Figure(figsize=(8.8, 5.4), dpi=100)

    canvas = FigureCanvasTkAgg(fig, master=canvas_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)

    toolbar = NavigationToolbar2Tk(canvas, toolbar_frame, pack_toolbar=False)
    toolbar.update()
    toolbar.pack(side="left", fill="x")

    status_var = tk.StringVar(value="Listo.")
    temperature_var = tk.BooleanVar(value=False)
    time_unit_var = tk.StringVar(value="s")
    voltage_line_var = tk.StringVar(value="-")
    temperature_line_var = tk.StringVar(value="--")
    x_tick_count_var = tk.IntVar(value=6)
    y_tick_count_var = tk.IntVar(value=6)
    plot_title_var = tk.StringVar(value="")
    title_fontsize_var = tk.StringVar(value="14")
    tick_fontsize_var = tk.StringVar(value="10")
    label_fontsize_var = tk.StringVar(value="11")
    legend_fontsize_var = tk.StringVar(value="10")
    line_width_var = tk.StringVar(value="1.5")
    t_min_var = tk.StringVar(value=default_limits["t_min"])
    t_max_var = tk.StringVar(value=default_limits["t_max"])
    v_min_var = tk.StringVar(value=default_limits["v_min"])
    v_max_var = tk.StringVar(value=default_limits["v_max"])
    temp_min_var = tk.StringVar(value=default_limits["temp_min"])
    temp_max_var = tk.StringVar(value=default_limits["temp_max"])

    initial_state = {
        "temperature": False,
        "time_unit": "s",
        "voltage_line": "-",
        "temperature_line": "--",
        "x_tick_count": 6,
        "y_tick_count": 6,
        "plot_title": "",
        "title_fontsize": "14",
        "tick_fontsize": "10",
        "label_fontsize": "11",
        "legend_fontsize": "10",
        "line_width": "1.5",
        "t_min": default_limits["t_min"],
        "t_max": default_limits["t_max"],
        "v_min": default_limits["v_min"],
        "v_max": default_limits["v_max"],
        "temp_min": default_limits["temp_min"],
        "temp_max": default_limits["temp_max"],
    }

    plot_job = {"id": None}
    current_time_unit = {"value": time_unit_var.get()}
    suspend_events = {"value": False}

    def _collect_limits():
        return dict(
            t_min=_optional_float(t_min_var.get()),
            t_max=_optional_float(t_max_var.get()),
            v_min=_optional_float(v_min_var.get()),
            v_max=_optional_float(v_max_var.get()),
            temp_min=_optional_float(temp_min_var.get()),
            temp_max=_optional_float(temp_max_var.get()),
        )

    def _format_time_limit(value: float | None, time_unit: str) -> str:
        if value is None:
            return ""
        return _format_limit_value(value, _time_unit_decimals(time_unit))

    def _convert_time_limit_text(value_text: str, from_unit: str, to_unit: str) -> str:
        value = _optional_float(value_text)
        if value is None or from_unit == to_unit:
            return value_text

        seconds = value * _time_unit_scale(from_unit)
        converted = seconds / _time_unit_scale(to_unit)
        return _format_time_limit(converted, to_unit)

    def _plot():
        plot_job["id"] = None
        try:
            has_plot = draw_v_vs_t_on_figure(
                fig=fig,
                parsed_items=parsed_items,
                show_temperature=temperature_var.get(),
                voltage_linestyle=voltage_line_var.get(),
                temperature_linestyle=temperature_line_var.get(),
                time_unit=time_unit_var.get(),
                x_tick_count=x_tick_count_var.get(),
                y_tick_count=y_tick_count_var.get(),
                plot_title=plot_title_var.get(),
                title_fontsize=_positive_float(title_fontsize_var.get(), "Title size"),
                tick_fontsize=_positive_float(tick_fontsize_var.get(), "Tick size"),
                label_fontsize=_positive_float(label_fontsize_var.get(), "Label size"),
                legend_fontsize=_positive_float(legend_fontsize_var.get(), "Legend size"),
                line_width=_positive_float(line_width_var.get(), "Line width"),
                **_collect_limits(),
            )
        except ValueError as exc:
            fig.clear()
            canvas.draw_idle()
            status_var.set(f"Error: {exc}")
            return

        if not has_plot:
            fig.clear()
            canvas.draw_idle()
            status_var.set("No se muestra grafico: active al menos una serie.")
            return

        canvas.draw_idle()
        status_var.set("Grafico actualizado.")

    def _schedule_plot(*_args):
        if suspend_events["value"]:
            return
        if plot_job["id"] is not None:
            win.after_cancel(plot_job["id"])
        plot_job["id"] = win.after(20, _plot)

    def _autofit():
        try:
            fitted = compute_autofit_v_vs_t_limits(
                parsed_items,
                temperature_var.get(),
                time_unit=time_unit_var.get(),
            )
        except ValueError as exc:
            status_var.set(f"Error: {exc}")
            return

        suspend_events["value"] = True
        try:
            t_min_var.set(fitted["t_min"])
            t_max_var.set(fitted["t_max"])
            v_min_var.set(fitted["v_min"])
            v_max_var.set(fitted["v_max"])
            temp_min_var.set(fitted["temp_min"])
            temp_max_var.set(fitted["temp_max"])
        finally:
            suspend_events["value"] = False

        _plot()
        status_var.set("Autoscale aplicado.")

    def _on_time_unit_changed(*_args):
        if suspend_events["value"]:
            return

        old_unit = current_time_unit["value"]
        new_unit = time_unit_var.get()

        suspend_events["value"] = True
        try:
            if t_min_var.get().strip():
                t_min_var.set(_convert_time_limit_text(t_min_var.get(), old_unit, new_unit))
            if t_max_var.get().strip():
                t_max_var.set(_convert_time_limit_text(t_max_var.get(), old_unit, new_unit))
            current_time_unit["value"] = new_unit
        finally:
            suspend_events["value"] = False

        _schedule_plot()

    def _reset():
        suspend_events["value"] = True
        try:
            temperature_var.set(initial_state["temperature"])
            current_time_unit["value"] = initial_state["time_unit"]
            time_unit_var.set(initial_state["time_unit"])
            voltage_line_var.set(initial_state["voltage_line"])
            temperature_line_var.set(initial_state["temperature_line"])
            x_tick_count_var.set(initial_state["x_tick_count"])
            y_tick_count_var.set(initial_state["y_tick_count"])
            plot_title_var.set(initial_state["plot_title"])
            title_fontsize_var.set(initial_state["title_fontsize"])
            tick_fontsize_var.set(initial_state["tick_fontsize"])
            label_fontsize_var.set(initial_state["label_fontsize"])
            legend_fontsize_var.set(initial_state["legend_fontsize"])
            line_width_var.set(initial_state["line_width"])
            t_min_var.set(initial_state["t_min"])
            t_max_var.set(initial_state["t_max"])
            v_min_var.set(initial_state["v_min"])
            v_max_var.set(initial_state["v_max"])
            temp_min_var.set(initial_state["temp_min"])
            temp_max_var.set(initial_state["temp_max"])
        finally:
            suspend_events["value"] = False

        _plot()
        status_var.set("Valores restaurados.")

    def _open_composer_placeholder():
        existing = getattr(win, "_composer_win", None)
        if existing is not None and existing.winfo_exists():
            existing.lift()
            existing.focus_force()
            return

        comp = tk.Toplevel(win)
        win._composer_win = comp  # type: ignore[attr-defined]
        comp.title("Composer (Deg)")
        comp.geometry("520x260")

        body = ttk.Frame(comp, padding=16)
        body.pack(fill="both", expand=True)

        ttk.Label(
            body,
            text="Composer placeholder",
            font=("TkDefaultFont", 11, "bold"),
        ).pack(anchor="w", pady=(0, 10))

        ttk.Label(
            body,
            text=(
                "Este espacio se reservará para componer curvas de degradación "
                "y superponer distintas etapas en un único gráfico."
            ),
            wraplength=460,
            justify="left",
        ).pack(anchor="w", pady=(0, 10))

        ttk.Label(
            body,
            text=(
                f"Etapas detectadas actualmente: "
                f"{', '.join(f'#{item.stage}' for item in deg_files)}"
            ),
            wraplength=460,
            justify="left",
        ).pack(anchor="w", pady=(0, 14))

        ttk.Button(body, text="Cerrar", command=comp.destroy).pack(anchor="e")

        def _on_close():
            comp.destroy()
            try:
                delattr(win, "_composer_win")
            except Exception:
                pass

        comp.protocol("WM_DELETE_WINDOW", _on_close)

    ttk.Label(
        controls_frame,
        text=f"Etapas detectadas: {', '.join(f'#{item.stage}' for item in deg_files)}",
        justify="left",
        wraplength=260,
    ).pack(anchor="w", pady=(0, 10))

    series_box = ttk.LabelFrame(controls_frame, text="Series")
    series_box.pack(fill="x", pady=5)

    style_box = ttk.LabelFrame(controls_frame, text="Estilo")
    style_box.pack(fill="x", pady=5)

    text_box = ttk.LabelFrame(controls_frame, text="Texto / tamanos")
    text_box.pack(fill="x", pady=5)

    limits_box = ttk.LabelFrame(controls_frame, text="Limites de ejes")
    limits_box.pack(fill="x", pady=5)

    ttk.Checkbutton(
        series_box,
        text="Temperatura",
        variable=temperature_var,
        command=_schedule_plot,
    ).pack(anchor="w", padx=8, pady=4)
    ttk.Label(series_box, text="Unidad tiempo").pack(anchor="w", padx=8, pady=(8, 2))
    time_unit_combo = ttk.Combobox(
        series_box,
        textvariable=time_unit_var,
        values=TIME_UNIT_OPTIONS,
        state="readonly",
        width=8,
    )
    time_unit_combo.pack(anchor="w", padx=8, pady=(0, 4))

    ttk.Label(style_box, text="Voltaje line").grid(row=0, column=0, sticky="w", padx=8, pady=3)
    voltage_line_combo = ttk.Combobox(
        style_box,
        textvariable=voltage_line_var,
        values=LINESTYLE_OPTIONS,
        state="readonly",
        width=10,
    )
    voltage_line_combo.grid(row=0, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(style_box, text="Temperatura line").grid(row=1, column=0, sticky="w", padx=8, pady=3)
    temperature_line_combo = ttk.Combobox(
        style_box,
        textvariable=temperature_line_var,
        values=LINESTYLE_OPTIONS,
        state="readonly",
        width=10,
    )
    temperature_line_combo.grid(row=1, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Titulo").grid(row=0, column=0, sticky="w", padx=8, pady=3)
    title_entry = ttk.Entry(text_box, textvariable=plot_title_var, width=28)
    title_entry.grid(row=0, column=1, sticky="we", padx=8, pady=3)

    ttk.Label(text_box, text="Title size").grid(row=1, column=0, sticky="w", padx=8, pady=3)
    title_size_spin = tk.Spinbox(text_box, from_=6, to=30, increment=0.5, textvariable=title_fontsize_var, width=8)
    title_size_spin.grid(row=1, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Tick size").grid(row=2, column=0, sticky="w", padx=8, pady=3)
    tick_size_spin = tk.Spinbox(text_box, from_=6, to=24, increment=0.5, textvariable=tick_fontsize_var, width=8)
    tick_size_spin.grid(row=2, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Label size").grid(row=3, column=0, sticky="w", padx=8, pady=3)
    label_size_spin = tk.Spinbox(text_box, from_=6, to=24, increment=0.5, textvariable=label_fontsize_var, width=8)
    label_size_spin.grid(row=3, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Legend size").grid(row=4, column=0, sticky="w", padx=8, pady=3)
    legend_size_spin = tk.Spinbox(text_box, from_=6, to=24, increment=0.5, textvariable=legend_fontsize_var, width=8)
    legend_size_spin.grid(row=4, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Line width").grid(row=5, column=0, sticky="w", padx=8, pady=3)
    line_width_spin = tk.Spinbox(text_box, from_=0.5, to=5.0, increment=0.1, textvariable=line_width_var, width=8)
    line_width_spin.grid(row=5, column=1, sticky="w", padx=8, pady=3)

    limit_specs = [
        ("t min", t_min_var),
        ("t max", t_max_var),
        ("V min", v_min_var),
        ("V max", v_max_var),
        ("T min", temp_min_var),
        ("T max", temp_max_var),
    ]

    for row_idx, (label, var) in enumerate(limit_specs):
        ttk.Label(limits_box, text=label).grid(row=row_idx, column=0, sticky="w", padx=8, pady=3)
        entry = ttk.Entry(limits_box, textvariable=var, width=12)
        entry.grid(row=row_idx, column=1, sticky="w", padx=8, pady=3)
        entry.bind("<Return>", _schedule_plot)
        entry.bind("<KP_Enter>", _schedule_plot)
        entry.bind("<FocusOut>", _schedule_plot)

    ttk.Label(limits_box, text="x-Ticks").grid(row=len(limit_specs), column=0, sticky="w", padx=8, pady=3)
    x_tick_spin = tk.Spinbox(limits_box, from_=2, to=10, textvariable=x_tick_count_var, width=8)
    x_tick_spin.grid(row=len(limit_specs), column=1, sticky="w", padx=8, pady=3)

    ttk.Label(limits_box, text="y-Ticks").grid(row=len(limit_specs) + 1, column=0, sticky="w", padx=8, pady=3)
    y_tick_spin = tk.Spinbox(limits_box, from_=2, to=10, textvariable=y_tick_count_var, width=8)
    y_tick_spin.grid(row=len(limit_specs) + 1, column=1, sticky="w", padx=8, pady=3)

    for combo in (voltage_line_combo, temperature_line_combo):
        combo.bind("<<ComboboxSelected>>", _schedule_plot)
    time_unit_combo.bind("<<ComboboxSelected>>", _on_time_unit_changed)
    for spin in (x_tick_spin, y_tick_spin):
        spin.bind("<Return>", _schedule_plot)
        spin.bind("<FocusOut>", _schedule_plot)
        spin.config(command=_schedule_plot)

    for widget in (
        title_entry,
        title_size_spin,
        tick_size_spin,
        label_size_spin,
        legend_size_spin,
        line_width_spin,
    ):
        widget.bind("<Return>", _schedule_plot)
        widget.bind("<KP_Enter>", _schedule_plot)
        widget.bind("<FocusOut>", _schedule_plot)

    for spin in (
        title_size_spin,
        tick_size_spin,
        label_size_spin,
        legend_size_spin,
        line_width_spin,
    ):
        spin.config(command=_schedule_plot)

    ttk.Label(
        controls_frame,
        textvariable=status_var,
        wraplength=240,
        justify="left",
    ).pack(anchor="w", fill="x", pady=(10, 10))

    buttons_frame = ttk.Frame(controls_frame)
    buttons_frame.pack(fill="x", pady=(5, 0))

    ttk.Button(buttons_frame, text="Reset", command=_reset).pack(side="left", padx=(0, 6))
    ttk.Button(buttons_frame, text="Composer", command=_open_composer_placeholder).pack(side="left", padx=(0, 6))
    ttk.Button(buttons_frame, text="Autoscale", command=_autofit).pack(side="left")

    _plot()


def open_dv_dt_window(input_dir: Path) -> None:
    deg_files = find_deg_files(Path(input_dir))
    if not deg_files:
        raise ValueError("No se encontraron archivos de degradacion validos.")

    parsed_items = [(deg_file, parse_gamry_dta(deg_file.path)) for deg_file in deg_files]
    default_limits = compute_default_dv_dt_limits(
        parsed_items,
        time_unit="s",
        smoothing_algorithm="Median filter",
        smoothing_window=1,
        logarithmic_y=True,
    )

    win = tk.Toplevel()
    win.title("Deg - dV/dt")
    win.geometry("1200x720")

    controls_frame = _build_scrollable_controls(win)

    plot_outer = ttk.Frame(win, padding=10)
    plot_outer.pack(side="right", fill="both", expand=True)

    toolbar_frame = ttk.Frame(plot_outer)
    toolbar_frame.pack(side="top", fill="x")

    canvas_frame = ttk.Frame(plot_outer)
    canvas_frame.pack(side="top", fill="both", expand=True)

    fig = Figure(figsize=(8.8, 5.4), dpi=100)

    canvas = FigureCanvasTkAgg(fig, master=canvas_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)

    toolbar = NavigationToolbar2Tk(canvas, toolbar_frame, pack_toolbar=False)
    toolbar.update()
    toolbar.pack(side="left", fill="x")

    status_var = tk.StringVar(value="Listo.")
    time_unit_var = tk.StringVar(value="s")
    logarithmic_y_var = tk.BooleanVar(value=True)
    dvdt_line_var = tk.StringVar(value="-")
    smoothing_algorithm_var = tk.StringVar(value="Median filter")
    smoothing_window_var = tk.IntVar(value=1)
    x_tick_count_var = tk.IntVar(value=6)
    y_tick_count_var = tk.IntVar(value=6)
    plot_title_var = tk.StringVar(value="")
    title_fontsize_var = tk.StringVar(value="14")
    tick_fontsize_var = tk.StringVar(value="10")
    label_fontsize_var = tk.StringVar(value="11")
    legend_fontsize_var = tk.StringVar(value="10")
    line_width_var = tk.StringVar(value="1.5")
    t_min_var = tk.StringVar(value=default_limits["t_min"])
    t_max_var = tk.StringVar(value=default_limits["t_max"])
    dvdt_min_var = tk.StringVar(value=default_limits["dvdt_min"])
    dvdt_max_var = tk.StringVar(value=default_limits["dvdt_max"])
    t_min_lock_var = tk.BooleanVar(value=False)
    t_max_lock_var = tk.BooleanVar(value=False)
    dvdt_min_lock_var = tk.BooleanVar(value=False)
    dvdt_max_lock_var = tk.BooleanVar(value=False)

    initial_state = {
        "time_unit": "s",
        "logarithmic_y": True,
        "dvdt_line": "-",
        "smoothing_algorithm": "Median filter",
        "smoothing_window": 1,
        "x_tick_count": 6,
        "y_tick_count": 6,
        "plot_title": "",
        "title_fontsize": "14",
        "tick_fontsize": "10",
        "label_fontsize": "11",
        "legend_fontsize": "10",
        "line_width": "1.5",
        "t_min": default_limits["t_min"],
        "t_max": default_limits["t_max"],
        "dvdt_min": default_limits["dvdt_min"],
        "dvdt_max": default_limits["dvdt_max"],
        "t_min_lock": False,
        "t_max_lock": False,
        "dvdt_min_lock": False,
        "dvdt_max_lock": False,
    }

    plot_job = {"id": None}
    current_time_unit = {"value": time_unit_var.get()}
    suspend_events = {"value": False}
    limit_entries: dict[str, ttk.Entry] = {}

    def _smoothing_config() -> dict[str, object]:
        return {
            "smoothing_algorithm": smoothing_algorithm_var.get(),
            "smoothing_window": smoothing_window_var.get(),
        }

    def _axis_mode_config() -> dict[str, object]:
        return {
            "logarithmic_y": logarithmic_y_var.get(),
        }

    def _collect_limits():
        return dict(
            t_min=_optional_float(t_min_var.get()),
            t_max=_optional_float(t_max_var.get()),
            dvdt_min=_optional_float(dvdt_min_var.get()),
            dvdt_max=_optional_float(dvdt_max_var.get()),
        )

    def _apply_lock_states() -> None:
        states = {
            "t_min": t_min_lock_var.get(),
            "t_max": t_max_lock_var.get(),
            "dvdt_min": dvdt_min_lock_var.get(),
            "dvdt_max": dvdt_max_lock_var.get(),
        }
        for key, entry in limit_entries.items():
            entry.state(["disabled"] if states[key] else ["!disabled"])

    def _format_time_limit(value: float | None, time_unit: str) -> str:
        if value is None:
            return ""
        return _format_limit_value(value, _time_unit_decimals(time_unit))

    def _convert_time_limit_text(value_text: str, from_unit: str, to_unit: str) -> str:
        value = _optional_float(value_text)
        if value is None or from_unit == to_unit:
            return value_text
        seconds = value * _time_unit_scale(from_unit)
        converted = seconds / _time_unit_scale(to_unit)
        return _format_time_limit(converted, to_unit)

    def _convert_dvdt_limit_text(value_text: str, from_unit: str, to_unit: str) -> str:
        value = _optional_float(value_text)
        if value is None or from_unit == to_unit:
            return value_text
        base_v_per_s = value / _dv_dt_scale(from_unit)
        converted = base_v_per_s * _dv_dt_scale(to_unit)
        return _format_log_limit_value(converted) if logarithmic_y_var.get() else _format_adaptive_limit_value(converted)

    def _plot():
        plot_job["id"] = None
        try:
            has_plot = draw_dv_dt_on_figure(
                fig=fig,
                parsed_items=parsed_items,
                dvdt_linestyle=dvdt_line_var.get(),
                time_unit=time_unit_var.get(),
                x_tick_count=x_tick_count_var.get(),
                y_tick_count=y_tick_count_var.get(),
                plot_title=plot_title_var.get(),
                title_fontsize=_positive_float(title_fontsize_var.get(), "Title size"),
                tick_fontsize=_positive_float(tick_fontsize_var.get(), "Tick size"),
                label_fontsize=_positive_float(label_fontsize_var.get(), "Label size"),
                legend_fontsize=_positive_float(legend_fontsize_var.get(), "Legend size"),
                line_width=_positive_float(line_width_var.get(), "Line width"),
                **_smoothing_config(),
                **_axis_mode_config(),
                **_collect_limits(),
            )
        except ValueError as exc:
            fig.clear()
            canvas.draw_idle()
            status_var.set(f"Error: {exc}")
            return

        if not has_plot:
            fig.clear()
            canvas.draw_idle()
            status_var.set("No se muestra grafico: active una linea visible.")
            return

        canvas.draw_idle()
        status_var.set("Grafico actualizado.")

    def _schedule_plot(*_args):
        if suspend_events["value"]:
            return
        if plot_job["id"] is not None:
            win.after_cancel(plot_job["id"])
        plot_job["id"] = win.after(20, _plot)

    def _autofit():
        try:
            fitted = compute_autofit_dv_dt_limits(
                parsed_items,
                time_unit=time_unit_var.get(),
                **_smoothing_config(),
                **_axis_mode_config(),
                locked_t_min=_optional_float(t_min_var.get()) if t_min_lock_var.get() else None,
                locked_t_max=_optional_float(t_max_var.get()) if t_max_lock_var.get() else None,
                locked_dvdt_min=_optional_float(dvdt_min_var.get()) if dvdt_min_lock_var.get() else None,
                locked_dvdt_max=_optional_float(dvdt_max_var.get()) if dvdt_max_lock_var.get() else None,
            )
        except ValueError as exc:
            status_var.set(f"Error: {exc}")
            return

        suspend_events["value"] = True
        try:
            if not t_min_lock_var.get():
                t_min_var.set(fitted["t_min"])
            if not t_max_lock_var.get():
                t_max_var.set(fitted["t_max"])
            if not dvdt_min_lock_var.get():
                dvdt_min_var.set(fitted["dvdt_min"])
            if not dvdt_max_lock_var.get():
                dvdt_max_var.set(fitted["dvdt_max"])
        finally:
            suspend_events["value"] = False

        _plot()
        status_var.set("Autoscale aplicado.")

    def _on_time_unit_changed(*_args):
        if suspend_events["value"]:
            return

        old_unit = current_time_unit["value"]
        new_unit = time_unit_var.get()

        suspend_events["value"] = True
        try:
            if not t_min_lock_var.get() and t_min_var.get().strip():
                t_min_var.set(_convert_time_limit_text(t_min_var.get(), old_unit, new_unit))
            if not t_max_lock_var.get() and t_max_var.get().strip():
                t_max_var.set(_convert_time_limit_text(t_max_var.get(), old_unit, new_unit))
            if not dvdt_min_lock_var.get() and dvdt_min_var.get().strip():
                dvdt_min_var.set(_convert_dvdt_limit_text(dvdt_min_var.get(), old_unit, new_unit))
            if not dvdt_max_lock_var.get() and dvdt_max_var.get().strip():
                dvdt_max_var.set(_convert_dvdt_limit_text(dvdt_max_var.get(), old_unit, new_unit))
            current_time_unit["value"] = new_unit
        finally:
            suspend_events["value"] = False

        _schedule_plot()

    def _reset():
        suspend_events["value"] = True
        try:
            current_time_unit["value"] = initial_state["time_unit"]
            time_unit_var.set(initial_state["time_unit"])
            logarithmic_y_var.set(initial_state["logarithmic_y"])
            dvdt_line_var.set(initial_state["dvdt_line"])
            smoothing_algorithm_var.set(initial_state["smoothing_algorithm"])
            smoothing_window_var.set(initial_state["smoothing_window"])
            x_tick_count_var.set(initial_state["x_tick_count"])
            y_tick_count_var.set(initial_state["y_tick_count"])
            plot_title_var.set(initial_state["plot_title"])
            title_fontsize_var.set(initial_state["title_fontsize"])
            tick_fontsize_var.set(initial_state["tick_fontsize"])
            label_fontsize_var.set(initial_state["label_fontsize"])
            legend_fontsize_var.set(initial_state["legend_fontsize"])
            line_width_var.set(initial_state["line_width"])
            t_min_var.set(initial_state["t_min"])
            t_max_var.set(initial_state["t_max"])
            dvdt_min_var.set(initial_state["dvdt_min"])
            dvdt_max_var.set(initial_state["dvdt_max"])
            t_min_lock_var.set(initial_state["t_min_lock"])
            t_max_lock_var.set(initial_state["t_max_lock"])
            dvdt_min_lock_var.set(initial_state["dvdt_min_lock"])
            dvdt_max_lock_var.set(initial_state["dvdt_max_lock"])
            _apply_lock_states()
        finally:
            suspend_events["value"] = False

        _plot()
        status_var.set("Valores restaurados.")

    ttk.Label(
        controls_frame,
        text=f"Etapas detectadas: {', '.join(f'#{item.stage}' for item in deg_files)}",
        justify="left",
        wraplength=260,
    ).pack(anchor="w", pady=(0, 10))

    series_box = ttk.LabelFrame(controls_frame, text="Series")
    series_box.pack(fill="x", pady=5)

    style_box = ttk.LabelFrame(controls_frame, text="Estilo")
    style_box.pack(fill="x", pady=5)

    smoothing_box = ttk.LabelFrame(controls_frame, text="Smoothing")
    smoothing_box.pack(fill="x", pady=5)

    text_box = ttk.LabelFrame(controls_frame, text="Texto / tamanos")
    text_box.pack(fill="x", pady=5)

    limits_box = ttk.LabelFrame(controls_frame, text="Limites de ejes")
    limits_box.pack(fill="x", pady=5)

    ttk.Label(series_box, text="Unidad tiempo").pack(anchor="w", padx=8, pady=(4, 2))
    time_unit_combo = ttk.Combobox(
        series_box,
        textvariable=time_unit_var,
        values=TIME_UNIT_OPTIONS,
        state="readonly",
        width=8,
    )
    time_unit_combo.pack(anchor="w", padx=8, pady=(0, 4))
    ttk.Checkbutton(
        series_box,
        text="Logarithmic y-axis",
        variable=logarithmic_y_var,
        command=_schedule_plot,
    ).pack(anchor="w", padx=8, pady=(4, 4))

    ttk.Label(style_box, text="dV/dt line").grid(row=0, column=0, sticky="w", padx=8, pady=3)
    dvdt_line_combo = ttk.Combobox(
        style_box,
        textvariable=dvdt_line_var,
        values=LINESTYLE_OPTIONS,
        state="readonly",
        width=10,
    )
    dvdt_line_combo.grid(row=0, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(smoothing_box, text="Algorithm").grid(row=0, column=0, sticky="w", padx=8, pady=3)
    smoothing_algo_combo = ttk.Combobox(
        smoothing_box,
        textvariable=smoothing_algorithm_var,
        values=SMOOTHING_ALGORITHMS,
        state="readonly",
        width=16,
    )
    smoothing_algo_combo.grid(row=0, column=1, sticky="w", padx=8, pady=3)

    smoothing_window_label = ttk.Label(smoothing_box, text="Median window")
    smoothing_window_label.grid(row=1, column=0, sticky="w", padx=8, pady=3)
    smoothing_window_spin = tk.Spinbox(
        smoothing_box,
        from_=1,
        to=500,
        increment=1,
        textvariable=smoothing_window_var,
        width=8,
    )
    smoothing_window_spin.grid(row=1, column=1, sticky="w", padx=8, pady=3)

    def _update_smoothing_ui(*_args):
        algorithm = smoothing_algorithm_var.get()
        if algorithm == "Rolling average":
            smoothing_window_label.config(text="Average window")
        else:
            smoothing_window_label.config(text="Median window")
        _schedule_plot()

    ttk.Label(text_box, text="Titulo").grid(row=0, column=0, sticky="w", padx=8, pady=3)
    title_entry = ttk.Entry(text_box, textvariable=plot_title_var, width=28)
    title_entry.grid(row=0, column=1, sticky="we", padx=8, pady=3)

    ttk.Label(text_box, text="Title size").grid(row=1, column=0, sticky="w", padx=8, pady=3)
    title_size_spin = tk.Spinbox(text_box, from_=6, to=30, increment=0.5, textvariable=title_fontsize_var, width=8)
    title_size_spin.grid(row=1, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Tick size").grid(row=2, column=0, sticky="w", padx=8, pady=3)
    tick_size_spin = tk.Spinbox(text_box, from_=6, to=24, increment=0.5, textvariable=tick_fontsize_var, width=8)
    tick_size_spin.grid(row=2, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Label size").grid(row=3, column=0, sticky="w", padx=8, pady=3)
    label_size_spin = tk.Spinbox(text_box, from_=6, to=24, increment=0.5, textvariable=label_fontsize_var, width=8)
    label_size_spin.grid(row=3, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Legend size").grid(row=4, column=0, sticky="w", padx=8, pady=3)
    legend_size_spin = tk.Spinbox(text_box, from_=6, to=24, increment=0.5, textvariable=legend_fontsize_var, width=8)
    legend_size_spin.grid(row=4, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(text_box, text="Line width").grid(row=5, column=0, sticky="w", padx=8, pady=3)
    line_width_spin = tk.Spinbox(text_box, from_=0.0, to=10.0, increment=0.1, textvariable=line_width_var, width=8)
    line_width_spin.grid(row=5, column=1, sticky="w", padx=8, pady=3)

    limit_specs = [
        ("t min", "t_min", t_min_var, t_min_lock_var),
        ("t max", "t_max", t_max_var, t_max_lock_var),
        ("dV/dt min", "dvdt_min", dvdt_min_var, dvdt_min_lock_var),
        ("dV/dt max", "dvdt_max", dvdt_max_var, dvdt_max_lock_var),
    ]

    for row_idx, (label, key, var, lock_var) in enumerate(limit_specs):
        ttk.Label(limits_box, text=label).grid(row=row_idx, column=0, sticky="w", padx=8, pady=3)
        entry = ttk.Entry(limits_box, textvariable=var, width=12)
        entry.grid(row=row_idx, column=1, sticky="w", padx=8, pady=3)
        entry.bind("<Return>", _schedule_plot)
        entry.bind("<KP_Enter>", _schedule_plot)
        entry.bind("<FocusOut>", _schedule_plot)
        limit_entries[key] = entry
        ttk.Checkbutton(
            limits_box,
            text="Lock",
            variable=lock_var,
            command=_apply_lock_states,
        ).grid(row=row_idx, column=2, sticky="w", padx=8, pady=3)

    ttk.Label(limits_box, text="x-Ticks").grid(row=len(limit_specs), column=0, sticky="w", padx=8, pady=3)
    x_tick_spin = tk.Spinbox(limits_box, from_=2, to=10, textvariable=x_tick_count_var, width=8)
    x_tick_spin.grid(row=len(limit_specs), column=1, sticky="w", padx=8, pady=3)

    ttk.Label(limits_box, text="y-Ticks").grid(row=len(limit_specs) + 1, column=0, sticky="w", padx=8, pady=3)
    y_tick_spin = tk.Spinbox(limits_box, from_=2, to=10, textvariable=y_tick_count_var, width=8)
    y_tick_spin.grid(row=len(limit_specs) + 1, column=1, sticky="w", padx=8, pady=3)

    dvdt_line_combo.bind("<<ComboboxSelected>>", _schedule_plot)
    time_unit_combo.bind("<<ComboboxSelected>>", _on_time_unit_changed)
    smoothing_algo_combo.bind("<<ComboboxSelected>>", _update_smoothing_ui)

    for spin in (smoothing_window_spin, x_tick_spin, y_tick_spin):
        spin.bind("<Return>", _schedule_plot)
        spin.bind("<FocusOut>", _schedule_plot)
        spin.config(command=_schedule_plot)

    for widget in (
        title_entry,
        title_size_spin,
        tick_size_spin,
        label_size_spin,
        legend_size_spin,
        line_width_spin,
    ):
        widget.bind("<Return>", _schedule_plot)
        widget.bind("<KP_Enter>", _schedule_plot)
        widget.bind("<FocusOut>", _schedule_plot)

    for spin in (
        title_size_spin,
        tick_size_spin,
        label_size_spin,
        legend_size_spin,
        line_width_spin,
    ):
        spin.config(command=_schedule_plot)

    ttk.Label(
        controls_frame,
        textvariable=status_var,
        wraplength=240,
        justify="left",
    ).pack(anchor="w", fill="x", pady=(10, 10))

    buttons_frame = ttk.Frame(controls_frame)
    buttons_frame.pack(fill="x", pady=(5, 0))

    ttk.Button(buttons_frame, text="Reset", command=_reset).pack(side="left", padx=(0, 6))
    ttk.Button(buttons_frame, text="Autoscale", command=_autofit).pack(side="left")

    _apply_lock_states()
    _update_smoothing_ui()
    _plot()


def _write_metadata_sheet(ws, parsed: ParsedDTA) -> None:
    ws["A1"] = "Campo"
    ws["B1"] = "Valor"
    ws["C1"] = "Unidad"

    for ref in ("A1", "B1", "C1"):
        ws[ref].font = Font(bold=True)

    for row_idx, (key, label) in enumerate(META_FIELDS, start=2):
        raw_value = parsed.meta_values.get(key, "")
        raw_unit = parsed.meta_units.get(key, "")
        ws.cell(row=row_idx, column=1, value=label)

        num = to_float(raw_value)
        ws.cell(row=row_idx, column=2, value=num if num is not None else raw_value)
        ws.cell(row=row_idx, column=3, value=raw_unit)

    ws.freeze_panes = "A2"


def _write_data_sheet(ws, parsed: ParsedDTA) -> None:
    col_idx = {name: idx for idx, name in enumerate(parsed.header)}
    selected_source_cols = [name for name, _label in DATA_EXPORT if name in col_idx]

    for col_num, source_name in enumerate(selected_source_cols, start=1):
        label = dict(DATA_EXPORT)[source_name]
        cell = ws.cell(row=1, column=col_num, value=label)
        cell.font = Font(bold=True)

    for col_num, source_name in enumerate(selected_source_cols, start=1):
        source_index = col_idx[source_name]
        unit_value = ""
        if parsed.units and source_index < len(parsed.units):
            unit_value = parsed.units[source_index]
            if unit_value == "#":
                unit_value = ""
        ws.cell(row=2, column=col_num, value=unit_value)

    for row_num, raw_parts in enumerate(parsed.rows, start=3):
        for col_num, source_name in enumerate(selected_source_cols, start=1):
            source_index = col_idx[source_name]
            raw_value = raw_parts[source_index] if source_index < len(raw_parts) else ""
            num = to_float(raw_value)
            ws.cell(row=row_num, column=col_num, value=num if num is not None else raw_value)

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


def export_file(source_path: Path, output_path: Path) -> None:
    parsed = parse_gamry_dta(source_path)
    wb = Workbook()
    ws_meta = wb.active
    ws_meta.title = "Metadata"
    _write_metadata_sheet(ws_meta, parsed)

    ws_data = wb.create_sheet("Data")
    _write_data_sheet(ws_data, parsed)

    for ws in (ws_meta, ws_data):
        _auto_format_sheet(ws)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def export_folder(input_dir: Path, output_dir: Path) -> list[Path]:
    exported: list[Path] = []
    for deg_file in find_deg_files(input_dir):
        out_path = output_dir / f"{deg_file.path.stem}.xlsx"
        export_file(deg_file.path, out_path)
        exported.append(out_path)
    return exported


def _show_deg_stub(title: str) -> None:
    messagebox.showinfo("Deg", f"{title} aun no esta implementado.")


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

    if "V vs t" in chosen:
        open_v_vs_t_window(input_dir)

    if "dV/dt" in chosen:
        open_dv_dt_window(input_dir)

    return exported_files
