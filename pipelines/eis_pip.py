
"""EIS (.DTA) -> Excel (.xlsx) exporter for Gamry Potentiostatic EIS.

What this version does:
  - Finds all .DTA files whose filename contains 'EISPOT'
  - Parses selected metadata fields
  - Parses the ZCURVE table
  - Exports ONE .xlsx per input file with two sheets:
        1) Metadata  -> Campo / Valor / Unidad
        2) Data      -> headers row, units row, then numeric data
  - Optionally creates plots depending on the GUI selection:
        * Nyquist plot        -> Zimag vs Zreal
        * Bode plot           -> Zmod vs Freq, and Zphz vs Freq
        * I vs pt             -> Idc vs Pt
        * T vs pt / T vs t    -> Temp vs Pt

This version is written to match the real structure of the uploaded Gamry files.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
import re

import math

from matplotlib.figure import Figure

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Labels to export (Spanish-friendly names)
# ---------------------------------------------------------------------------

META_FIELDS = [
    ("TITLE", "Técnica"),
    ("DATE", "Fecha"),
    ("TIME", "Hora"),
    ("VDC", "Vdc"),
    ("FREQINIT", "Frecuencia inicial"),
    ("FREQFINAL", "Frecuencia final"),
    ("PTSPERDEC", "Puntos por década"),
    ("VAC", "Amplitud"),
    ("AREA", "Área"),
]

DATA_MAP = {
    "Pt": "Pt",
    "Freq": "Frecuencia",
    "Zreal": "Zreal",
    "Zimag": "Zimag",
    "Zsig": "Zsig",
    "Zmod": "Zmod",
    "Zphz": "Zphz",
    "Idc": "Idc",
    "Vdc": "Vdc",
    "Temp": "Temperatura",
}


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
    """Gamry ZCURVE rows usually start with a leading tab."""
    if parts and parts[0] == "":
        return parts[1:]
    return parts


def _extract_parenthesized_unit(text: str) -> str:
    """Extract the last '(...)' group from a description line."""
    matches = re.findall(r"\(([^()]*)\)", text)
    if matches:
        return matches[-1].strip()
    return ""


def _extract_meta_unit(key: str, description: str) -> str:
    """Extract a clean metadata unit from the descriptive text."""
    unit = _extract_parenthesized_unit(description)
    if unit:
        return unit

    if key == "PTSPERDEC":
        return "puntos/década"

    if key in {"TITLE", "DATE", "TIME"}:
        return ""

    return ""


# ---------------------------------------------------------------------------
# Parsed container
# ---------------------------------------------------------------------------

@dataclass
class ParsedDTA:
    meta_values: dict[str, str]
    meta_units: dict[str, str]
    header: list[str]
    units: list[str]
    rows: list[list[str]]


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

def parse_gamry_dta(path: Path) -> ParsedDTA:
    """Parse one Gamry .DTA file containing a ZCURVE table."""
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
            if line.startswith("ZCURVE") and "TABLE" in line:
                table_started = True
                continue

            if not line.strip():
                continue

            parts = line.split("\t")
            if len(parts) >= 3 and parts[0].strip():
                key = parts[0].strip()
                value = parts[2].strip()
                description = parts[-1].strip() if len(parts) >= 4 else ""

                meta_values[key] = value
                meta_units[key] = _extract_meta_unit(key, description)

            continue

        if not line.strip():
            continue

        parts = _drop_leading_blank(line.rstrip("\r\n").split("\t"))
        if not parts:
            continue

        if header is None:
            if parts[0] == "Pt":
                header = parts
            continue

        if parts[0] == "#":
            units = parts
            continue

        if re.fullmatch(r"-?\d+", parts[0]):
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
# Data extraction for plotting
# ---------------------------------------------------------------------------

def _column_index(parsed: ParsedDTA, column_name: str) -> int | None:
    try:
        return parsed.header.index(column_name)
    except ValueError:
        return None


def _column_unit(parsed: ParsedDTA, column_name: str) -> str:
    idx = _column_index(parsed, column_name)
    if idx is None or not parsed.units or idx >= len(parsed.units):
        return ""
    unit = parsed.units[idx]
    if unit == "#":
        return ""
    return unit


def _numeric_series(parsed: ParsedDTA, column_name: str) -> list[float]:
    idx = _column_index(parsed, column_name)
    if idx is None:
        return []

    values: list[float] = []
    for row in parsed.rows:
        if idx >= len(row):
            continue
        num = to_float(row[idx])
        if num is None:
            continue
        values.append(num)
    return values


def _paired_series(
    parsed: ParsedDTA,
    x_name: str,
    y_name: str,
    *,
    require_positive_x: bool = False,
) -> tuple[list[float], list[float]]:
    x_idx = _column_index(parsed, x_name)
    y_idx = _column_index(parsed, y_name)
    if x_idx is None or y_idx is None:
        return [], []

    xs: list[float] = []
    ys: list[float] = []

    for row in parsed.rows:
        if x_idx >= len(row) or y_idx >= len(row):
            continue

        x_val = to_float(row[x_idx])
        y_val = to_float(row[y_idx])

        if x_val is None or y_val is None:
            continue
        if require_positive_x and x_val <= 0:
            continue

        xs.append(x_val)
        ys.append(y_val)

    return xs, ys


def _technique_name(parsed: ParsedDTA) -> str:
    return parsed.meta_values.get("TITLE", "").strip() or "Potentiostatic EIS"


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def export_to_xlsx(parsed: ParsedDTA, out_path: Path) -> None:
    """Create one .xlsx with Metadata + Data sheets."""
    wb = Workbook()
    wb.remove(wb.active)

    ws_meta = wb.create_sheet("Metadata")
    ws_data = wb.create_sheet("Data")

    # ---------------- Metadata sheet ---------------------------------------
    ws_meta["A1"] = "Campo"
    ws_meta["B1"] = "Valor"
    ws_meta["C1"] = "Unidad"
    for ref in ("A1", "B1", "C1"):
        ws_meta[ref].font = Font(bold=True)
    ws_meta.freeze_panes = "A2"

    numeric_meta_keys = {"VDC", "FREQINIT", "FREQFINAL", "PTSPERDEC", "VAC", "AREA"}

    for row_idx, (key, label) in enumerate(META_FIELDS, start=2):
        raw_value = parsed.meta_values.get(key, "")
        raw_unit = parsed.meta_units.get(key, "")

        ws_meta.cell(row=row_idx, column=1, value=label)

        if key in numeric_meta_keys:
            num = to_float(raw_value)
            ws_meta.cell(row=row_idx, column=2, value=num if num is not None else raw_value)
        else:
            ws_meta.cell(row=row_idx, column=2, value=raw_value)

        ws_meta.cell(row=row_idx, column=3, value=raw_unit)

    # ---------------- Data sheet -------------------------------------------
    col_idx = {name: idx for idx, name in enumerate(parsed.header)}
    selected_source_cols = [name for name in DATA_MAP if name in col_idx]
    selected_output_headers = [DATA_MAP[name] for name in selected_source_cols]

    # Row 1 = names
    for col_num, header_name in enumerate(selected_output_headers, start=1):
        cell = ws_data.cell(row=1, column=col_num, value=header_name)
        cell.font = Font(bold=True)

    # Row 2 = units
    for col_num, source_name in enumerate(selected_source_cols, start=1):
        unit_value = ""
        source_index = col_idx[source_name]

        if parsed.units and source_index < len(parsed.units):
            unit_value = parsed.units[source_index]
            if unit_value == "#":
                unit_value = ""

        ws_data.cell(row=2, column=col_num, value=unit_value)

    ws_data.freeze_panes = "A3"

    # Row 3 onward = numeric data
    for row_num, raw_parts in enumerate(parsed.rows, start=3):
        for col_num, source_name in enumerate(selected_source_cols, start=1):
            source_index = col_idx[source_name]
            raw_value = raw_parts[source_index] if source_index < len(raw_parts) else ""

            num = to_float(raw_value)
            ws_data.cell(row=row_num, column=col_num, value=num if num is not None else raw_value)

    # ---------------- Light formatting -------------------------------------
    for ws in (ws_meta, ws_data):
        for col_num in range(1, ws.max_column + 1):
            max_len = 0
            for row_num in range(1, min(ws.max_row, 50) + 1):
                value = ws.cell(row=row_num, column=col_num).value
                if value is None:
                    continue
                max_len = max(max_len, len(str(value)))

            ws.column_dimensions[get_column_letter(col_num)].width = min(max(10, max_len + 2), 45)

        for row_cells in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row_cells:
                cell.alignment = Alignment(vertical="top")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


# ---------------------------------------------------------------------------
# Plotting (interactive: build Figures and show them in a Tk window)
# ---------------------------------------------------------------------------

def _new_figure() -> Figure:
    # You can tweak size if you want bigger/smaller default tabs
    return Figure(figsize=(6.8, 4.8), dpi=100)

def fig_nyquist(parsed: ParsedDTA) -> Figure | None:
    x, y = _paired_series(parsed, "Zreal", "Zimag")
    if not x or not y:
        return None

    y_plot = [-v for v in y]  # Nyquist convention

    fig = _new_figure()
    ax = fig.add_subplot(111)

    ax.plot(x, y_plot, marker="o", linestyle="-", markerfacecolor="none")  # default styling

    # Square intervals: 1 unit in x == 1 unit in y
    ax.set_aspect("equal", adjustable="box")
    ax.margins(0.05)  # small padding

    x_unit = _column_unit(parsed, "Zreal")
    y_unit = _column_unit(parsed, "Zimag")
    ax.set_title(f"{_technique_name(parsed)} - Nyquist")
    ax.set_xlabel(f"Zreal ({x_unit})" if x_unit else "Zreal")
    ax.set_ylabel(f"-Zimag ({y_unit})" if y_unit else "-Zimag")
    ax.grid(True)

    fig.tight_layout()
    return fig

def figs_bode(parsed: ParsedDTA) -> list[tuple[str, Figure]]:
    """Two Bode figures: Zmod and Zphz."""
    out: list[tuple[str, Figure]] = []

    freq_unit = _column_unit(parsed, "Freq")
    zmod_unit = _column_unit(parsed, "Zmod")
    zphz_unit = _column_unit(parsed, "Zphz")

    x1, y1 = _paired_series(parsed, "Freq", "Zmod", require_positive_x=True)
    if x1 and y1:
        fig = _new_figure()
        ax = fig.add_subplot(111)
        ax.semilogx(x1, y1, marker="o")
        ax.set_title(f"{_technique_name(parsed)} - Bode (Zmod)")
        ax.set_xlabel(f"Frecuencia ({freq_unit})" if freq_unit else "Frecuencia")
        ax.set_ylabel(f"Zmod ({zmod_unit})" if zmod_unit else "Zmod")
        ax.grid(True, which="both")
        fig.tight_layout()
        out.append(("Bode (Zmod)", fig))

    x2, y2 = _paired_series(parsed, "Freq", "Zphz", require_positive_x=True)
    if x2 and y2:
        fig = _new_figure()
        ax = fig.add_subplot(111)
        ax.semilogx(x2, y2, marker="o")
        ax.set_title(f"{_technique_name(parsed)} - Bode (Zphz)")
        ax.set_xlabel(f"Frecuencia ({freq_unit})" if freq_unit else "Frecuencia")
        ax.set_ylabel(f"Zphz ({zphz_unit})" if zphz_unit else "Zphz")
        ax.grid(True, which="both")
        fig.tight_layout()
        out.append(("Bode (Zphz)", fig))

    return out

def fig_idc_vs_pt(parsed: ParsedDTA) -> Figure | None:
    x, y = _paired_series(parsed, "Pt", "Idc")
    if not x or not y:
        return None

    x_unit = _column_unit(parsed, "Pt")
    y_unit = _column_unit(parsed, "Idc")

    fig = _new_figure()
    ax = fig.add_subplot(111)
    ax.plot(x, y, marker="o")
    ax.set_title(f"{_technique_name(parsed)} - Idc vs Pt")
    ax.set_xlabel(f"Pt ({x_unit})" if x_unit else "Pt")
    ax.set_ylabel(f"Idc ({y_unit})" if y_unit else "Idc")
    ax.grid(True)
    fig.tight_layout()
    return fig

def fig_temp_vs_pt(parsed: ParsedDTA) -> Figure | None:
    x, y = _paired_series(parsed, "Pt", "Temp")
    if not x or not y:
        return None

    x_unit = _column_unit(parsed, "Pt")
    y_unit = _column_unit(parsed, "Temp")

    fig = _new_figure()
    ax = fig.add_subplot(111)
    ax.plot(x, y, marker="o")
    ax.set_title(f"{_technique_name(parsed)} - Temperatura vs Pt")
    ax.set_xlabel(f"Pt ({x_unit})" if x_unit else "Pt")
    ax.set_ylabel(f"Temperatura ({y_unit})" if y_unit else "Temperatura")
    ax.grid(True)
    fig.tight_layout()
    return fig

def build_figures(parsed: ParsedDTA, base_name: str, selected_options: Iterable[str] | None) -> list[tuple[str, Figure]]:
    """Create only the figures requested by the GUI."""
    if not selected_options:
        return []

    chosen = set(selected_options)
    figs: list[tuple[str, Figure]] = []

    if "Nyquist plot" in chosen:
        f = fig_nyquist(parsed)
        if f is not None:
            figs.append((f"{base_name} — Nyquist", f))

    if "Bode plot" in chosen:
        for label, f in figs_bode(parsed):
            figs.append((f"{base_name} — {label}", f))

    if "I vs pt" in chosen:
        f = fig_idc_vs_pt(parsed)
        if f is not None:
            figs.append((f"{base_name} — Idc vs Pt", f))

    if "T vs pt" in chosen or "T vs t" in chosen:
        f = fig_temp_vs_pt(parsed)
        if f is not None:
            figs.append((f"{base_name} — Temp vs Pt", f))

    # "Equivalent circuit fit" is listed in GUI but not implemented here yet.
    return figs

def show_figures_tk(figures: list[tuple[str, Figure]], window_title: str = "EIS plots") -> None:
    if not figures:
        return

    import tkinter as tk
    from tkinter import ttk, colorchooser
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

    root = tk._default_root
    created_root = False
    if root is None:
        root = tk.Tk()
        root.withdraw()
        created_root = True

    win = tk.Toplevel(root)
    win.title(window_title)
    win.geometry("1200x780")

    nb = ttk.Notebook(win)
    nb.pack(fill="both", expand=True)

    win._mpl_refs = []  # type: ignore[attr-defined]

    def _on_close():
        for _, fig in figures:
            fig.clear()
        win.destroy()
        if created_root:
            root.destroy()

    win.protocol("WM_DELETE_WINDOW", _on_close)

    marker_opts = ["o", ".", "x", "+", "s", "^", "v", "D", "None"]
    linestyle_opts = ["-", "--", "-.", ":", "None"]

    def _fmt(v: float) -> str:
        return f"{v:.6g}"
    
    from matplotlib.ticker import MaxNLocator

    def _snap_linear_limits(vmin: float, vmax: float, nbins: int = 6) -> tuple[float, float]:
        # Ensure order and non-degenerate span
        if vmax < vmin:
            vmin, vmax = vmax, vmin
        if vmin == vmax:
            pad = 1.0 if vmin == 0 else abs(vmin) * 0.1
            vmin -= pad
            vmax += pad

        loc = MaxNLocator(nbins=nbins)
        ticks = list(loc.tick_values(vmin, vmax))
        if not ticks:
            return vmin, vmax

        lo = max([t for t in ticks if t <= vmin], default=min(ticks))
        hi = min([t for t in ticks if t >= vmax], default=max(ticks))
        return float(lo), float(hi)

    def _add_tab(tab_title: str, fig: Figure) -> None:
        is_nyquist = ("nyquist" in tab_title.lower())

        tab = ttk.Frame(nb)
        nb.add(tab, text=tab_title[:28] + ("…" if len(tab_title) > 28 else ""))

        outer = ttk.Frame(tab)
        outer.pack(fill="both", expand=True)

        plot_frame = ttk.Frame(outer)
        plot_frame.pack(side="left", fill="both", expand=True)

        ctrl_frame = ttk.Frame(outer, padding=10)
        ctrl_frame.pack(side="right", fill="y")

        canvas = FigureCanvasTkAgg(fig, master=plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(canvas, plot_frame)
        toolbar.update()

        ax = fig.axes[0] if fig.axes else fig.add_subplot(111)
        line = ax.lines[0] if ax.lines else None

        if is_nyquist:
            ax.set_aspect("equal", adjustable="box")

        x0, x1 = ax.get_xlim()
        y0, y1 = ax.get_ylim()
        xmin_var = tk.StringVar(value=_fmt(x0))
        xmax_var = tk.StringVar(value=_fmt(x1))
        ymin_var = tk.StringVar(value=_fmt(y0))
        ymax_var = tk.StringVar(value=_fmt(y1))

        init_xlim = (x0, x1)
        init_ylim = (y0, y1)

        init_color = line.get_color() if line else ""
        init_marker = (line.get_marker() if line else "o") or "None"
        init_ls = (line.get_linestyle() if line else "-") or "None"
        init_lw = float(line.get_linewidth()) if line else 1.0
        init_ms = float(line.get_markersize()) if line else 4.0

        if init_marker in (None, "", "None"):
            init_marker = "None"
        if init_ls in (None, "", "None"):
            init_ls = "None"

        color_var = tk.StringVar(value=str(init_color))
        marker_var = tk.StringVar(value=str(init_marker))
        linestyle_var = tk.StringVar(value=str(init_ls))
        lw_var = tk.DoubleVar(value=init_lw)
        ms_var = tk.DoubleVar(value=init_ms)

        def _update_limit_entries():
            a0, a1 = ax.get_xlim()
            b0, b1 = ax.get_ylim()
            xmin_var.set(_fmt(a0))
            xmax_var.set(_fmt(a1))
            ymin_var.set(_fmt(b0))
            ymax_var.set(_fmt(b1))

        def _parse_float(s: str) -> float | None:
            s = s.strip()
            if not s:
                return None
            try:
                return float(s)
            except ValueError:
                return None

        def apply_axes():
            cur_x0, cur_x1 = ax.get_xlim()
            cur_y0, cur_y1 = ax.get_ylim()

            nx0 = _parse_float(xmin_var.get())
            nx1 = _parse_float(xmax_var.get())
            ny0 = _parse_float(ymin_var.get())
            ny1 = _parse_float(ymax_var.get())

            new_x0 = cur_x0 if nx0 is None else nx0
            new_x1 = cur_x1 if nx1 is None else nx1
            new_y0 = cur_y0 if ny0 is None else ny0
            new_y1 = cur_y1 if ny1 is None else ny1

            if ax.get_xscale() == "log" and (new_x0 <= 0 or new_x1 <= 0):
                _update_limit_entries()
                return
            if ax.get_yscale() == "log" and (new_y0 <= 0 or new_y1 <= 0):
                _update_limit_entries()
                return

            ax.set_xlim(new_x0, new_x1)
            ax.set_ylim(new_y0, new_y1)

            if is_nyquist:
                ax.set_aspect("equal", adjustable="box")

            canvas.draw_idle()
            _update_limit_entries()

        def reset_axes():
            ax.set_xlim(*init_xlim)
            ax.set_ylim(*init_ylim)
            if is_nyquist:
                ax.set_aspect("equal", adjustable="box")
            canvas.draw_idle()
            _update_limit_entries()

        def autoscale_axes():
            # Default: normal autoscale for log axes / non-Nyquist
            if ax.get_xscale() == "log" or ax.get_yscale() == "log" or not is_nyquist:
                ax.relim()
                ax.autoscale_view()
                if is_nyquist:
                    ax.set_aspect("equal", adjustable="box")
                canvas.draw_idle()
                _update_limit_entries()
                return

            # Nyquist + linear axes: Ymin = 0 and snap to "nice" ticks
            if line is None:
                ax.relim()
                ax.autoscale_view()
                x0, x1 = ax.get_xlim()
                y0, y1 = ax.get_ylim()
            else:
                xdata = list(line.get_xdata(orig=False))
                ydata = list(line.get_ydata(orig=False))
                # filter finite
                pts = [(x, y) for x, y in zip(xdata, ydata) if (x is not None and y is not None)]
                if not pts:
                    return
                xs = [float(x) for x, _ in pts]
                ys = [float(y) for _, y in pts]
                x0, x1 = min(xs), max(xs)
                y1 = max(ys)
                y0 = 0.0  # as requested

            # Snap to nearest "nice" axis marks
            x0s, x1s = _snap_linear_limits(x0, x1)
            y0s, y1s = _snap_linear_limits(0.0, y1)

            ax.set_xlim(x0s, x1s)
            ax.set_ylim(y0s, y1s)

            ax.set_aspect("equal", adjustable="box")  # preserve square units
            canvas.draw_idle()
            _update_limit_entries()

        def apply_style():
            if line is None:
                return

            c = color_var.get().strip()
            if c:
                line.set_color(c)

            ls = linestyle_var.get().strip()
            line.set_linestyle("" if ls == "None" else ls)

            mk = marker_var.get().strip()
            line.set_marker("" if mk == "None" else mk)

            # Keep markers hollow by default
            if line.get_marker() not in ("", None):
                line.set_markerfacecolor("none")
                line.set_markeredgecolor(line.get_color())

            try:
                line.set_linewidth(float(lw_var.get()))
            except Exception:
                pass
            try:
                line.set_markersize(float(ms_var.get()))
            except Exception:
                pass

            canvas.draw_idle()

        def reset_style():
            if line is None:
                return
            color_var.set(str(init_color))
            marker_var.set(str(init_marker))
            linestyle_var.set(str(init_ls))
            lw_var.set(init_lw)
            ms_var.set(init_ms)
            apply_style()

        def pick_color():
            if line is None:
                return
            chosen = colorchooser.askcolor(title="Choose line color")
            if chosen and chosen[1]:
                color_var.set(chosen[1])
                apply_style()

        # --- Controls UI ---
        axes_box = ttk.LabelFrame(ctrl_frame, text="Axes limits", padding=8)
        axes_box.pack(fill="x", pady=(0, 10))

        def _row(parent, r, label, var):
            ttk.Label(parent, text=label, width=5).grid(row=r, column=0, sticky="w", padx=(0, 6), pady=2)
            e = ttk.Entry(parent, textvariable=var, width=12)
            e.grid(row=r, column=1, sticky="w", pady=2)
            return e

        exmin = _row(axes_box, 0, "Xmin", xmin_var)
        exmax = _row(axes_box, 1, "Xmax", xmax_var)
        eymin = _row(axes_box, 2, "Ymin", ymin_var)
        eymax = _row(axes_box, 3, "Ymax", ymax_var)

        pending = {"id": None}

        def _schedule_apply_axes(_evt=None):
            if pending["id"] is not None:
                tab.after_cancel(pending["id"])
            pending["id"] = tab.after(300, apply_axes)

        for e in (exmin, exmax, eymin, eymax):
            e.bind("<Return>", apply_axes)
            e.bind("<FocusOut>", apply_axes)
            e.bind("<KeyRelease>", _schedule_apply_axes)

        btns_axes = ttk.Frame(axes_box)
        btns_axes.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(btns_axes, text="Apply", command=apply_axes).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(btns_axes, text="Autoscale", command=autoscale_axes).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(btns_axes, text="Reset", command=reset_axes).pack(side="left", expand=True, fill="x")

        style_box = ttk.LabelFrame(ctrl_frame, text="Style", padding=8)
        style_box.pack(fill="x", pady=(0, 10))

        ttk.Label(style_box, text="Color").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        color_entry = ttk.Entry(style_box, textvariable=color_var, width=12)
        color_entry.grid(row=0, column=1, sticky="w", pady=2)
        ttk.Button(style_box, text="Pick…", command=pick_color).grid(row=0, column=2, sticky="w", padx=(6, 0), pady=2)

        ttk.Label(style_box, text="Line").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=2)
        linestyle_cb = ttk.Combobox(style_box, textvariable=linestyle_var, values=linestyle_opts, width=9, state="readonly")
        linestyle_cb.grid(row=1, column=1, sticky="w", pady=2)

        ttk.Label(style_box, text="Marker").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=2)
        marker_cb = ttk.Combobox(style_box, textvariable=marker_var, values=marker_opts, width=9, state="readonly")
        marker_cb.grid(row=2, column=1, sticky="w", pady=2)

        ttk.Label(style_box, text="LW").grid(row=3, column=0, sticky="w", padx=(0, 6), pady=2)
        lw_spin = ttk.Spinbox(style_box, from_=0.0, to=10.0, increment=0.1, textvariable=lw_var, width=10)
        lw_spin.grid(row=3, column=1, sticky="w", pady=2)

        ttk.Label(style_box, text="MS").grid(row=4, column=0, sticky="w", padx=(0, 6), pady=2)
        ms_spin = ttk.Spinbox(style_box, from_=0.0, to=20.0, increment=0.5, textvariable=ms_var, width=10)
        ms_spin.grid(row=4, column=1, sticky="w", pady=2)

        # Comboboxes apply instantly on selection
        linestyle_cb.bind("<<ComboboxSelected>>", lambda e: apply_style())
        marker_cb.bind("<<ComboboxSelected>>", lambda e: apply_style())

        # Spinboxes: apply on arrow clicks + typing
        lw_spin.configure(command=apply_style)
        ms_spin.configure(command=apply_style)
        lw_spin.bind("<KeyRelease>", lambda e: apply_style())
        ms_spin.bind("<KeyRelease>", lambda e: apply_style())

        # Color entry: apply on Enter or leaving the field
        color_entry.bind("<Return>", lambda e: apply_style())
        color_entry.bind("<FocusOut>", lambda e: apply_style())

        btns_style = ttk.Frame(style_box)
        btns_style.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Button(btns_style, text="Apply", command=apply_style).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(btns_style, text="Reset", command=reset_style).pack(side="left", expand=True, fill="x")
        

        if line is None:
            for child in style_box.winfo_children():
                try:
                    child.configure(state="disabled")
                except Exception:
                    pass

        win._mpl_refs.append((canvas, toolbar, fig, ax, line))  # type: ignore[attr-defined]

    for tab_title, fig in figures:
        _add_tab(tab_title, fig)

    if created_root:
        win.mainloop()
        


# ---------------------------------------------------------------------------
# Folder export
# ---------------------------------------------------------------------------

def export_folder(
    input_dir: Path,
    output_dir: Path,
    selected_options: Iterable[str] | None = None,
) -> list[Path]:
    """Find all EISPOT .DTA files in input_dir and export them to output_dir.

    Returns the list of created .xlsx files.
    Plot files are also created as side effects when selected_options is provided.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    dta_files = sorted(
        [
            p for p in input_dir.iterdir()
            if p.is_file() and p.suffix.lower() == ".dta" and "EISPOT" in p.name
        ]
    )

    if not dta_files:
        return []

    exported_xlsx: list[Path] = []
    all_figs: list[tuple[str, Figure]] = []

    for dta_file in dta_files:
        parsed = parse_gamry_dta(dta_file)

        xlsx_path = output_dir / f"{dta_file.stem}.xlsx"
        export_to_xlsx(parsed, xlsx_path)
        exported_xlsx.append(xlsx_path)

        # NEW: create figures instead of exporting .svg
        all_figs.extend(build_figures(parsed, dta_file.stem, selected_options))
    
    # NEW: show one window with tabs for all figures
    show_figures_tk(all_figs, window_title="EIS plots")

    return exported_xlsx


def main() -> None:
    """Manual standalone test."""
    input_dir = Path(r"C:\\path\\to\\your\\input")
    output_dir = Path(r"C:\\path\\to\\your\\output")

    exported = export_folder(input_dir, output_dir, selected_options=["Nyquist plot", "Bode plot"])
    print(f"Exported {len(exported)} file(s)")
    for path in exported:
        print(" -", path)


if __name__ == "__main__":
    main()
