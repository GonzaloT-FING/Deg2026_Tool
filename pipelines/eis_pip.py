
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
from matplotlib.ticker import MaxNLocator

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

def _triplet_series(
    parsed: ParsedDTA,
    x_name: str,
    y_name: str,
    f_name: str,
    *,
    require_positive_f: bool = True,
) -> tuple[list[float], list[float], list[float]]:
    x_idx = _column_index(parsed, x_name)
    y_idx = _column_index(parsed, y_name)
    f_idx = _column_index(parsed, f_name)

    if x_idx is None or y_idx is None or f_idx is None:
        return [], [], []

    xs: list[float] = []
    ys: list[float] = []
    fs: list[float] = []

    for row in parsed.rows:
        if x_idx >= len(row) or y_idx >= len(row) or f_idx >= len(row):
            continue

        x_val = to_float(row[x_idx])
        y_val = to_float(row[y_idx])
        f_val = to_float(row[f_idx])

        if x_val is None or y_val is None or f_val is None:
            continue
        if require_positive_f and f_val <= 0:
            continue

        xs.append(x_val)
        ys.append(y_val)
        fs.append(f_val)

    return xs, ys, fs

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
    x, y, f = _triplet_series(parsed, "Zreal", "Zimag", "Freq")
    if not x or not y or not f:
        return None

    y_plot = [-v for v in y]

    fig = _new_figure()
    ax = fig.add_subplot(111)

    (line,) = ax.plot(x, y_plot, marker="o", linestyle="-", markerfacecolor="none")
    line._eis_freq = f  # attach frequency array (same index as points)

    ax.set_aspect("equal", adjustable="box")
    ax.margins(0.05)

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
        ax.semilogx(x1, y1, marker="o", linestyle="-", markerfacecolor="none")
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
        ax.semilogx(x2, y2, marker="o", linestyle="-", markerfacecolor="none")
        ax.set_title(f"{_technique_name(parsed)} - Bode (Zphz)")
        ax.set_xlabel(f"Frecuencia ({freq_unit})" if freq_unit else "Frecuencia")
        ax.set_ylabel(f"Zphz ({zphz_unit})" if zphz_unit else "Zphz")
        ax.grid(True, which="both")
        fig.tight_layout()
        out.append(("Bode (Zphz)", fig))

    return out


def fig_series_vs_pt(parsed: ParsedDTA) -> Figure | None:
    fig = _new_figure()
    axI = fig.add_subplot(111)

    lines: dict[str, object] = {}
    axes: dict[str, object] = {"I": axI}
    ylabels: dict[str, str] = {}

    # Create extra axes (always created; lines may or may not exist)
    axV = axI.twinx()
    axT = axI.twinx()
    axT.spines["right"].set_position(("outward", 60))

    # After:
    # axV = axI.twinx()
    # axT = axI.twinx()
    # axT.spines["right"].set_position(("outward", 60))

    axV.spines["left"].set_visible(False)
    axT.spines["left"].set_visible(False)

    # make sure their ticks are on the right (usually automatic, but explicit is fine)
    axV.yaxis.tick_right()
    axV.yaxis.set_label_position("right")
    axT.yaxis.tick_right()
    axT.yaxis.set_label_position("right")

    axes["V"] = axV
    axes["T"] = axT

    # X label
    x_unit = _column_unit(parsed, "Pt")
    axI.set_xlabel(f"Pt ({x_unit})" if x_unit else "Pt")

    # Grid only on base axis to avoid clutter
    axI.grid(True)
    axV.grid(False)
    axT.grid(False)

    def _add_series(key: str, col: str, label: str, ax):
        x, y, f = _triplet_series(parsed, "Pt", col, "Freq")
        if not x or not y or not f:
            return

        unit = _column_unit(parsed, col)
        lab = f"{label} ({unit})" if unit else label

        (ln,) = ax.plot(x, y, marker="o", linestyle="-", markerfacecolor="none", label=lab)
        ln._eis_freq = f
        lines[key] = ln
        ylabels[key] = lab

    _add_series("I", "Idc", "Idc", axI)
    _add_series("V", "Vdc", "Vdc", axV)
    _add_series("T", "Temp", "Temp", axT)

    if not lines:
        return None

    # Default visibility: show only Idc if present, else first available
    for k, ln in lines.items():
        ln.set_visible(False)

    base_key = "I" if "I" in lines else next(iter(lines.keys()))
    lines[base_key].set_visible(True)

    # Set ylabels for each axis (they can be colored later in UI)
    if "I" in lines: axI.set_ylabel(ylabels["I"])
    if "V" in lines: axV.set_ylabel(ylabels["V"])
    if "T" in lines: axT.set_ylabel(ylabels["T"])

    # Title
    axI.set_title(f"{_technique_name(parsed)} - Series vs Pt")

    # Save metadata for the UI
    fig._pt_series = True
    fig._pt_lines = lines
    fig._pt_axes = axes
    fig._pt_ylabels = ylabels
    fig._pt_base_key = base_key  # which axis drives the grid/tick alignment

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

    if "Series by Pt" in chosen:
        f = fig_series_vs_pt(parsed)
        if f is not None:
            figs.append((f"{base_name} — Series vs Pt", f))

    # "Equivalent circuit fit" is listed in GUI but not implemented here yet.
    return figs

def show_figures_tk(figures: list[tuple[str, Figure]], window_title: str = "EIS plots") -> None:

    tab_id_by_title: dict[str, str] = {}
    title_by_tab_id: dict[str, str] = {}

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

    # after nb = ttk.Notebook(win) ...
    topbar = ttk.Frame(win, padding=(8, 6))
    topbar.pack(fill="x")

    ttk.Label(topbar, text="Select plot:").pack(side="left")

    nyquist_sources: dict[str, dict] = {}
    bode_sources = {"zmod": {}, "zphz": {}}

    plot_names_var = tk.StringVar()
    plot_select = ttk.Combobox(topbar, textvariable=plot_names_var, state="readonly", width=55)
    plot_select.pack(side="left", padx=(6, 0), fill="x", expand=True)

    compose_btn = ttk.Button(topbar, text="Componer")
    compose_btn.pack(side="left", padx=(8, 0))

    def _sync_y_ticks(master_ax, other_axes, nbins: int = 6):
        # Ensure master has "nice" ticks
        loc = MaxNLocator(nbins=nbins)
        master_ax.yaxis.set_major_locator(loc)
        master_ax.figure.canvas.draw_idle()  # safe

        ticks = list(master_ax.get_yticks())
        if len(ticks) < 2:
            return

        y0, y1 = master_ax.get_ylim()
        span = (y1 - y0) if (y1 != y0) else 1.0
        fracs = [(t - y0) / span for t in ticks]

        for axk in other_axes:
            a0, a1 = axk.get_ylim()
            aspan = (a1 - a0) if (a1 != a0) else 1.0
            axk.set_yticks([a0 + f * aspan for f in fracs])

    def _goto_selected(_evt=None):
        wanted = plot_names_var.get()
        tab_id = tab_id_by_title.get(wanted)
        if tab_id is not None:
            nb.select(tab_id)

    plot_select.bind("<<ComboboxSelected>>", _goto_selected)

    def _sync_combo(_evt=None):
        tab_id = nb.select()
        title = title_by_tab_id.get(tab_id)
        if title:
            plot_names_var.set(title)

    nb.bind("<<NotebookTabChanged>>", _sync_combo)

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

        tab_id = str(tab)  # notebook tab identifier
        tab_id_by_title[tab_title] = tab_id
        title_by_tab_id[tab_id] = tab_title

        is_pt_series = bool(getattr(fig, "_pt_series", False))

        if is_pt_series:
            lines = getattr(fig, "_pt_lines", {})   # dict like {"I": line, "V": line, "T": line}
            axes  = getattr(fig, "_pt_axes",  {})
            # ✅ EVERYTHING that uses `lines` must be inside this block:
            # - Idc/Vdc/Temp checkboxes
            # - color buttons
            # - style notebook (linestyle/marker/lw/ms per series)
            # - tick sync / axis coloring helpers

        tlow = tab_title.lower()

        def _refresh_plot_dropdown():
            visible_titles = []
            for tab_id in nb.tabs():
                if nb.tab(tab_id, "state") != "hidden":
                    title = title_by_tab_id.get(tab_id)
                    if title:
                        visible_titles.append(title)

            plot_select["values"] = visible_titles

            # if current selection got hidden, jump to first visible
            if plot_names_var.get() not in visible_titles and visible_titles:
                plot_names_var.set(visible_titles[0])
                nb.select(tab_id_by_title[visible_titles[0]])

        # right after nb.add(tab, text=...)
        current = list(plot_select["values"])
        current.append(tab_title)
        plot_select["values"] = current

        # also set initial value once
        if not plot_names_var.get():
            plot_names_var.set(tab_title)

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

        # --- Pt Series controls (only on Series vs Pt figures) ---
        if getattr(fig, "_pt_series", False):

            def _apply_axis_color(ax, color: str, enabled: bool, side: str):
                # side: "left" or "right"
                c = color if enabled else "black"

                if side == "left":
                    ax.tick_params(axis="y", colors=c, labelleft=True, labelright=False)
                    ax.yaxis.label.set_color(c)
                    if "left" in ax.spines:
                        ax.spines["left"].set_color(c)
                    # do NOT touch right spine here
                else:
                    ax.tick_params(axis="y", colors=c, labelright=True, labelleft=False)
                    ax.yaxis.label.set_color(c)
                    if "right" in ax.spines:
                        ax.spines["right"].set_color(c)
                    # do NOT touch left spine here

            color_axes_var = tk.BooleanVar(value=True)

            def _refresh_axis_colors():
                if "I" in lines and "I" in axes:
                    _apply_axis_color(
                        axes["I"], lines["I"].get_color(),
                        enabled=(color_axes_var.get() and lines["I"].get_visible()),
                        side="left",
                    )
                if "V" in lines and "V" in axes:
                    _apply_axis_color(
                        axes["V"], lines["V"].get_color(),
                        enabled=(color_axes_var.get() and lines["V"].get_visible()),
                        side="right",
                    )
                if "T" in lines and "T" in axes:
                    _apply_axis_color(
                        axes["T"], lines["T"].get_color(),
                        enabled=(color_axes_var.get() and lines["T"].get_visible()),
                        side="right",
                    )
            
            def _pick_series_color(k: str):
                if k not in lines:
                    return
                chosen = colorchooser.askcolor(title=f"Choose color for {k}")
                if chosen and chosen[1]:
                    ln = lines[k]
                    ln.set_color(chosen[1])
                    # keep hollow marker consistent
                    try:
                        ln.set_markerfacecolor("none")
                        ln.set_markeredgecolor(chosen[1])
                    except Exception:
                        pass
                    _refresh_axis_colors()
                    _update_legend()
                    canvas.draw_idle()

            pt_box = ttk.LabelFrame(ctrl_frame, text="Series vs Pt", padding=8)
            pt_box.pack(fill="x", pady=(0, 10))

            lines = getattr(fig, "_pt_lines", {})
            ylabels = getattr(fig, "_pt_ylabels", {})

            style_nb = ttk.Notebook(pt_box)
            style_nb.pack(fill="x", pady=(8, 0))

            style_vars = {}

            ttk.Checkbutton(pt_box, text="Color y-axes", variable=color_axes_var,
                command=lambda: (_refresh_axis_colors(), canvas.draw_idle())).pack(anchor="w", pady=(6,0))
            
            row = ttk.Frame(pt_box)
            row.pack(fill="x", pady=(6,0))
            if "I" in lines: ttk.Button(row, text="Color Idc", command=lambda: _pick_series_color("I")).pack(side="left", padx=(0,6))
            if "V" in lines: ttk.Button(row, text="Color Vdc", command=lambda: _pick_series_color("V")).pack(side="left", padx=(0,6))
            if "T" in lines: ttk.Button(row, text="Color Temp", command=lambda: _pick_series_color("T")).pack(side="left")

            show_I = tk.BooleanVar(value=("I" in lines and lines["I"].get_visible()))
            show_V = tk.BooleanVar(value=("V" in lines and lines["V"].get_visible()))
            show_T = tk.BooleanVar(value=("T" in lines and lines["T"].get_visible()))

            axes = getattr(fig, "_pt_axes", {})          # {"I": axI, "V": axV, "T": axT}
            base_key = getattr(fig, "_pt_base_key", "I")
            base_ax = axes.get(base_key, ax)

            def _update_legend():
                handles = []
                labels = []
                for k, ln in lines.items():
                    if ln.get_visible():
                        handles.append(ln)
                        labels.append(ln.get_label())

                leg = base_ax.get_legend()
                if leg is not None:
                    leg.remove()

                if handles:
                    base_ax.legend(handles, labels, loc="best", fontsize=9)

            def _autoscale_visible_axes():
                for k, ln in lines.items():
                    axk = axes.get(k)
                    if axk is None:
                        continue
                    if ln.get_visible():
                        axk.relim()
                        axk.autoscale_view()

            def _apply_visibility():
                if "I" in lines: lines["I"].set_visible(bool(show_I.get()))
                if "V" in lines: lines["V"].set_visible(bool(show_V.get()))
                if "T" in lines: lines["T"].set_visible(bool(show_T.get()))

                _autoscale_visible_axes()

                # Choose master for tick/grid alignment: first visible in I->V->T order
                visible_keys = [k for k in ("I", "V", "T") if k in lines and lines[k].get_visible()]
                if visible_keys:
                    master_key = visible_keys[0]
                    master_ax = axes[master_key]

                    # grid only from master axis to avoid clutter
                    for axx in axes.values():
                        axx.grid(False)
                    master_ax.grid(True)

                    others = [axes[k] for k in visible_keys if k != master_key]
                    _sync_y_ticks(master_ax, others, nbins=6)

                _update_legend()
                _refresh_axis_colors()
                canvas.draw_idle()

            # ---- Per-series style notebook (Idc / Vdc / Temp) ----
            style_vars = {}

            def _apply_series_style(k: str):
                ln = lines[k]
                v = style_vars[k]

                c = v["color"].get().strip()
                if c:
                    ln.set_color(c)

                ls = v["ls"].get()
                mk = v["mk"].get()
                ln.set_linestyle("" if ls == "None" else ls)
                ln.set_marker("" if mk == "None" else mk)
                ln.set_linewidth(float(v["lw"].get()))
                ln.set_markersize(float(v["ms"].get()))

                # keep hollow markers
                if ln.get_marker() not in ("", None):
                    ln.set_markerfacecolor("none")
                    ln.set_markeredgecolor(ln.get_color())

                _update_legend()
                _refresh_axis_colors()
                canvas.draw_idle()

            for k, title in (("I", "Idc"), ("V", "Vdc"), ("T", "Temp")):
                if k not in lines:
                    continue

                f = ttk.Frame(style_nb, padding=6)
                style_nb.add(f, text=title)

                ln = lines[k]
                style_vars[k] = {
                    "color": tk.StringVar(value=str(ln.get_color())),
                    "ls": tk.StringVar(value=str(ln.get_linestyle() or "-") or "-"),
                    "mk": tk.StringVar(value=str(ln.get_marker() or "o") or "o"),
                    "lw": tk.DoubleVar(value=float(ln.get_linewidth())),
                    "ms": tk.DoubleVar(value=float(ln.get_markersize())),
                }

                ttk.Label(f, text="Color").grid(row=0, column=0, sticky="w")
                ce = ttk.Entry(f, textvariable=style_vars[k]["color"], width=10)
                ce.grid(row=0, column=1, sticky="w", padx=(6, 0))
                ttk.Button(f, text="Pick…", command=lambda kk=k: _pick_series_color(kk)).grid(row=0, column=2, sticky="w", padx=(6, 0))

                ttk.Label(f, text="Line").grid(row=1, column=0, sticky="w", pady=(6, 0))
                cb_ls = ttk.Combobox(f, textvariable=style_vars[k]["ls"], values=linestyle_opts, state="readonly", width=8)
                cb_ls.grid(row=1, column=1, sticky="w", padx=(6, 0), pady=(6, 0))

                ttk.Label(f, text="Marker").grid(row=2, column=0, sticky="w", pady=(6, 0))
                cb_mk = ttk.Combobox(f, textvariable=style_vars[k]["mk"], values=marker_opts, state="readonly", width=8)
                cb_mk.grid(row=2, column=1, sticky="w", padx=(6, 0), pady=(6, 0))

                ttk.Label(f, text="LW").grid(row=3, column=0, sticky="w", pady=(6, 0))
                sp_lw = ttk.Spinbox(f, from_=0.0, to=10.0, increment=0.1, textvariable=style_vars[k]["lw"], width=8)
                sp_lw.grid(row=3, column=1, sticky="w", padx=(6, 0), pady=(6, 0))

                ttk.Label(f, text="MS").grid(row=4, column=0, sticky="w", pady=(6, 0))
                sp_ms = ttk.Spinbox(f, from_=0.0, to=20.0, increment=0.5, textvariable=style_vars[k]["ms"], width=8)
                sp_ms.grid(row=4, column=1, sticky="w", padx=(6, 0), pady=(6, 0))

                # auto-apply
                ce.bind("<Return>", lambda e, kk=k: _apply_series_style(kk))
                ce.bind("<FocusOut>", lambda e, kk=k: _apply_series_style(kk))
                cb_ls.bind("<<ComboboxSelected>>", lambda e, kk=k: _apply_series_style(kk))
                cb_mk.bind("<<ComboboxSelected>>", lambda e, kk=k: _apply_series_style(kk))
                sp_lw.configure(command=lambda kk=k: _apply_series_style(kk))
                sp_ms.configure(command=lambda kk=k: _apply_series_style(kk))
                sp_lw.bind("<KeyRelease>", lambda e, kk=k: _apply_series_style(kk))
                sp_ms.bind("<KeyRelease>", lambda e, kk=k: _apply_series_style(kk))

                # auto-apply bindings
                ce.bind("<Return>", lambda e, kk=k: _apply_series_style(kk))
                ce.bind("<FocusOut>", lambda e, kk=k: _apply_series_style(kk))
                cb_ls.bind("<<ComboboxSelected>>", lambda e, kk=k: _apply_series_style(kk))
                cb_mk.bind("<<ComboboxSelected>>", lambda e, kk=k: _apply_series_style(kk))
                sp_lw.configure(command=lambda kk=k: _apply_series_style(kk))
                sp_ms.configure(command=lambda kk=k: _apply_series_style(kk))
                sp_lw.bind("<KeyRelease>", lambda e, kk=k: _apply_series_style(kk))
                sp_ms.bind("<KeyRelease>", lambda e, kk=k: _apply_series_style(kk))

            ttk.Checkbutton(pt_box, text="Idc", variable=show_I, command=_apply_visibility).pack(anchor="w")
            ttk.Checkbutton(pt_box, text="Vdc", variable=show_V, command=_apply_visibility).pack(anchor="w")
            ttk.Checkbutton(pt_box, text="Temp", variable=show_T, command=_apply_visibility).pack(anchor="w")

        if line is not None:
            if "nyquist" in tlow:
                nyquist_sources[tab_title] = {"line": line, "ax": ax, "fig": fig}
            elif "bode" in tlow and "zmod" in tlow:
                bode_sources["zmod"][tab_title] = {"line": line, "ax": ax, "fig": fig}
            elif "bode" in tlow and "zphz" in tlow:
                bode_sources["zphz"][tab_title] = {"line": line, "ax": ax, "fig": fig}

        cat = None
        if "vs pt" in tlow:
            if "idc" in tlow:
                cat = "pt_i"
            elif "vdc" in tlow:
                cat = "pt_v"
            elif "temp" in tlow:
                cat = "pt_t"
        

        if is_nyquist and line is not None:
            nyquist_sources[tab_title] = {
                "line": line,   # live reference (so we can copy current formatting)
                "ax": ax,
                "fig": fig,
            }

        init_title_text = ax.get_title()
        title_text_var = tk.StringVar(value=init_title_text)

        def apply_title_text():
            ax.set_title(title_text_var.get(), fontsize=float(title_fs_var.get()))
            try:
                fig.tight_layout()
            except Exception:
                pass
            canvas.draw_idle()

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

        def _safe_fg_for_bg(bg: str) -> str:
            # bg like "#RRGGBB"; returns black/white for readability
            if not (isinstance(bg, str) and bg.startswith("#") and len(bg) == 7):
                return "black"
            try:
                r = int(bg[1:3], 16)
                g = int(bg[3:5], 16)
                b = int(bg[5:7], 16)
            except ValueError:
                return "black"
            # perceived luminance
            lum = 0.2126*r + 0.7152*g + 0.0722*b
            return "black" if lum > 140 else "white"

        def _update_color_entry_bg():
            c = color_var.get().strip()
            if c.startswith("#") and len(c) == 7:
                try:
                    color_entry.configure(background=c, foreground=_safe_fg_for_bg(c))
                except Exception:
                    # ttk.Entry may ignore background on some themes; fallback below
                    pass

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
            _update_color_entry_bg()
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

        # ---------------- Fonts ----------------
        # Initial font sizes (grab from current artists)
        try:
            init_tick_fs = float(ax.get_xticklabels()[0].get_fontsize()) if ax.get_xticklabels() else 10.0
        except Exception:
            init_tick_fs = 10.0

        try:
            init_label_fs = float(ax.xaxis.label.get_fontsize() or 12.0)
        except Exception:
            init_label_fs = 12.0

        try:
            init_title_fs = float(ax.title.get_fontsize() or 14.0)
        except Exception:
            init_title_fs = 14.0

        tick_fs_var = tk.DoubleVar(value=init_tick_fs)
        label_fs_var = tk.DoubleVar(value=init_label_fs)
        title_fs_var = tk.DoubleVar(value=init_title_fs)

        def apply_fonts():
            # Apply to all axes in the figure (safe even if later you add multi-axes figs)
            try:
                tfs = float(tick_fs_var.get())
                lfs = float(label_fs_var.get())
                hfs = float(title_fs_var.get())
            except Exception:
                return

            for ax_ in fig.axes:
                ax_.tick_params(labelsize=tfs)
                ax_.xaxis.label.set_fontsize(lfs)
                ax_.yaxis.label.set_fontsize(lfs)
                ax_.title.set_fontsize(hfs)

            # Layout may need refresh when fonts change
            try:
                fig.tight_layout()
            except Exception:
                pass

            canvas.draw_idle()

        def reset_fonts():
            tick_fs_var.set(init_tick_fs)
            label_fs_var.set(init_label_fs)
            title_fs_var.set(init_title_fs)
            title_text_var.set(init_title_text)
            apply_fonts()
            apply_title_text()

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
        ttk.Button(btns_axes, text="Autoscale", command=autoscale_axes).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(btns_axes, text="Reset", command=reset_axes).pack(side="left", expand=True, fill="x")

        style_box = ttk.LabelFrame(ctrl_frame, text="Style", padding=8)
        style_box.pack(fill="x", pady=(0, 10))

        fonts_box = ttk.LabelFrame(ctrl_frame, text="Fonts", padding=8)
        fonts_box.pack(fill="x", pady=(0, 10))
        fonts_box.columnconfigure(1, weight=1)

        ttk.Label(fonts_box, text="Ticks").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        tick_spin = ttk.Spinbox(fonts_box, from_=6.0, to=30.0, increment=0.5,
                                textvariable=tick_fs_var, width=10)
        tick_spin.grid(row=0, column=1, sticky="w", pady=2)

        ttk.Label(fonts_box, text="Labels").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=2)
        label_spin = ttk.Spinbox(fonts_box, from_=6.0, to=40.0, increment=0.5,
                                textvariable=label_fs_var, width=10)
        label_spin.grid(row=1, column=1, sticky="w", pady=2)

        ttk.Label(fonts_box, text="Title").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=2)
        title_spin = ttk.Spinbox(fonts_box, from_=6.0, to=50.0, increment=0.5,
                                textvariable=title_fs_var, width=10)
        title_spin.grid(row=2, column=1, sticky="w", pady=2)

        ttk.Label(fonts_box, text="Text").grid(row=3, column=0, sticky="w", padx=(0, 6), pady=2)
        title_entry = ttk.Entry(fonts_box, textvariable=title_text_var, width=18)
        title_entry.grid(row=3, column=1, sticky="w", pady=2)
        title_entry.grid(row=3, column=1, sticky="ew", pady=2)

        btns_fonts = ttk.Frame(fonts_box)
        btns_fonts.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(btns_fonts, text="Reset", command=reset_fonts).pack(side="left", expand=True, fill="x")

        ttk.Label(style_box, text="Color").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)

        color_entry = tk.Entry(style_box, textvariable=color_var, width=12)
        color_entry.grid(row=0, column=1, sticky="w", pady=2)
        _update_color_entry_bg()
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

        # Apply on arrow clicks + typing
        tick_spin.configure(command=apply_fonts)
        label_spin.configure(command=apply_fonts)
        title_spin.configure(command=apply_fonts)

        tick_spin.bind("<KeyRelease>", lambda e: apply_fonts())
        label_spin.bind("<KeyRelease>", lambda e: apply_fonts())
        title_spin.bind("<KeyRelease>", lambda e: apply_fonts())

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
        ttk.Button(btns_style, text="Reset", command=reset_style).pack(side="left", expand=True, fill="x")

        pending_title = {"id": None}

        def _schedule_title(_evt=None):
            if pending_title["id"] is not None:
                tab.after_cancel(pending_title["id"])
            pending_title["id"] = tab.after(250, apply_title_text)

        # ---------------- Frequency tools (Nyquist only) ----------------
        freqs = getattr(line, "_eis_freq", None) if line is not None else None

        def _fmt_freq_hz(v: float) -> str:
            # nice readable frequency formatting
            av = abs(v)
            if av >= 1e6:
                return f"{v/1e6:.3g} MHz"
            if av >= 1e3:
                return f"{v/1e3:.3g} kHz"
            return f"{v:.3g} Hz"

        # Only enable this panel for Nyquist plots that actually have freq data
        if is_nyquist and line is not None and isinstance(freqs, list) and len(freqs) == len(line.get_xdata()):
            freq_box = ttk.LabelFrame(ctrl_frame, text="Frequency", padding=8)
            freq_box.pack(fill="x", pady=(0, 10))

            # --- Hover tooltip ---
            hover_var = tk.BooleanVar(value=True)

            hover_annot = ax.annotate(
                "",
                xy=(0, 0),
                xytext=(10, 10),
                textcoords="offset points",
                bbox=dict(boxstyle="round", fc="white", alpha=0.9),
                arrowprops=dict(arrowstyle="->", alpha=0.7),
            )
            hover_annot.set_visible(False)

            def _toggle_hover():
                if not hover_var.get():
                    hover_annot.set_visible(False)
                    canvas.draw_idle()

            ttk.Checkbutton(freq_box, text="Hover shows frequency", variable=hover_var, command=_toggle_hover)\
                .grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

            def _on_move(event):
                if not hover_var.get():
                    return
                if event.inaxes != ax or event.x is None or event.y is None:
                    if hover_annot.get_visible():
                        hover_annot.set_visible(False)
                        canvas.draw_idle()
                    return

                xdata = list(line.get_xdata(orig=False))
                ydata = list(line.get_ydata(orig=False))

                best_i = None
                best_d2 = 1e18
                # threshold ~12 px
                thresh2 = 12.0 * 12.0

                for i, (xv, yv) in enumerate(zip(xdata, ydata)):
                    xp, yp = ax.transData.transform((xv, yv))
                    d2 = (xp - event.x) ** 2 + (yp - event.y) ** 2
                    if d2 < best_d2:
                        best_d2 = d2
                        best_i = i

                if best_i is None or best_d2 > thresh2:
                    if hover_annot.get_visible():
                        hover_annot.set_visible(False)
                        canvas.draw_idle()
                    return

                fval = freqs[best_i]
                hover_annot.xy = (xdata[best_i], ydata[best_i])
                hover_annot.set_text(f"f = {_fmt_freq_hz(float(fval))}")
                if not hover_annot.get_visible():
                    hover_annot.set_visible(True)
                canvas.draw_idle()

            canvas.mpl_connect("motion_notify_event", _on_move)

            # --- Static labels for export ---
            ttk.Separator(freq_box).grid(row=1, column=0, columnspan=2, sticky="ew", pady=6)

            # Defaults: full range, labels OFF (N=0)
            fmin_default = min(freqs)
            fmax_default = max(freqs)

            freqmin_var = tk.StringVar(value=f"{fmin_default:.6g}")
            freqmax_var = tk.StringVar(value=f"{fmax_default:.6g}")
            nlabels_var = tk.IntVar(value=0)

            static_artists: list[object] = []

            def _clear_static():
                nonlocal static_artists
                for a in static_artists:
                    try:
                        a.remove()
                    except Exception:
                        pass
                static_artists = []

                # NEW: clear metadata so composite knows there are no labels
                try:
                    line._freq_label_idxs = []        # type: ignore[attr-defined]
                    line._freq_label_spec = None      # type: ignore[attr-defined]
                except Exception:
                    pass

                canvas.draw_idle()

            def _parse_float_or_none(s: str) -> float | None:
                s = s.strip()
                if not s:
                    return None
                try:
                    return float(s)
                except ValueError:
                    return None

            def _apply_static_labels():
                _clear_static()

                n = int(nlabels_var.get())
                if n <= 0:
                    return

                fmin_in = _parse_float_or_none(freqmin_var.get())
                fmax_in = _parse_float_or_none(freqmax_var.get())

                # if blank/invalid, use full range
                fmin_use = fmin_default if fmin_in is None else fmin_in
                fmax_use = fmax_default if fmax_in is None else fmax_in

                lo = min(fmin_use, fmax_use)
                hi = max(fmin_use, fmax_use)

                # candidate indices in the requested freq window
                candidates = [i for i, fv in enumerate(freqs) if lo <= float(fv) <= hi]

                if not candidates:
                    # fallback: whole curve
                    candidates = list(range(len(freqs)))

                # Choose indices:
                if n == 1:
                    target = fmin_use  # spec: nearest to Freqmin
                    idx = min(candidates, key=lambda i: abs(float(freqs[i]) - target))
                    idxs = [idx]
                    # NEW: store label selection on the line (for composite import)
                    try:
                        line._freq_label_idxs = list(idxs)  # type: ignore[attr-defined]
                        line._freq_label_spec = {           # type: ignore[attr-defined]
                            "freqmin": fmin_use,
                            "freqmax": fmax_use,
                            "n": int(n),
                        }
                    except Exception:
                        pass
                else:
                    # evenly spaced in point number *within* candidates
                    m = len(candidates)
                    if n > m:
                        n = m
                    if n == 1:
                        idxs = [candidates[0]]
                    else:
                        pos = [round(k * (m - 1) / (n - 1)) for k in range(n)]
                        idxs = sorted({candidates[p] for p in pos})

                xdata = list(line.get_xdata(orig=False))
                ydata = list(line.get_ydata(orig=False))
                fs = float(tick_fs_var.get()) if "tick_fs_var" in locals() else 10.0
                label_fs = max(7.0, fs * 0.9)

                for i in idxs:
                    txt = _fmt_freq_hz(float(freqs[i]))
                    a = ax.annotate(
                        txt,
                        xy=(xdata[i], ydata[i]),
                        xytext=(6, 6),
                        textcoords="offset points",
                        fontsize=label_fs,
                        bbox=dict(boxstyle="round,pad=0.15", fc="white", alpha=0.7),
                    )
                    static_artists.append(a)

                canvas.draw_idle()

            # UI widgets
            ttk.Label(freq_box, text="Freqmin").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=2)
            fmin_entry = ttk.Entry(freq_box, textvariable=freqmin_var, width=12)
            fmin_entry.grid(row=2, column=1, sticky="w", pady=2)

            ttk.Label(freq_box, text="Freqmax").grid(row=3, column=0, sticky="w", padx=(0, 6), pady=2)
            fmax_entry = ttk.Entry(freq_box, textvariable=freqmax_var, width=12)
            fmax_entry.grid(row=3, column=1, sticky="w", pady=2)

            ttk.Label(freq_box, text="N labels").grid(row=4, column=0, sticky="w", padx=(0, 6), pady=2)
            n_spin = ttk.Spinbox(freq_box, from_=0, to=50, increment=1, textvariable=nlabels_var, width=10)
            n_spin.grid(row=4, column=1, sticky="w", pady=2)

            btns_f = ttk.Frame(freq_box)
            btns_f.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(8, 0))
            ttk.Button(btns_f, text="Clear", command=lambda: (nlabels_var.set(0), _apply_static_labels()))\
                .pack(side="left", expand=True, fill="x")

            # Debounced auto-apply for static labels
            pending_freq = {"id": None}

            def _schedule_labels(_evt=None):
                if pending_freq["id"] is not None:
                    tab.after_cancel(pending_freq["id"])
                pending_freq["id"] = tab.after(300, _apply_static_labels)

            for e in (fmin_entry, fmax_entry):
                e.bind("<Return>", lambda ev: _apply_static_labels())
                e.bind("<FocusOut>", lambda ev: _apply_static_labels())
                e.bind("<KeyRelease>", _schedule_labels)

            n_spin.configure(command=_apply_static_labels)
            n_spin.bind("<KeyRelease>", lambda ev: _apply_static_labels())

        title_entry.bind("<Return>", lambda e: apply_title_text())
        title_entry.bind("<FocusOut>", lambda e: apply_title_text())
        title_entry.bind("<KeyRelease>", _schedule_title)
        

        if line is None:
            for child in style_box.winfo_children():
                try:
                    child.configure(state="disabled")
                except Exception:
                    pass

        win._mpl_refs.append((canvas, toolbar, fig, ax, line))  # type: ignore[attr-defined]

    def open_composer_nyquist():
        import tkinter as tk
        from tkinter import ttk
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

        if not nyquist_sources:
            return

        # If already open, bring to front
        existing = getattr(win, "_composer_win", None)
        if existing is not None and existing.winfo_exists():
            existing.lift()
            existing.focus_force()
            return

        def _src_label(key: str) -> str:
            # IMPORTANT: uses the *current* title of the source plot (user-editable)
            ax_src = nyquist_sources[key]["ax"]
            t = (ax_src.get_title() or "").strip()
            return t if t else key

        comp = tk.Toplevel(win)
        win._composer_win = comp  # type: ignore[attr-defined]
        comp.title("Composite (Nyquist)")
        comp.geometry("1250x780")

        outer = ttk.Frame(comp)
        outer.pack(fill="both", expand=True)

        # Left: plot
        plot_frame = ttk.Frame(outer)
        plot_frame.pack(side="left", fill="both", expand=True)

        figc = _new_figure()
        axc = figc.add_subplot(111)

        def _reset_composite_axes():
            axc.set_aspect("equal", adjustable="box")
            axc.grid(True)
            axc.set_title("Composite - Nyquist")
            axc.set_xlabel("Zreal")
            axc.set_ylabel("-Zimag")

        axc.set_aspect("equal", adjustable="box")
        axc.grid(True)
        axc.set_title("Composite - Nyquist")
        axc.set_xlabel("Zreal")
        axc.set_ylabel("-Zimag")

        canvas = FigureCanvasTkAgg(figc, master=plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(canvas, plot_frame)
        toolbar.update()

        # Right: controls
        ctrl = ttk.Frame(outer, padding=10)
        ctrl.pack(side="right", fill="y")

        # ---- Sources list (show current title + key) ----
        src_box = ttk.LabelFrame(ctrl, text="Nyquist sources", padding=8)
        src_box.pack(fill="x", pady=(0, 10))

        lb = tk.Listbox(src_box, selectmode="extended", height=12, exportselection=False)
        lb.pack(fill="x", expand=False)

        # mapping listbox index -> key
        idx_to_key: list[str] = []

        def _rebuild_listbox():
            nonlocal idx_to_key
            lb.delete(0, "end")
            idx_to_key = []

            # sort by label for nicer UX
            keys = list(nyquist_sources.keys())
            keys.sort(key=lambda k: _src_label(k).lower())

            for k in keys:
                lb.insert("end", f"{_src_label(k)}   [{k}]")
                idx_to_key.append(k)

        _rebuild_listbox()

        def _selected_keys() -> list[str]:
            return [idx_to_key[i] for i in lb.curselection()]

        # ---- Composite line storage: key -> line ----
        comp_lines: dict[str, object] = {}
        comp_label_artists: dict[str, list[object]] = {}  # NEW: per-curve annotations

        def _fmt_freq_hz(v: float) -> str:
            av = abs(v)
            if av >= 1e6:
                return f"{v/1e6:.3g} MHz"
            if av >= 1e3:
                return f"{v/1e3:.3g} kHz"
            return f"{v:.3g} Hz"
        
        def _clear_comp_labels(key: str):
            arts = comp_label_artists.pop(key, [])
            for a in arts:
                try:
                    a.remove()
                except Exception:
                    pass

        def _apply_comp_labels_from_source(key: str, dst_line):
            # remove any old ones
            _clear_comp_labels(key)

            src_line = nyquist_sources[key]["line"]
            idxs = getattr(src_line, "_freq_label_idxs", None)
            freqs = getattr(src_line, "_eis_freq", None)

            if not idxs or freqs is None:
                return
            # defensive: keep only valid indices
            npts = len(list(dst_line.get_xdata(orig=False)))
            idxs = [int(i) for i in idxs if 0 <= int(i) < npts]
            if not idxs:
                return

            xdata = list(dst_line.get_xdata(orig=False))
            ydata = list(dst_line.get_ydata(orig=False))

            arts: list[object] = []
            for i in idxs:
                txt = _fmt_freq_hz(float(freqs[i]))
                a = axc.annotate(
                    txt,
                    xy=(xdata[i], ydata[i]),
                    xytext=(6, 6),
                    textcoords="offset points",
                    fontsize=8.5,
                    bbox=dict(boxstyle="round,pad=0.15", fc="white", alpha=0.7),
                )
                arts.append(a)

            comp_label_artists[key] = arts

        legend_var = tk.BooleanVar(value=True)

        def _apply_legend():
            # Always remove existing legend first (prevents stacking / stale legends)
            leg = axc.get_legend()
            if leg is not None:
                leg.remove()

            if not legend_var.get():
                canvas.draw_idle()
                return

            handles, labels = axc.get_legend_handles_labels()
            # Keep only meaningful labels (ignore '_' internal ones)
            pairs = [(h, l) for h, l in zip(handles, labels) if l and not l.startswith("_")]
            if not pairs:
                canvas.draw_idle()
                return

            h2, l2 = zip(*pairs)
            axc.legend(h2, l2, loc="best", fontsize=9)
            canvas.draw_idle()

        def _copy_style(src_line, dst_line):
            # Copy *current* formatting from source line
            dst_line.set_color(src_line.get_color())
            dst_line.set_linestyle(src_line.get_linestyle())
            dst_line.set_marker(src_line.get_marker())
            dst_line.set_linewidth(src_line.get_linewidth())
            dst_line.set_markersize(src_line.get_markersize())

            # Also copy marker fill/edge if present
            try:
                dst_line.set_markerfacecolor(src_line.get_markerfacecolor())
            except Exception:
                pass
            try:
                dst_line.set_markeredgecolor(src_line.get_markeredgecolor())
            except Exception:
                pass
            try:
                dst_line.set_alpha(src_line.get_alpha())
            except Exception:
                pass

        def _fit_all():
            if not comp_lines:
                return

            xs: list[float] = []
            ys: list[float] = []

            for ln in comp_lines.values():
                x = [float(v) for v in ln.get_xdata(orig=False)]
                y = [float(v) for v in ln.get_ydata(orig=False)]
                xs.extend(x)
                ys.extend(y)

            if not xs or not ys:
                return

            x0, x1 = min(xs), max(xs)
            y0, y1 = min(ys), max(ys)

            dx = (x1 - x0) if x1 != x0 else (abs(x0) * 0.1 + 1.0)
            dy = (y1 - y0) if y1 != y0 else (abs(y0) * 0.1 + 1.0)

            pad_x = 0.05 * dx
            pad_y = 0.05 * dy

            axc.set_xlim(x0 - pad_x, x1 + pad_x)
            axc.set_ylim(y0 - pad_y, y1 + pad_y)
            axc.set_aspect("equal", adjustable="box")

            canvas.draw_idle()
            _sync_limit_entries()

        def add_selected():
            for key in _selected_keys():
                if key in comp_lines:
                    continue

                src_line = nyquist_sources[key]["line"]
                x = list(src_line.get_xdata(orig=False))
                y = list(src_line.get_ydata(orig=False))

                # label must match the *source plot title*
                (ln,) = axc.plot(x, y, label=_src_label(key))
                _copy_style(src_line, ln)

                # keep freqs too (future hover freq on composite)
                freqs = getattr(src_line, "_eis_freq", None)
                if freqs is not None:
                    ln._eis_freq = freqs  # type: ignore[attr-defined]

                comp_lines[key] = ln
                _apply_comp_labels_from_source(key, ln)   # NEW

            axc.set_aspect("equal", adjustable="box")
            _apply_legend()
            _fit_all()

        def remove_selected():
            removed = False
            for key in _selected_keys():
                ln = comp_lines.pop(key, None)
                if ln is not None:
                    try:
                        _clear_comp_labels(key)   # NEW
                        ln.remove()
                    except Exception:
                        pass
                    removed = True
            if removed:
                _apply_legend()
                _fit_all()

        def clear_all():
            # Clear our bookkeeping first
            comp_lines.clear()

            # If you implemented composed freq labels:
            try:
                comp_label_artists.clear()  # type: ignore[name-defined]
            except Exception:
                pass

            # Clear the axes in one shot (fast, removes lines + annotations + legend)
            axc.cla()
            _reset_composite_axes()

            # Redraw + sync UI fields
            canvas.draw_idle()
            _sync_limit_entries()

            # Optional: clear selection so user doesn't accidentally "Remove" nothing
            try:
                lb.selection_clear(0, "end")
            except Exception:
                pass

            # Legend should now be empty (but keep checkbox state consistent)
            _apply_legend()

        def refresh_formatting():
            # Refresh BOTH style and legend labels from the current state of source tabs
            for key, ln in comp_lines.items():
                src = nyquist_sources.get(key)
                if not src:
                    continue
                src_line = src["line"]
                _copy_style(src_line, ln)
                ln.set_label(_src_label(key))  # <-- ensures legend matches edited titles
                _apply_comp_labels_from_source(key, ln)

            _rebuild_listbox()
            _apply_legend()
            canvas.draw_idle()

        btns = ttk.Frame(src_box)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Add", command=add_selected).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(btns, text="Remove", command=remove_selected).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(btns, text="Clear", command=clear_all).pack(side="left", expand=True, fill="x")

        ttk.Button(src_box, text="Refresh formatting", command=refresh_formatting).pack(fill="x", pady=(8, 0))
        ttk.Checkbutton(ctrl, text="Legend", variable=legend_var, command=_apply_legend).pack(anchor="w", pady=(0, 10))

        # ---- Axis limits (independent in composite) ----
        lim_box = ttk.LabelFrame(ctrl, text="Axes limits", padding=8)
        lim_box.pack(fill="x", pady=(0, 10))

        def _fmt(v: float) -> str:
            return f"{v:.6g}"

        xmin_var = tk.StringVar()
        xmax_var = tk.StringVar()
        ymin_var = tk.StringVar()
        ymax_var = tk.StringVar()

        def _sync_limit_entries():
            x0, x1 = axc.get_xlim()
            y0, y1 = axc.get_ylim()
            xmin_var.set(_fmt(x0))
            xmax_var.set(_fmt(x1))
            ymin_var.set(_fmt(y0))
            ymax_var.set(_fmt(y1))

        _sync_limit_entries()

        def _parse_float(s: str) -> float | None:
            s = s.strip()
            if not s:
                return None
            try:
                return float(s)
            except ValueError:
                return None

        def apply_limits():
            cx0, cx1 = axc.get_xlim()
            cy0, cy1 = axc.get_ylim()

            nx0 = _parse_float(xmin_var.get())
            nx1 = _parse_float(xmax_var.get())
            ny0 = _parse_float(ymin_var.get())
            ny1 = _parse_float(ymax_var.get())

            axc.set_xlim(cx0 if nx0 is None else nx0, cx1 if nx1 is None else nx1)
            axc.set_ylim(cy0 if ny0 is None else ny0, cy1 if ny1 is None else ny1)
            axc.set_aspect("equal", adjustable="box")
            canvas.draw_idle()
            _sync_limit_entries()

        def _row(parent, r, label, var):
            ttk.Label(parent, text=label, width=5).grid(row=r, column=0, sticky="w", padx=(0, 6), pady=2)
            e = ttk.Entry(parent, textvariable=var, width=12)
            e.grid(row=r, column=1, sticky="w", pady=2)
            return e

        exmin = _row(lim_box, 0, "Xmin", xmin_var)
        exmax = _row(lim_box, 1, "Xmax", xmax_var)
        eymin = _row(lim_box, 2, "Ymin", ymin_var)
        eymax = _row(lim_box, 3, "Ymax", ymax_var)

        pending = {"id": None}

        def _schedule(_evt=None):
            if pending["id"] is not None:
                comp.after_cancel(pending["id"])
            pending["id"] = comp.after(300, apply_limits)

        for e in (exmin, exmax, eymin, eymax):
            e.bind("<Return>", lambda ev: apply_limits())
            e.bind("<FocusOut>", lambda ev: apply_limits())
            e.bind("<KeyRelease>", _schedule)

        b2 = ttk.Frame(lim_box)
        b2.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(b2, text="Fit all", command=_fit_all).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(b2, text="Refresh fields", command=_sync_limit_entries).pack(side="left", expand=True, fill="x")

        def _on_close():
            comp.destroy()
            try:
                delattr(win, "_composer_win")
            except Exception:
                pass

        comp.protocol("WM_DELETE_WINDOW", _on_close)

    def open_composer_bode(kind: str):  # kind in {"zmod","zphz"}
        import tkinter as tk
        from tkinter import ttk
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

        sources = bode_sources.get(kind, {})
        if not sources:
            return

        win_attr = f"_composer_win_bode_{kind}"
        existing = getattr(win, win_attr, None)
        if existing is not None and existing.winfo_exists():
            existing.lift()
            existing.focus_force()
            return

        def _src_label(key: str) -> str:
            ax_src = sources[key]["ax"]
            t = (ax_src.get_title() or "").strip()
            return t if t else key

        # Use first source labels for composite axes labels (keeps units)
        any_key = next(iter(sources.keys()))
        src_ax0 = sources[any_key]["ax"]
        default_xlabel = src_ax0.get_xlabel() or "Frecuencia"
        default_ylabel = src_ax0.get_ylabel() or ("Zmod" if kind == "zmod" else "Zphz")
        default_title = f"Composite - Bode ({'Zmod' if kind == 'zmod' else 'Zphz'})"

        comp = tk.Toplevel(win)
        setattr(win, win_attr, comp)
        comp.title(f"Composite (Bode - {'Zmod' if kind == 'zmod' else 'Zphz'})")
        comp.geometry("1250x780")

        outer = ttk.Frame(comp)
        outer.pack(fill="both", expand=True)

        plot_frame = ttk.Frame(outer)
        plot_frame.pack(side="left", fill="both", expand=True)

        figc = _new_figure()
        axc = figc.add_subplot(111)

        def _reset_axes():
            axc.set_xscale("log")
            axc.grid(True, which="both")
            axc.set_title(default_title)
            axc.set_xlabel(default_xlabel)
            axc.set_ylabel(default_ylabel)

        _reset_axes()

        canvas = FigureCanvasTkAgg(figc, master=plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(canvas, plot_frame)
        toolbar.update()

        ctrl = ttk.Frame(outer, padding=10)
        ctrl.pack(side="right", fill="y")

        # -------- source list --------
        src_box = ttk.LabelFrame(ctrl, text="Bode sources", padding=8)
        src_box.pack(fill="x", pady=(0, 10))

        lb = tk.Listbox(src_box, selectmode="extended", height=12, exportselection=False)
        lb.pack(fill="x", expand=False)

        idx_to_key: list[str] = []

        def _rebuild_listbox():
            nonlocal idx_to_key
            lb.delete(0, "end")
            idx_to_key = []
            keys = list(sources.keys())
            keys.sort(key=lambda k: _src_label(k).lower())
            for k in keys:
                lb.insert("end", f"{_src_label(k)}   [{k}]")
                idx_to_key.append(k)

        _rebuild_listbox()

        def _selected_keys() -> list[str]:
            return [idx_to_key[i] for i in lb.curselection()]

        # -------- composite lines --------
        comp_lines: dict[str, object] = {}
        legend_var = tk.BooleanVar(value=True)

        def _apply_legend():
            leg = axc.get_legend()
            if leg is not None:
                leg.remove()

            if not legend_var.get():
                canvas.draw_idle()
                return

            handles, labels = axc.get_legend_handles_labels()
            pairs = [(h, l) for h, l in zip(handles, labels) if l and not l.startswith("_")]
            if not pairs:
                canvas.draw_idle()
                return

            h2, l2 = zip(*pairs)
            axc.legend(h2, l2, loc="best", fontsize=9)
            canvas.draw_idle()

        def _copy_style(src_line, dst_line):
            dst_line.set_color(src_line.get_color())
            dst_line.set_linestyle(src_line.get_linestyle())
            dst_line.set_marker(src_line.get_marker())
            dst_line.set_linewidth(src_line.get_linewidth())
            dst_line.set_markersize(src_line.get_markersize())
            try:
                dst_line.set_markerfacecolor(src_line.get_markerfacecolor())
            except Exception:
                pass
            try:
                dst_line.set_markeredgecolor(src_line.get_markeredgecolor())
            except Exception:
                pass
            try:
                dst_line.set_alpha(src_line.get_alpha())
            except Exception:
                pass

        def _fit_all():
            if not comp_lines:
                return

            xs: list[float] = []
            ys: list[float] = []
            for ln in comp_lines.values():
                x = [float(v) for v in ln.get_xdata(orig=False)]
                y = [float(v) for v in ln.get_ydata(orig=False)]
                for xv, yv in zip(x, y):
                    if xv > 0:
                        xs.append(xv)
                        ys.append(yv)

            if not xs or not ys:
                return

            x0, x1 = min(xs), max(xs)
            y0, y1 = min(ys), max(ys)

            # log-friendly padding on x, linear padding on y
            axc.set_xlim(x0 / 1.2, x1 * 1.2)

            dy = (y1 - y0) if y1 != y0 else (abs(y0) * 0.1 + 1.0)
            pad_y = 0.05 * dy
            axc.set_ylim(y0 - pad_y, y1 + pad_y)

            canvas.draw_idle()
            _sync_limit_entries()

        def add_selected():
            for key in _selected_keys():
                if key in comp_lines:
                    continue
                src_line = sources[key]["line"]
                x = list(src_line.get_xdata(orig=False))
                y = list(src_line.get_ydata(orig=False))
                (ln,) = axc.plot(x, y, label=_src_label(key))
                _copy_style(src_line, ln)
                comp_lines[key] = ln

            _apply_legend()
            _fit_all()

        def remove_selected():
            removed = False
            for key in _selected_keys():
                ln = comp_lines.pop(key, None)
                if ln is not None:
                    try:
                        ln.remove()
                    except Exception:
                        pass
                    removed = True
            if removed:
                _apply_legend()
                _fit_all()

        def clear_all():
            comp_lines.clear()
            axc.cla()
            _reset_axes()
            canvas.draw_idle()
            _sync_limit_entries()
            _apply_legend()

        def refresh_formatting():
            # refresh BOTH style and legend labels (titles may have been edited)
            for key, ln in comp_lines.items():
                src_line = sources[key]["line"]
                _copy_style(src_line, ln)
                ln.set_label(_src_label(key))
            _rebuild_listbox()
            _apply_legend()
            canvas.draw_idle()

        btns = ttk.Frame(src_box)
        btns.pack(fill="x", pady=(8, 0))
        ttk.Button(btns, text="Add", command=add_selected).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(btns, text="Remove", command=remove_selected).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(btns, text="Clear", command=clear_all).pack(side="left", expand=True, fill="x")

        ttk.Button(src_box, text="Refresh formatting", command=refresh_formatting).pack(fill="x", pady=(8, 0))
        ttk.Checkbutton(ctrl, text="Legend", variable=legend_var, command=_apply_legend).pack(anchor="w", pady=(0, 10))

        # -------- axis limits UI --------
        lim_box = ttk.LabelFrame(ctrl, text="Axes limits", padding=8)
        lim_box.pack(fill="x", pady=(0, 10))

        def _fmt(v: float) -> str:
            return f"{v:.6g}"

        xmin_var = tk.StringVar()
        xmax_var = tk.StringVar()
        ymin_var = tk.StringVar()
        ymax_var = tk.StringVar()

        def _sync_limit_entries():
            x0, x1 = axc.get_xlim()
            y0, y1 = axc.get_ylim()
            xmin_var.set(_fmt(x0))
            xmax_var.set(_fmt(x1))
            ymin_var.set(_fmt(y0))
            ymax_var.set(_fmt(y1))

        _sync_limit_entries()

        def _parse_float(s: str) -> float | None:
            s = s.strip()
            if not s:
                return None
            try:
                return float(s)
            except ValueError:
                return None

        def apply_limits():
            cx0, cx1 = axc.get_xlim()
            cy0, cy1 = axc.get_ylim()

            nx0 = _parse_float(xmin_var.get())
            nx1 = _parse_float(xmax_var.get())
            ny0 = _parse_float(ymin_var.get())
            ny1 = _parse_float(ymax_var.get())

            new_x0 = cx0 if nx0 is None else nx0
            new_x1 = cx1 if nx1 is None else nx1

            # log-x must be > 0
            if new_x0 <= 0 or new_x1 <= 0:
                _sync_limit_entries()
                return

            axc.set_xlim(new_x0, new_x1)
            axc.set_ylim(cy0 if ny0 is None else ny0, cy1 if ny1 is None else ny1)

            canvas.draw_idle()
            _sync_limit_entries()

        def _row(parent, r, label, var):
            ttk.Label(parent, text=label, width=5).grid(row=r, column=0, sticky="w", padx=(0, 6), pady=2)
            e = ttk.Entry(parent, textvariable=var, width=12)
            e.grid(row=r, column=1, sticky="w", pady=2)
            return e

        exmin = _row(lim_box, 0, "Xmin", xmin_var)
        exmax = _row(lim_box, 1, "Xmax", xmax_var)
        eymin = _row(lim_box, 2, "Ymin", ymin_var)
        eymax = _row(lim_box, 3, "Ymax", ymax_var)

        pending = {"id": None}

        def _schedule(_evt=None):
            if pending["id"] is not None:
                comp.after_cancel(pending["id"])
            pending["id"] = comp.after(300, apply_limits)

        for e in (exmin, exmax, eymin, eymax):
            e.bind("<Return>", lambda ev: apply_limits())
            e.bind("<FocusOut>", lambda ev: apply_limits())
            e.bind("<KeyRelease>", _schedule)

        b2 = ttk.Frame(lim_box)
        b2.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(b2, text="Fit all", command=_fit_all).pack(side="left", expand=True, fill="x", padx=(0, 6))
        ttk.Button(b2, text="Refresh fields", command=_sync_limit_entries).pack(side="left", expand=True, fill="x")

        def _on_close():
            comp.destroy()
            try:
                delattr(win, win_attr)
            except Exception:
                pass

        comp.protocol("WM_DELETE_WINDOW", _on_close)

    def open_composer_for_current():
        key = plot_names_var.get().lower()

        if "nyquist" in key:
            open_composer_nyquist()
            return

        if "bode" in key and "zmod" in key:
            open_composer_bode(kind="zmod")
            return

        if "bode" in key and "zphz" in key:
            open_composer_bode(kind="zphz")
            return

        # optional: small message
        import tkinter.messagebox as mb
        mb.showinfo("Componer", "Composer is available for Nyquist and Bode plots only.")

    compose_btn.configure(command=open_composer_for_current)

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
