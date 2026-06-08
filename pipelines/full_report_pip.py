from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Callable
import colorsys
import math
import re

from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.figure import Figure
from matplotlib.ticker import LinearLocator, MaxNLocator, StrMethodFormatter

from i18n import normalize_language, translate
from plot_defaults import PlotFontDefaults, apply_plot_font_defaults, make_legend_draggable, resolve_plot_font_defaults
from pipelines import activ_pip as activ
from pipelines import cic_vol_pip as cv
from pipelines import deg_pip as deg
from pipelines import eis_pip as eis
from pipelines import pol_cur_pip as pc


ProgressCallback = Callable[[str, int | None, int | None], None]
SummaryIndicatorRow = tuple[str, str, str, object, str]


def _safe_filename_part(text: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", text.strip())
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned.strip("._") or "Full_Report"


def _emit(progress_callback: ProgressCallback | None, message: str, done: int | None = None, total: int | None = None) -> None:
    if progress_callback is not None:
        progress_callback(message, done, total)


def _optional_float(value: object) -> float | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(text.replace(",", "."))
    except ValueError:
        return None


def _format_sig(value: object, digits: int = 6) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if not math.isfinite(value):
            return ""
        if value == 0:
            return "0"
        abs_value = abs(value)
        if 1e-6 <= abs_value < 1e6:
            return f"{value:.{digits}f}".rstrip("0").rstrip(".")
        return f"{value:.{digits}g}"
    return str(value)


def _hls_to_hex(hue: float, lightness: float = 0.46, saturation: float = 0.75) -> str:
    r, g, b = colorsys.hls_to_rgb(hue % 1.0, lightness, saturation)
    return f"#{int(round(r * 255)):02x}{int(round(g * 255)):02x}{int(round(b * 255)):02x}"


def _color_for_index(index: int, total: int) -> str:
    palette = [
        "#1f77b4",
        "#d62728",
        "#2ca02c",
        "#9467bd",
        "#ff7f0e",
        "#17becf",
        "#8c564b",
        "#e377c2",
        "#7f7f7f",
        "#bcbd22",
    ]
    if index < len(palette):
        return palette[index]
    return _hls_to_hex(index / max(1, total))


def _stage_label(stage_number: int | None, fallback: str, language: str) -> str:
    if stage_number is None:
        return fallback
    return f"{translate('stage', language)} {stage_number}"


def _new_plot_figure(figsize: tuple[float, float] = (9.5, 6.2), dpi: int = 150) -> Figure:
    fig = Figure(figsize=figsize, dpi=dpi)
    FigureCanvasAgg(fig)
    return fig


def _save_figure(pdf: PdfPages, fig: Figure) -> None:
    if fig.canvas is None:
        FigureCanvasAgg(fig)
    pdf.savefig(fig, bbox_inches="tight")


def _add_report_table(
    ax,
    title: str,
    rows: list[tuple[object, object, object]],
    bbox: list[float],
    language: str,
    *,
    font_size: float = 8.4,
) -> None:
    table_width = min(float(bbox[2]), 0.82)
    bbox = [(1.0 - table_width) / 2.0, bbox[1], table_width, bbox[3]]
    ax.text(
        bbox[0],
        bbox[1] + bbox[3] + 0.025,
        title,
        fontsize=13,
        fontweight="bold",
        ha="left",
        va="bottom",
        transform=ax.transAxes,
    )
    cell_rows = [[_format_sig(name), _format_sig(value), _format_sig(unit)] for name, value, unit in rows]
    table = ax.table(
        cellText=cell_rows,
        colLabels=[
            translate("name", language),
            translate("value", language),
            translate("unit", language),
        ],
        cellLoc="left",
        colLoc="left",
        bbox=bbox,
    )
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)
    for (row_idx, _col_idx), cell in table.get_celld().items():
        cell.set_edgecolor("#b8c0c8")
        cell.set_linewidth(0.5)
        if row_idx == 0:
            cell.set_facecolor("#edf2f7")
            cell.set_text_props(weight="bold")
        else:
            cell.set_facecolor("white")


def _add_summary_indicator_table(
    ax,
    rows: list[SummaryIndicatorRow],
    bbox: list[float],
    language: str,
    *,
    font_size: float = 6.8,
) -> None:
    table = ax.table(
        cellText=[
            [_format_sig(pipeline), _format_sig(item), _format_sig(indicator), _format_sig(value), _format_sig(unit)]
            for pipeline, item, indicator, value, unit in rows
        ],
        colLabels=[
            translate("pipeline", language),
            translate("item", language),
            translate("indicator", language),
            translate("value", language),
            translate("unit", language),
        ],
        cellLoc="left",
        colLoc="left",
        colWidths=[0.13, 0.22, 0.37, 0.15, 0.13],
        bbox=bbox,
    )
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)
    for (row_idx, _col_idx), cell in table.get_celld().items():
        cell.set_edgecolor("#b8c0c8")
        cell.set_linewidth(0.45)
        if row_idx == 0:
            cell.set_facecolor("#edf2f7")
            cell.set_text_props(weight="bold")
        else:
            cell.set_facecolor("white")


def _table_page(
    title: str,
    subtitle: str,
    table_specs: list[tuple[str, list[tuple[object, object, object]], list[float]]],
    language: str,
) -> Figure:
    fig = _new_plot_figure(figsize=(8.5, 11.0), dpi=150)
    ax = fig.add_subplot(111)
    ax.axis("off")
    fig.text(0.05, 0.965, title, fontsize=18, fontweight="bold", ha="left", va="top")
    if subtitle:
        fig.text(0.05, 0.935, subtitle, fontsize=10, ha="left", va="top", color="#4a5568")
    for table_title, rows, bbox in table_specs:
        if rows:
            _add_report_table(ax, table_title, rows, bbox, language)
    return fig


def _title_page(input_dir: Path, output_dir: Path, counts: dict[str, int], language: str) -> Figure:
    fig = _new_plot_figure(figsize=(8.5, 11.0), dpi=150)
    ax = fig.add_subplot(111)
    ax.axis("off")
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    fig.text(
        0.08,
        0.88,
        translate("full_report_title", language),
        fontsize=26,
        fontweight="bold",
        ha="left",
        va="top",
    )
    fig.text(0.08, 0.835, now, fontsize=11, color="#4a5568", ha="left", va="top")
    fig.text(0.08, 0.78, f"{translate('input_folder', language)}:", fontsize=12, fontweight="bold", ha="left")
    fig.text(0.08, 0.755, str(input_dir), fontsize=9.5, color="#2d3748", ha="left")
    fig.text(0.08, 0.715, f"{translate('output_folder', language)}:", fontsize=12, fontweight="bold", ha="left")
    fig.text(0.08, 0.690, str(output_dir), fontsize=9.5, color="#2d3748", ha="left")

    rows = [
        ("Activacion", counts.get("activation", 0), translate("curve", language)),
        ("PC", counts.get("pc", 0), translate("curve", language)),
        ("EIS", counts.get("eis", 0), translate("nyquist_plot", language)),
        ("EIS Pre", counts.get("pre_stab", 0), translate("pre_stabilization", language)),
        ("CV", counts.get("cv", 0), "CV"),
        ("Deg", counts.get("deg", 0), translate("deg_stage_count", language)),
    ]
    _add_report_table(
        ax,
        translate("full_report_contents", language),
        [(name, value, unit) for name, value, unit in rows],
        [0.08, 0.39, 0.84, 0.24],
        language,
        font_size=9.0,
    )
    return fig


def _section_page(title: str, subtitle: str, language: str) -> Figure:
    fig = _new_plot_figure(figsize=(8.5, 11.0), dpi=150)
    ax = fig.add_subplot(111)
    ax.axis("off")
    fig.text(0.08, 0.62, title, fontsize=25, fontweight="bold", ha="left", va="center")
    if subtitle:
        fig.text(0.08, 0.57, subtitle, fontsize=12, color="#4a5568", ha="left", va="center")
    fig.text(0.08, 0.08, translate("full_report", language), fontsize=9, color="#718096", ha="left")
    return fig


def _index_page(section_titles: list[str], language: str) -> Figure:
    fig = _new_plot_figure(figsize=(8.5, 11.0), dpi=150)
    ax = fig.add_subplot(111)
    ax.axis("off")
    fig.text(0.08, 0.88, translate("full_report_index", language), fontsize=25, fontweight="bold", ha="left", va="top")
    fig.text(0.08, 0.84, translate("full_report_index_subtitle", language), fontsize=12, color="#4a5568", ha="left", va="top")
    y = 0.76
    for idx, title in enumerate(section_titles, start=1):
        fig.text(0.12, y, f"{idx}. {title}", fontsize=12, ha="left", va="center", color="#2d3748")
        y -= 0.052
        if y < 0.12:
            break
    fig.text(0.08, 0.08, translate("full_report", language), fontsize=9, color="#718096", ha="left")
    return fig


def _add_pdf_outline(output_path: Path, section_entries: list[tuple[str, int]]) -> None:
    if not section_entries:
        return
    try:
        from pypdf import PdfReader, PdfWriter
    except Exception:
        try:
            from PyPDF2 import PdfReader, PdfWriter
        except Exception:
            return

    try:
        reader = PdfReader(str(output_path))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        for title, page_index in section_entries:
            if 0 <= page_index < len(reader.pages):
                if hasattr(writer, "add_outline_item"):
                    writer.add_outline_item(title, page_index)
                elif hasattr(writer, "addBookmark"):
                    writer.addBookmark(title, page_index)
        temp_path = output_path.with_name(f"{output_path.stem}.tmp{output_path.suffix}")
        with temp_path.open("wb") as handle:
            writer.write(handle)
        temp_path.replace(output_path)
    except Exception:
        return


def _bookmark(section_entries: list[tuple[str, int]], title: str, page_index: int) -> None:
    section_entries.append((title, page_index))


def _pc_current_axis_mode(bundles: list[pc.CurveBundle]) -> bool:
    if not bundles:
        return False
    for bundle in bundles:
        area_cm2 = pc._bundle_area_cm2(bundle)
        if area_cm2 is None or area_cm2 <= 0:
            return False
    return True


def _pc_bundle_use_density(bundle: pc.CurveBundle) -> bool:
    area_cm2 = pc._bundle_area_cm2(bundle)
    return area_cm2 is not None and area_cm2 > 0


def _draw_pc_ascending_summary(
    bundles: list[pc.CurveBundle],
    font_defaults: PlotFontDefaults,
    language: str,
) -> Figure | None:
    if not bundles:
        return None

    use_current_density = _pc_current_axis_mode(bundles)
    fig = _new_plot_figure(figsize=(10.5, 6.6), dpi=150)
    ax = fig.add_subplot(111)
    all_x: list[float] = []
    all_y: list[float] = []

    for index, bundle in enumerate(bundles):
        if not bundle.asc_files:
            continue
        asc_rows = pc.concatenate_curve_data(bundle.asc_files)
        last_rows = pc.find_last_point_of_each_step(asc_rows, pc.infer_current_tolerance(bundle.asc_files))
        if not last_rows:
            continue
        area_cm2 = pc._bundle_area_cm2(bundle) if use_current_density else None
        points = [
            (pc._scaled_current(row["Corriente"], use_current_density, area_cm2), row["Voltaje"])
            for row in last_rows
        ]
        points.sort(key=lambda item: item[0])
        x_values = [point[0] for point in points]
        y_values = [point[1] for point in points]
        all_x.extend(x_values)
        all_y.extend(y_values)
        color = _color_for_index(index, len(bundles))
        ax.plot(
            x_values,
            y_values,
            marker="o",
            linestyle="-",
            linewidth=1.8,
            markersize=5.5,
            markerfacecolor="none",
            markeredgewidth=1.1,
            color=color,
            label=f"{_stage_label(bundle.curve_id, bundle.description, language)} - {bundle.description}",
        )

    if not all_x or not all_y:
        return None

    ax.set_title(translate("full_report_pc_summary_title", language))
    ax.set_xlabel(
        f"{translate('current_density', language)} (A/cm^2)"
        if use_current_density
        else f"{translate('current', language)} (A)"
    )
    ax.set_ylabel(f"{translate('voltage', language)} (V)")
    ax.grid(True)
    ax.xaxis.set_major_locator(MaxNLocator(nbins=6))
    ax.yaxis.set_major_locator(MaxNLocator(nbins=6))
    ax.xaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
    ax.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
    legend = ax.legend(fontsize=font_defaults.legend, loc="best")
    make_legend_draggable(legend)
    apply_plot_font_defaults(fig, font_defaults)
    fig.tight_layout()
    return fig


def _eis_current_group_key(entry: eis.EISPlotEntry) -> tuple[float | str, str]:
    if entry.current_value is not None:
        label = entry.current_label or f"{eis._format_report_value(entry.current_value)}A"
        return round(entry.current_value, 12), label
    label = entry.current_label or entry.voltage_label or entry.display_name
    return label, label


def _eis_current_groups(entries: list[eis.EISPlotEntry]) -> list[tuple[str, list[eis.EISPlotEntry]]]:
    grouped: dict[tuple[float | str, str], list[eis.EISPlotEntry]] = defaultdict(list)
    for entry in entries:
        if entry.current_value is None:
            continue
        grouped[_eis_current_group_key(entry)].append(entry)

    ordered_keys = sorted(
        grouped,
        key=lambda key: (
            not isinstance(key[0], float),
            key[0] if isinstance(key[0], float) else math.inf,
            key[1],
        ),
    )
    return [(key[1], grouped[key]) for key in ordered_keys]


def _draw_eis_nyquist_summary(
    current_label: str,
    entries: list[eis.EISPlotEntry],
    font_defaults: PlotFontDefaults,
    language: str,
) -> Figure | None:
    if not entries:
        return None

    fig = _new_plot_figure(figsize=(9.5, 6.2), dpi=150)
    ax = fig.add_subplot(111)
    all_x: list[float] = []
    all_y: list[float] = []
    first_parsed = entries[0].parsed

    ordered_entries = sorted(
        entries,
        key=lambda entry: (
            entry.stage_number is None,
            entry.stage_number if entry.stage_number is not None else math.inf,
            entry.display_name,
        ),
    )
    for index, entry in enumerate(ordered_entries):
        x_values, z_imag_values, freqs = eis._triplet_series(entry.parsed, "Zreal", "Zimag", "Freq")
        if not x_values or not z_imag_values:
            continue
        y_values = [-value for value in z_imag_values]
        all_x.extend(x_values)
        all_y.extend(y_values)
        color = entry.nyquist_color or _color_for_index(index, len(ordered_entries))
        marker = entry.default_marker or "o"
        ax.plot(
            x_values,
            y_values,
            marker=marker,
            linestyle="-",
            linewidth=1.6,
            markersize=4.8,
            markerfacecolor="none",
            markeredgecolor=color,
            color=color,
            label=_stage_label(entry.stage_number, entry.display_name, language),
        )
        eis._annotate_nyquist_max_y(ax, x_values, y_values, freqs, fontsize=font_defaults.tick)

    if not all_x or not all_y:
        return None

    x_unit = eis._impedance_unit(first_parsed, "Zreal")
    y_unit = eis._impedance_unit(first_parsed, "Zimag")
    ax.set_title(translate("full_report_eis_summary_title", language, current=current_label))
    ax.set_xlabel(f"Zreal ({x_unit})" if x_unit else "Zreal")
    ax.set_ylabel(f"-Zimag ({y_unit})" if y_unit else "-Zimag")
    ax.grid(True)
    eis._apply_nyquist_limits(ax, all_x, all_y, current_associated=True)
    ax.set_aspect("equal", adjustable="box")
    legend = ax.legend(fontsize=font_defaults.legend, loc="best")
    make_legend_draggable(legend)
    apply_plot_font_defaults(fig, font_defaults)
    fig.tight_layout()
    return fig


def _activation_visible_ramp_keys(bundle: activ.ActivationBundle) -> set[str]:
    return {ramp.key for ramp in activ.build_activation_ramps(bundle)}


def _draw_activation_local_summary(
    bundle: activ.ActivationBundle,
    font_defaults: PlotFontDefaults,
    language: str,
) -> Figure | None:
    visible_ramp_keys = _activation_visible_ramp_keys(bundle)
    if not visible_ramp_keys:
        return None

    time_unit = "h"
    limits = activ.compute_autofit_v_vs_t_limits(
        bundle,
        visible_ramp_keys=visible_ramp_keys,
        show_voltage=True,
        show_current=False,
        show_temperature=True,
        time_unit=time_unit,
        local_cycle_time=True,
    )
    fig = _new_plot_figure(figsize=(10.0, 6.4), dpi=150)
    has_plot = activ.draw_v_vs_t_on_figure(
        fig=fig,
        bundle=bundle,
        visible_ramp_keys=visible_ramp_keys,
        show_voltage=True,
        show_current=False,
        show_temperature=True,
        voltage_linestyle="-",
        current_linestyle="none",
        temperature_linestyle="--",
        time_unit=time_unit,
        local_cycle_time=True,
        x_tick_count=6,
        y_tick_count=6,
        t_min=_optional_float(limits.get("t_min")),
        t_max=_optional_float(limits.get("t_max")),
        v_min=_optional_float(limits.get("v_min")),
        v_max=_optional_float(limits.get("v_max")),
        temp_min=_optional_float(limits.get("temp_min")),
        temp_max=_optional_float(limits.get("temp_max")),
        current_min=None,
        current_max=None,
        plot_title=translate("full_report_activation_summary_title", language, curve=bundle.label),
        show_title=True,
        title_fontsize=font_defaults.title,
        tick_fontsize=font_defaults.tick,
        label_fontsize=font_defaults.label,
        legend_fontsize=font_defaults.legend,
        legend_scale=0.85,
        color_axes_by_magnitude=False,
        line_width=1.5,
        language=language,
    )
    return fig if has_plot else None


def _nyquist_y0_indicator(entry: eis.EISPlotEntry) -> tuple[float, str] | None:
    rows = eis.build_nyquist_indicator_rows(entry.parsed)
    if not rows:
        return None
    _label, value, unit = rows[0]
    numeric = _optional_float(value)
    if numeric is None:
        return None
    return numeric, unit


def _eis_stage_key(entry: eis.EISPlotEntry) -> int | str:
    return entry.stage_number if entry.stage_number is not None else entry.display_name


def _draw_eis_y0_bar_summary(
    current_groups: list[tuple[str, list[eis.EISPlotEntry]]],
    font_defaults: PlotFontDefaults,
    language: str,
) -> Figure | None:
    if not current_groups:
        return None

    stage_keys: list[int | str] = []
    stage_labels: dict[int | str, str] = {}
    stage_colors: dict[int | str, str] = {}
    group_values: list[dict[int | str, float]] = []
    y_unit = ""

    for _current_label, entries in current_groups:
        values_for_group: dict[int | str, float] = {}
        ordered_entries = sorted(
            entries,
            key=lambda entry: (
                entry.stage_number is None,
                entry.stage_number if entry.stage_number is not None else math.inf,
                entry.display_name,
            ),
        )
        for entry in ordered_entries:
            y0 = _nyquist_y0_indicator(entry)
            if y0 is None:
                continue
            value, unit = y0
            if not y_unit and unit:
                y_unit = unit
            stage_key = _eis_stage_key(entry)
            if stage_key not in stage_labels:
                stage_keys.append(stage_key)
                stage_labels[stage_key] = _stage_label(entry.stage_number, entry.display_name, language)
                if entry.nyquist_color:
                    stage_colors[stage_key] = entry.nyquist_color
            values_for_group[stage_key] = value
        group_values.append(values_for_group)

    if not stage_keys or not any(group_values):
        return None

    fig = _new_plot_figure(figsize=(10.0, 6.2), dpi=150)
    ax = fig.add_subplot(111)
    x_positions = list(range(len(current_groups)))
    bar_width = min(0.26, 0.78 / max(1, len(stage_keys)))

    for stage_index, stage_key in enumerate(stage_keys):
        x_values: list[float] = []
        y_values: list[float] = []
        offset = (stage_index - ((len(stage_keys) - 1) / 2.0)) * bar_width
        for group_index, values_for_group in enumerate(group_values):
            if stage_key not in values_for_group:
                continue
            x_values.append(x_positions[group_index] + offset)
            y_values.append(values_for_group[stage_key])
        if not x_values:
            continue
        ax.bar(
            x_values,
            y_values,
            width=bar_width * 0.88,
            label=stage_labels[stage_key],
            color=stage_colors.get(stage_key, _color_for_index(stage_index, len(stage_keys))),
            edgecolor="#2d3748",
            linewidth=0.35,
        )

    ax.set_title(translate("full_report_eis_y0_summary_title", language))
    ax.set_xlabel(translate("current", language))
    y_label = translate("y0_intersection", language)
    ax.set_ylabel(f"{y_label} ({y_unit})" if y_unit else y_label)
    ax.set_xticks(x_positions)
    ax.set_xticklabels([label for label, _entries in current_groups], rotation=20, ha="right")
    ax.grid(True, axis="y", alpha=0.35)
    ax.axhline(0.0, color="#4a5568", linewidth=0.8)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=6))
    ax.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
    ax.tick_params(axis="both", labelsize=font_defaults.tick)
    legend = ax.legend(fontsize=max(6.0, font_defaults.legend * 0.9), loc="best")
    make_legend_draggable(legend)
    apply_plot_font_defaults(fig, font_defaults)
    fig.tight_layout()
    return fig


def _sorted_deg_items(parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]]) -> list[tuple[deg.DegFile, deg.ParsedDTA]]:
    return sorted(parsed_items, key=lambda item: (item[0].stage, item[0].path.name.lower()))


def _deg_v_vs_t_kwargs(
    parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
    font_defaults: PlotFontDefaults,
    language: str,
    *,
    title_key: str,
) -> dict[str, object]:
    limits = deg.compute_autofit_v_vs_t_limits(parsed_items, show_temperature=True, time_unit="h")
    return {
        "show_temperature": True,
        "voltage_linestyle": "-",
        "temperature_linestyle": "--",
        "time_unit": "h",
        "x_tick_count": 6,
        "y_tick_count": 6,
        "plot_title": translate(title_key, language),
        "show_title": True,
        "title_fontsize": font_defaults.title,
        "tick_fontsize": font_defaults.tick,
        "label_fontsize": font_defaults.label,
        "legend_fontsize": max(6.0, font_defaults.legend * 0.85),
        "line_width": 1.5,
        "t_min": deg._parse_time_limit_text(limits["t_min"], "h"),
        "t_max": deg._parse_time_limit_text(limits["t_max"], "h"),
        "v_min": _optional_float(limits.get("v_min")),
        "v_max": _optional_float(limits.get("v_max")),
        "temp_min": _optional_float(limits.get("temp_min")),
        "temp_max": _optional_float(limits.get("temp_max")),
        "show_fit_line": False,
        "show_fit_range": False,
    }


def _deg_simple_slope_uv_h(parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]]) -> float | None:
    absolute_points: list[tuple[float, float]] = []
    fallback_first_voltage: float | None = None
    fallback_last_voltage: float | None = None
    fallback_total_seconds = 0.0

    for deg_file, parsed in _sorted_deg_items(parsed_items):
        try:
            time_values, voltage_values = deg._required_numeric_series(parsed, "T", "Vf")
        except ValueError:
            continue
        pairs = sorted(zip(time_values, voltage_values), key=lambda item: item[0])
        if len(pairs) < 2:
            continue

        stage_start_t, stage_start_v = pairs[0]
        stage_end_t, stage_end_v = pairs[-1]
        stage_duration = max(0.0, stage_end_t - stage_start_t)
        if stage_duration <= 0.0:
            continue

        try:
            stage_start_dt = deg._start_datetime(parsed, deg_file.path.name)
        except ValueError:
            stage_start_dt = None
        if stage_start_dt is not None:
            stage_epoch = stage_start_dt.timestamp()
            absolute_points.append((stage_epoch + stage_start_t, stage_start_v))
            absolute_points.append((stage_epoch + stage_end_t, stage_end_v))

        if fallback_first_voltage is None:
            fallback_first_voltage = stage_start_v
        fallback_last_voltage = stage_end_v
        fallback_total_seconds += stage_duration

    if len(absolute_points) >= 2:
        first_time, first_voltage = min(absolute_points, key=lambda item: item[0])
        last_time, last_voltage = max(absolute_points, key=lambda item: item[0])
        total_seconds = last_time - first_time
        if total_seconds > 0.0:
            return ((last_voltage - first_voltage) / total_seconds) * 1e6 * deg.SECONDS_PER_HOUR

    if fallback_first_voltage is None or fallback_last_voltage is None or fallback_total_seconds <= 0.0:
        return None
    return ((fallback_last_voltage - fallback_first_voltage) / fallback_total_seconds) * 1e6 * deg.SECONDS_PER_HOUR

def _deg_simple_slope_rows(
    parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
    language: str,
) -> list[tuple[str, object, str]]:
    slope = _deg_simple_slope_uv_h(parsed_items)
    if slope is None:
        return []
    return [(translate("simple_slope", language), _format_sig(slope), "µV/h")]


def _append_summary_indicator_rows(
    rows: list[SummaryIndicatorRow],
    pipeline: str,
    item: str,
    indicator_rows: list[tuple[str, object, str]],
) -> None:
    for indicator, value, unit in indicator_rows:
        rows.append((pipeline, item, indicator, value, unit))


def _summary_indicator_rows(
    activation_bundles: list[activ.ActivationBundle],
    pc_bundles: list[pc.CurveBundle],
    eis_entries: list[eis.EISPlotEntry],
    pre_entries: list[eis.EISPlotEntry],
    cv_datasets: list[cv.CVDataset],
    deg_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
    language: str,
) -> list[SummaryIndicatorRow]:
    rows: list[SummaryIndicatorRow] = []

    for bundle in activation_bundles:
        try:
            _append_summary_indicator_rows(
                rows,
                "Activacion",
                bundle.label,
                activ.build_activation_report_indicators(bundle, language=language),
            )
        except Exception:
            continue

    for bundle in pc_bundles:
        try:
            use_density = _pc_bundle_use_density(bundle)
            curve_label = f"{bundle.description} #{bundle.curve_id}"
            _append_summary_indicator_rows(
                rows,
                "PC",
                curve_label,
                pc.build_pc_report_indicators(bundle, use_current_density=use_density, language=language),
            )
            _append_summary_indicator_rows(
                rows,
                "PC dV/dI",
                curve_label,
                pc.build_pc_dv_di_report_indicators(
                    bundle,
                    point_fraction=1.0,
                    smoothing_algorithm="Median filter",
                    smoothing_window=1,
                    use_current_density=use_density,
                    high_current_fraction=0.90,
                    language=language,
                ),
            )
            _append_summary_indicator_rows(
                rows,
                "PC Step",
                curve_label,
                pc.build_pc_step_stability_report_indicators(bundle, use_current_density=use_density, language=language),
            )
        except Exception:
            continue

    for entry in eis_entries:
        try:
            _append_summary_indicator_rows(rows, "EIS Nyquist", entry.display_name, eis.build_nyquist_indicator_rows(entry.parsed))
            _append_summary_indicator_rows(rows, "EIS Bode", entry.display_name, eis.build_bode_indicator_rows(entry.parsed))
            _append_summary_indicator_rows(rows, "EIS Series Pt", entry.display_name, eis.build_series_pt_indicator_rows(entry.parsed))
        except Exception:
            continue

    for entry in pre_entries:
        try:
            _append_summary_indicator_rows(
                rows,
                "EIS Pre",
                entry.display_name,
                eis.build_pre_stabilization_indicator_rows(entry.parsed),
            )
        except Exception:
            continue

    for dataset in cv_datasets:
        try:
            visible_segment_keys = {segment.key for segment in dataset.segments if segment.cycle >= 2}
            if visible_segment_keys:
                _append_summary_indicator_rows(
                    rows,
                    "CV",
                    cv._dataset_stage_label(dataset),
                    cv._build_cv_report_indicator_rows(dataset, visible_segment_keys, language=language),
                )
        except Exception:
            continue

    if deg_items:
        try:
            _append_summary_indicator_rows(
                rows,
                "Deg",
                translate("deg_report_title", language),
                deg.build_deg_report_indicators(deg_items, language=language),
            )
        except Exception:
            pass

    return rows


def _draw_summary_indicators_pages(
    activation_bundles: list[activ.ActivationBundle],
    pc_bundles: list[pc.CurveBundle],
    eis_entries: list[eis.EISPlotEntry],
    pre_entries: list[eis.EISPlotEntry],
    cv_datasets: list[cv.CVDataset],
    deg_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
    language: str,
) -> list[Figure]:
    rows = _summary_indicator_rows(activation_bundles, pc_bundles, eis_entries, pre_entries, cv_datasets, deg_items, language)
    if not rows:
        return []

    rows_per_page = 30
    pages: list[Figure] = []
    for page_index, start in enumerate(range(0, len(rows), rows_per_page), start=1):
        page_rows = rows[start : start + rows_per_page]
        fig = _new_plot_figure(figsize=(11.0, 8.5), dpi=150)
        ax = fig.add_subplot(111)
        ax.axis("off")
        title = translate("full_report_summary_indicators_title", language)
        if len(rows) > rows_per_page:
            title = f"{title} ({page_index})"
        fig.text(0.04, 0.955, title, fontsize=17, fontweight="bold", ha="left", va="top")
        fig.text(
            0.04,
            0.915,
            translate("full_report_summary_indicators_subtitle", language),
            fontsize=9.5,
            ha="left",
            va="top",
            color="#4a5568",
        )
        _add_summary_indicator_table(ax, page_rows, [0.04, 0.06, 0.92, 0.80], language)
        pages.append(fig)
    return pages


def _draw_deg_summary(
    parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
    font_defaults: PlotFontDefaults,
    language: str,
) -> Figure | None:
    parsed_items = _sorted_deg_items(parsed_items)
    if not parsed_items:
        return None

    fig = _new_plot_figure(figsize=(10.0, 6.4), dpi=150)
    has_plot = deg.draw_v_vs_t_on_figure(
        fig=fig,
        parsed_items=parsed_items,
        **_deg_v_vs_t_kwargs(
            parsed_items,
            font_defaults,
            language,
            title_key="full_report_deg_summary_title",
        ),
    )
    if not has_plot:
        return None

    rows = _deg_simple_slope_rows(parsed_items, language)
    if rows and fig.axes:
        label, value, unit = rows[0]
        fig.axes[0].text(
            0.015,
            0.985,
            f"{label}: {value} {unit}".strip(),
            transform=fig.axes[0].transAxes,
            fontsize=max(7.0, font_defaults.legend * 0.9),
            va="top",
            ha="left",
            bbox=dict(boxstyle="round,pad=0.25", facecolor="white", edgecolor="#cbd5e0", alpha=0.88),
        )
    return fig


def _pc_v_vs_i_kwargs(bundle: pc.CurveBundle, font_defaults: PlotFontDefaults, language: str) -> dict[str, object]:
    use_density = _pc_bundle_use_density(bundle)
    limits = pc.compute_default_v_vs_i_limits(bundle, use_current_density=use_density)
    return {
        "show_asc": True,
        "show_dsc": True,
        "show_voltage": True,
        "show_temperature": False,
        "use_current_density": use_density,
        "point_fraction": 1.0,
        "asc_marker": "^",
        "dsc_marker": "v",
        "voltage_linestyle": "-",
        "temperature_linestyle": "none",
        "x_tick_count": 6,
        "y_tick_count": 6,
        "x_min": _optional_float(limits.get("x_min")),
        "x_max": _optional_float(limits.get("x_max")),
        "v_min": _optional_float(limits.get("v_min")),
        "v_max": _optional_float(limits.get("v_max")),
        "temp_min": None,
        "temp_max": None,
        "plot_title": "",
        "show_title": True,
        "title_fontsize": font_defaults.title,
        "tick_fontsize": font_defaults.tick,
        "label_fontsize": font_defaults.label,
        "legend_fontsize": font_defaults.legend,
        "marker_size": 6,
        "hollow_markers": False,
        "line_width": 1.5,
        "show_slope_guides": False,
        "indicator_current": 0.0,
        "language": language,
    }


def _pc_series_kwargs(bundle: pc.CurveBundle, font_defaults: PlotFontDefaults, language: str) -> dict[str, object]:
    time_unit = "min"
    use_density = _pc_bundle_use_density(bundle)
    limits = pc.compute_default_series_by_time_limits(bundle, time_unit=time_unit, use_current_density=use_density)
    return {
        "show_asc": True,
        "show_dsc": True,
        "show_voltage": True,
        "show_current": True,
        "show_temperature": True,
        "asc_marker": "^",
        "dsc_marker": "v",
        "voltage_linestyle": "-",
        "current_linestyle": "-.",
        "temperature_linestyle": "--",
        "use_current_density": use_density,
        "time_unit": time_unit,
        "x_tick_count": 6,
        "y_tick_count": 6,
        "t_min": _optional_float(limits.get("t_min")),
        "t_max": _optional_float(limits.get("t_max")),
        "v_min": _optional_float(limits.get("v_min")),
        "v_max": _optional_float(limits.get("v_max")),
        "i_min": _optional_float(limits.get("i_min")),
        "i_max": _optional_float(limits.get("i_max")),
        "temp_min": _optional_float(limits.get("temp_min")),
        "temp_max": _optional_float(limits.get("temp_max")),
        "plot_title": "",
        "show_title": True,
        "title_fontsize": font_defaults.title,
        "tick_fontsize": font_defaults.tick,
        "label_fontsize": font_defaults.label,
        "legend_fontsize": font_defaults.legend,
        "marker_size": 6,
        "line_width": 1.5,
        "hollow_markers": False,
        "language": language,
    }


def _pc_dvdi_kwargs(bundle: pc.CurveBundle, font_defaults: PlotFontDefaults, language: str) -> dict[str, object]:
    use_density = _pc_bundle_use_density(bundle)
    limits = pc.compute_default_dv_di_limits(bundle, use_current_density=use_density)
    return {
        "show_asc": True,
        "show_dsc": True,
        "point_fraction": 1.0,
        "smoothing_algorithm": "Median filter",
        "smoothing_window": 1,
        "logarithmic_y": True,
        "use_current_density": use_density,
        "asc_marker": "^",
        "dsc_marker": "v",
        "dvdi_linestyle": "-",
        "x_tick_count": 6,
        "y_tick_count": 6,
        "x_min": _optional_float(limits.get("x_min")),
        "x_max": _optional_float(limits.get("x_max")),
        "dvdi_min": _optional_float(limits.get("dvdi_min")),
        "dvdi_max": _optional_float(limits.get("dvdi_max")),
        "plot_title": "",
        "show_title": True,
        "title_fontsize": font_defaults.title,
        "tick_fontsize": font_defaults.tick,
        "label_fontsize": font_defaults.label,
        "legend_fontsize": font_defaults.legend,
        "marker_size": 6,
        "hollow_markers": False,
        "line_width": 1.5,
        "language": language,
    }


def _pc_step_kwargs(bundle: pc.CurveBundle, font_defaults: PlotFontDefaults, language: str) -> dict[str, object]:
    use_density = _pc_bundle_use_density(bundle)
    limits = pc.compute_default_step_stability_limits(bundle, use_current_density=use_density)
    return {
        "show_asc": True,
        "show_dsc": True,
        "show_voltage_delta": True,
        "show_temperature_delta": True,
        "use_current_density": use_density,
        "asc_marker": "^",
        "dsc_marker": "v",
        "voltage_linestyle": "-",
        "temperature_linestyle": "--",
        "x_tick_count": 6,
        "y_tick_count": 6,
        "x_min": _optional_float(limits.get("x_min")),
        "x_max": _optional_float(limits.get("x_max")),
        "dv_min": _optional_float(limits.get("dv_min")),
        "dv_max": _optional_float(limits.get("dv_max")),
        "dt_min": _optional_float(limits.get("dt_min")),
        "dt_max": _optional_float(limits.get("dt_max")),
        "plot_title": "",
        "show_title": True,
        "title_fontsize": font_defaults.title,
        "tick_fontsize": font_defaults.tick,
        "label_fontsize": font_defaults.label,
        "legend_fontsize": font_defaults.legend,
        "marker_size": 6,
        "hollow_markers": False,
        "line_width": 1.5,
        "language": language,
    }


def _append_pc_individual_pages(
    pdf: PdfPages,
    bundles: list[pc.CurveBundle],
    font_defaults: PlotFontDefaults,
    language: str,
    warnings: list[str],
    progress_callback: ProgressCallback | None,
    step: list[int],
    total: int,
) -> None:
    for bundle in bundles:
        curve_label = f"{bundle.description} #{bundle.curve_id}"
        use_density = _pc_bundle_use_density(bundle)
        _emit(progress_callback, f"PC: {curve_label}", step[0], total)

        try:
            v_kwargs = _pc_v_vs_i_kwargs(bundle, font_defaults, language)
            plot_fig = _new_plot_figure()
            if pc.draw_v_vs_i_on_figure(plot_fig, bundle=bundle, **v_kwargs):
                _save_figure(pdf, plot_fig)
                step[0] += 1

            temp_fig = _new_plot_figure()
            if pc.draw_temperature_vs_current_on_figure(
                temp_fig,
                bundle,
                point_fraction=1.0,
                use_current_density=use_density,
                x_tick_count=6,
                y_tick_count=6,
                title_fontsize=font_defaults.title,
                tick_fontsize=font_defaults.tick,
                label_fontsize=font_defaults.label,
                legend_fontsize=font_defaults.legend,
                marker_size=6,
                line_width=1.5,
                hollow_markers=False,
                language=language,
            ):
                _save_figure(pdf, temp_fig)
                step[0] += 1

            metadata_rows = pc.build_pc_report_metadata(bundle, use_current_density=use_density, language=language)
            indicator_rows = pc.build_pc_report_indicators(
                bundle,
                point_fraction=1.0,
                use_current_density=use_density,
                language=language,
                time_unit="s",
            )
            table_fig = _table_page(
                translate("pc_report_title", language, curve=curve_label),
                translate("pc_report_subtitle", language),
                [
                    (translate("metadata", language), metadata_rows, [0.05, 0.53, 0.90, 0.34]),
                    (translate("indicators", language), indicator_rows, [0.05, 0.08, 0.90, 0.36]),
                ],
                language,
            )
            _save_figure(pdf, table_fig)
            step[0] += 1
        except Exception as exc:
            warnings.append(f"PC V vs I {curve_label}: {type(exc).__name__}: {exc}")

        try:
            series_kwargs = _pc_series_kwargs(bundle, font_defaults, language)
            series_fig = _new_plot_figure(figsize=(10.0, 6.4))
            if pc.draw_series_by_time_on_figure(series_fig, bundle=bundle, **series_kwargs):
                _save_figure(pdf, series_fig)
                step[0] += 1

            metadata_rows = pc.build_pc_report_metadata(bundle, use_current_density=use_density, language=language)
            indicator_rows = pc.build_pc_report_indicators(
                bundle,
                point_fraction=1.0,
                use_current_density=use_density,
                language=language,
                time_unit="min",
            )
            table_fig = _table_page(
                translate("pc_series_report_title", language, curve=curve_label),
                translate("pc_series_report_subtitle", language),
                [
                    (translate("metadata", language), metadata_rows, [0.05, 0.53, 0.90, 0.34]),
                    (translate("indicators", language), indicator_rows, [0.05, 0.08, 0.90, 0.36]),
                ],
                language,
            )
            _save_figure(pdf, table_fig)
            step[0] += 1
        except Exception as exc:
            warnings.append(f"PC Series {curve_label}: {type(exc).__name__}: {exc}")

        try:
            dvdi_kwargs = _pc_dvdi_kwargs(bundle, font_defaults, language)
            dvdi_fig = _new_plot_figure()
            if pc.draw_dv_di_on_figure(dvdi_fig, bundle=bundle, **dvdi_kwargs):
                _save_figure(pdf, dvdi_fig)
                step[0] += 1

            metadata_rows = pc.build_pc_report_metadata(bundle, use_current_density=use_density, language=language)
            indicator_rows = pc.build_pc_dv_di_report_indicators(
                bundle,
                point_fraction=1.0,
                smoothing_algorithm="Median filter",
                smoothing_window=1,
                use_current_density=use_density,
                language=language,
            )
            table_fig = _table_page(
                translate("pc_dvdi_report_title", language, curve=curve_label),
                translate("pc_dvdi_report_subtitle", language),
                [
                    (translate("metadata", language), metadata_rows, [0.05, 0.53, 0.90, 0.34]),
                    (translate("indicators", language), indicator_rows, [0.05, 0.08, 0.90, 0.36]),
                ],
                language,
            )
            _save_figure(pdf, table_fig)
            step[0] += 1
        except Exception as exc:
            warnings.append(f"PC dV/dI {curve_label}: {type(exc).__name__}: {exc}")

        try:
            step_kwargs = _pc_step_kwargs(bundle, font_defaults, language)
            voltage_kwargs = dict(step_kwargs)
            voltage_kwargs.update(
                {
                    "show_voltage_delta": True,
                    "show_temperature_delta": False,
                    "plot_title": translate("voltage_step_range", language),
                }
            )
            step_fig = _new_plot_figure()
            if pc.draw_step_stability_on_figure(step_fig, bundle=bundle, **voltage_kwargs):
                _save_figure(pdf, step_fig)
                step[0] += 1

            temp_kwargs = dict(step_kwargs)
            temp_kwargs.update(
                {
                    "show_voltage_delta": False,
                    "show_temperature_delta": True,
                    "plot_title": translate("temperature_step_range", language),
                    "dv_min": None,
                    "dv_max": None,
                }
            )
            temp_step_fig = _new_plot_figure()
            if pc.draw_step_stability_on_figure(temp_step_fig, bundle=bundle, **temp_kwargs):
                _save_figure(pdf, temp_step_fig)
                step[0] += 1

            metadata_rows = pc.build_pc_report_metadata(bundle, use_current_density=use_density, language=language)
            indicator_rows = pc.build_pc_step_stability_report_indicators(
                bundle,
                use_current_density=use_density,
                language=language,
            )
            table_fig = _table_page(
                translate("pc_step_stability_report_title", language, curve=curve_label),
                translate("pc_step_stability_report_subtitle", language),
                [
                    (translate("metadata", language), metadata_rows, [0.05, 0.53, 0.90, 0.34]),
                    (translate("indicators", language), indicator_rows, [0.05, 0.08, 0.90, 0.36]),
                ],
                language,
            )
            _save_figure(pdf, table_fig)
            step[0] += 1
        except Exception as exc:
            warnings.append(f"PC Step Stability {curve_label}: {type(exc).__name__}: {exc}")


def _pre_stab_metadata_rows(parsed: eis.ParsedDTA, language: str) -> list[tuple[str, object, str]]:
    rows: list[tuple[str, object, str]] = []
    for key, label in eis.PRE_STAB_META_FIELDS:
        raw_value = parsed.meta_values.get(key, "")
        numeric = eis.to_float(raw_value) if key in {"ISTEP1", "TSTEP1", "SAMPLETIME", "AREA"} else None
        rows.append(
            (
                eis._localized_meta_label(label, language),
                numeric if numeric is not None else raw_value,
                parsed.meta_units.get(key, ""),
            )
        )
    return rows


def _append_eis_individual_pages(
    pdf: PdfPages,
    entries: list[eis.EISPlotEntry],
    pre_entries: list[eis.EISPlotEntry],
    font_defaults: PlotFontDefaults,
    language: str,
    warnings: list[str],
    progress_callback: ProgressCallback | None,
    step: list[int],
    total: int,
) -> None:
    for entry in entries:
        _emit(progress_callback, f"EIS: {entry.display_name}", step[0], total)
        try:
            nyquist_fig = eis.fig_nyquist(
                entry.parsed,
                plot_title=entry.display_name,
                line_color=entry.nyquist_color,
                marker_style=entry.default_marker,
                current_value=entry.current_value,
                voltage_label=entry.voltage_label,
                annotate_max_imaginary=True,
                font_defaults=font_defaults,
            )
            if nyquist_fig is not None:
                FigureCanvasAgg(nyquist_fig)
                _save_figure(pdf, nyquist_fig)
                step[0] += 1

            bode_fig = eis.fig_bode(
                entry.parsed,
                plot_title=f"{entry.display_name} - Bode",
                marker_style=entry.default_marker,
                font_defaults=font_defaults,
            )
            if bode_fig is not None:
                FigureCanvasAgg(bode_fig)
                _save_figure(pdf, bode_fig)
                step[0] += 1

            series_fig = eis.fig_series_vs_pt(
                entry.parsed,
                plot_title=f"{entry.display_name} - Series vs Pt",
                font_defaults=font_defaults,
            )
            if series_fig is not None:
                eis._update_right_axis_spacing(series_fig, FigureCanvasAgg(series_fig), getattr(series_fig, "_pt_axes", {}))
                _save_figure(pdf, series_fig)
                step[0] += 1

            table_fig = _table_page(
                f"EIS - {entry.display_name}",
                "",
                [
                    (
                        translate("metadata", language),
                        eis._build_eis_report_metadata_rows(entry.parsed, language),
                        [0.05, 0.62, 0.90, 0.26],
                    ),
                    (
                        translate("nyquist_indicators", language),
                        eis.build_nyquist_indicator_rows(entry.parsed),
                        [0.05, 0.41, 0.90, 0.13],
                    ),
                    (
                        translate("bode_indicators", language),
                        eis.build_bode_indicator_rows(entry.parsed),
                        [0.05, 0.14, 0.90, 0.19],
                    ),
                ],
                language,
            )
            _save_figure(pdf, table_fig)
            step[0] += 1
        except Exception as exc:
            warnings.append(f"EIS {entry.display_name}: {type(exc).__name__}: {exc}")

    for entry in pre_entries:
        _emit(progress_callback, f"EIS Pre: {entry.display_name}", step[0], total)
        try:
            pre_fig = eis.fig_pre_stabilization(entry, font_defaults=font_defaults)
            if pre_fig is not None:
                eis._update_right_axis_spacing(pre_fig, FigureCanvasAgg(pre_fig), getattr(pre_fig, "_pre_stab_axes", {}))
                _save_figure(pdf, pre_fig)
                step[0] += 1

            table_fig = _table_page(
                translate("eis_pre_stabilization_report_title", language, curve=entry.display_name),
                "",
                [
                    (
                        translate("metadata", language),
                        _pre_stab_metadata_rows(entry.parsed, language),
                        [0.05, 0.58, 0.90, 0.29],
                    ),
                    (
                        translate("pre_stabilization_indicators", language),
                        eis.build_pre_stabilization_indicator_rows(entry.parsed),
                        [0.05, 0.06, 0.90, 0.43],
                    ),
                ],
                language,
            )
            _save_figure(pdf, table_fig)
            step[0] += 1
        except Exception as exc:
            warnings.append(f"EIS Pre {entry.display_name}: {type(exc).__name__}: {exc}")


def _append_cv_individual_pages(
    pdf: PdfPages,
    datasets: list[cv.CVDataset],
    font_defaults: PlotFontDefaults,
    language: str,
    warnings: list[str],
    progress_callback: ProgressCallback | None,
    step: list[int],
    total: int,
) -> None:
    for dataset in datasets:
        label = cv._dataset_stage_label(dataset)
        _emit(progress_callback, f"CV: {label}", step[0], total)
        try:
            visible_segment_keys = {segment.key for segment in dataset.segments}
            report_segment_keys = {
                segment.key
                for segment in dataset.segments
                if segment.key in visible_segment_keys and segment.cycle >= 2
            }
            if not report_segment_keys:
                continue

            limits = cv.compute_autofit_i_vs_v_limits(
                dataset=dataset,
                visible_segment_keys=report_segment_keys,
                show_current=True,
                show_temperature=False,
            )
            plot_fig = _new_plot_figure()
            if cv.draw_i_vs_v_on_figure(
                fig=plot_fig,
                dataset=dataset,
                visible_segment_keys=report_segment_keys,
                show_current=True,
                show_temperature=False,
                current_linestyle="-",
                temperature_linestyle="none",
                color_axes_by_magnitude=False,
                x_tick_count=6,
                y_tick_count=6,
                v_min=_optional_float(limits.get("v_min")),
                v_max=_optional_float(limits.get("v_max")),
                i_min=_optional_float(limits.get("i_min")),
                i_max=_optional_float(limits.get("i_max")),
                temp_min=None,
                temp_max=None,
                plot_title=f"j vs V - CV {label}",
                show_title=True,
                title_fontsize=font_defaults.title,
                tick_fontsize=font_defaults.tick,
                label_fontsize=font_defaults.label,
                legend_fontsize=font_defaults.legend,
                legend_scale=1.0,
                line_width=1.5,
                language=language,
            ):
                _save_figure(pdf, plot_fig)
                step[0] += 1

            table_fig = _table_page(
                translate("cv_i_vs_v_report_title", language, curve=label),
                "",
                [
                    (translate("metadata", language), cv._build_metadata_rows(dataset.parsed, language=language), [0.05, 0.53, 0.90, 0.34]),
                    (
                        translate("cv_indicators", language),
                        cv._build_cv_report_indicator_rows(dataset, report_segment_keys, language=language),
                        [0.05, 0.24, 0.90, 0.18],
                    ),
                ],
                language,
            )
            _save_figure(pdf, table_fig)
            step[0] += 1
        except Exception as exc:
            warnings.append(f"CV {label}: {type(exc).__name__}: {exc}")


def _append_activation_individual_pages(
    pdf: PdfPages,
    bundles: list[activ.ActivationBundle],
    font_defaults: PlotFontDefaults,
    language: str,
    warnings: list[str],
    progress_callback: ProgressCallback | None,
    step: list[int],
    total: int,
) -> None:
    for bundle in bundles:
        _emit(progress_callback, f"Activacion: {bundle.label}", step[0], total)
        visible_ramp_keys = _activation_visible_ramp_keys(bundle)
        if not visible_ramp_keys:
            continue

        try:
            global_fig = _new_plot_figure(figsize=(10.0, 6.4), dpi=150)
            if activ.draw_activation_report_time_on_figure(
                global_fig,
                bundle,
                visible_ramp_keys,
                local_cycle_time=False,
                time_unit="h",
                language=language,
                title_fontsize=font_defaults.title,
                tick_fontsize=font_defaults.tick,
                label_fontsize=font_defaults.label,
                legend_fontsize=font_defaults.legend,
                line_width=1.5,
                x_tick_count=6,
                y_tick_count=6,
            ):
                _save_figure(pdf, global_fig)
                step[0] += 1

            local_fig = _new_plot_figure(figsize=(10.0, 6.4), dpi=150)
            if activ.draw_activation_report_time_on_figure(
                local_fig,
                bundle,
                visible_ramp_keys,
                local_cycle_time=True,
                time_unit="h",
                language=language,
                title_fontsize=font_defaults.title,
                tick_fontsize=font_defaults.tick,
                label_fontsize=font_defaults.label,
                legend_fontsize=font_defaults.legend,
                line_width=1.5,
                x_tick_count=6,
                y_tick_count=6,
            ):
                _save_figure(pdf, local_fig)
                step[0] += 1

            table_fig = _table_page(
                translate("activation_report_title", language, curve=bundle.label),
                translate("activation_report_subtitle", language),
                [
                    (
                        translate("metadata", language),
                        activ.build_activation_report_metadata(bundle, language=language),
                        [0.05, 0.55, 0.90, 0.32],
                    ),
                    (
                        translate("indicators", language),
                        activ.build_activation_report_indicators(bundle, language=language),
                        [0.05, 0.08, 0.90, 0.38],
                    ),
                ],
                language,
            )
            _save_figure(pdf, table_fig)
            step[0] += 1
        except Exception as exc:
            warnings.append(f"Activacion {bundle.label}: {type(exc).__name__}: {exc}")


def _append_deg_individual_pages(
    pdf: PdfPages,
    parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
    font_defaults: PlotFontDefaults,
    language: str,
    warnings: list[str],
    progress_callback: ProgressCallback | None,
    step: list[int],
    total: int,
) -> None:
    parsed_items = _sorted_deg_items(parsed_items)
    if not parsed_items:
        return

    _emit(progress_callback, translate("deg_report_title", language), step[0], total)
    try:
        plot_fig = _new_plot_figure(figsize=(10.0, 6.4), dpi=150)
        if deg.draw_v_vs_t_on_figure(
            fig=plot_fig,
            parsed_items=parsed_items,
            **_deg_v_vs_t_kwargs(
                parsed_items,
                font_defaults,
                language,
                title_key="deg_report_plot_title",
            ),
        ):
            _save_figure(pdf, plot_fig)
            step[0] += 1

        table_fig = _table_page(
            translate("deg_report_title", language),
            translate("deg_report_subtitle", language),
            [
                (
                    translate("metadata", language),
                    deg.build_deg_report_metadata(parsed_items, language=language),
                    [0.05, 0.53, 0.90, 0.34],
                ),
                (
                    translate("indicators", language),
                    deg.build_deg_report_indicators(parsed_items, language=language),
                    [0.05, 0.23, 0.90, 0.20],
                ),
            ],
            language,
        )
        _save_figure(pdf, table_fig)
        step[0] += 1
    except Exception as exc:
        warnings.append(f"Deg: {type(exc).__name__}: {exc}")


def _append_warnings_page(pdf: PdfPages, warnings: list[str], language: str) -> None:
    if not warnings:
        return
    fig = _new_plot_figure(figsize=(8.5, 11.0), dpi=150)
    ax = fig.add_subplot(111)
    ax.axis("off")
    fig.text(0.05, 0.965, translate("full_report_warnings", language), fontsize=18, fontweight="bold", ha="left", va="top")
    y = 0.91
    for warning in warnings[:35]:
        fig.text(0.07, y, f"- {warning}", fontsize=8.5, ha="left", va="top", color="#4a5568", wrap=True)
        y -= 0.028
    if len(warnings) > 35:
        fig.text(0.07, y, f"... {len(warnings) - 35} more", fontsize=8.5, ha="left", va="top", color="#4a5568")
    _save_figure(pdf, fig)


def _estimate_total_pages(
    activation_bundles: list[activ.ActivationBundle],
    pc_bundles: list[pc.CurveBundle],
    eis_entries: list[eis.EISPlotEntry],
    pre_entries: list[eis.EISPlotEntry],
    cv_datasets: list[cv.CVDataset],
    current_groups: list[tuple[str, list[eis.EISPlotEntry]]],
    deg_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
) -> int:
    total = 2
    if activation_bundles or pc_bundles or current_groups or pre_entries or cv_datasets or deg_items:
        total += 1
    try:
        summary_indicator_count = len(
            _summary_indicator_rows(activation_bundles, pc_bundles, eis_entries, pre_entries, cv_datasets, deg_items, "es")
        )
    except Exception:
        summary_indicator_count = 1
    if summary_indicator_count:
        total += max(1, math.ceil(summary_indicator_count / 30))
    total += len(activation_bundles)
    if pc_bundles:
        total += 1
    total += len(current_groups)
    if current_groups:
        total += 1
    if deg_items:
        total += 1
    if activation_bundles:
        total += 1 + (3 * len(activation_bundles))
    if pc_bundles:
        total += 1 + (10 * len(pc_bundles))
    if eis_entries or pre_entries:
        total += 1 + (4 * len(eis_entries)) + (2 * len(pre_entries))
    if cv_datasets:
        total += 1 + (2 * len(cv_datasets))
    if deg_items:
        total += 3
    return max(total, 1)


def generate_full_report(
    input_dir: Path,
    output_dir: Path,
    *,
    language: str = "es",
    font_defaults: PlotFontDefaults | None = None,
    progress_callback: ProgressCallback | None = None,
) -> Path:
    language = normalize_language(language)
    font_defaults = resolve_plot_font_defaults(font_defaults)
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    _emit(progress_callback, translate("full_report_discovering", language), 0, None)
    pc.PC_LANGUAGE = language
    eis.EIS_LANGUAGE = language
    cv.CV_LANGUAGE = language

    activation_bundles = activ.discover_activation_bundles(input_dir)
    pc_bundles = pc.discover_curve_bundles(input_dir)
    eis_files = eis.find_eis_files(input_dir)
    eis_entries = eis._collect_eis_plot_entries(eis_files, language=language)
    pre_files = eis.find_pre_stabilization_files(input_dir)
    pre_entries = eis._collect_pre_stabilization_entries(pre_files, language=language)
    cv_datasets = cv.discover_cv_datasets(input_dir)
    deg_files = deg.find_deg_files(input_dir)
    deg_items = _sorted_deg_items([(deg_file, deg.parse_gamry_dta(deg_file.path)) for deg_file in deg_files])
    current_groups = _eis_current_groups(eis_entries)

    if not (activation_bundles or pc_bundles or eis_entries or pre_entries or cv_datasets or deg_items):
        raise ValueError(translate("full_report_no_data", language))

    counts = {
        "activation": len(activation_bundles),
        "pc": len(pc_bundles),
        "eis": len(eis_entries),
        "pre_stab": len(pre_entries),
        "cv": len(cv_datasets),
        "deg": len(deg_items),
    }
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"Full_Report_{_safe_filename_part(input_dir.name)}_{timestamp}.pdf"
    total = _estimate_total_pages(
        activation_bundles,
        pc_bundles,
        eis_entries,
        pre_entries,
        cv_datasets,
        current_groups,
        deg_items,
    )
    step = [0]
    warnings: list[str] = []
    section_titles: list[str] = []
    if activation_bundles or pc_bundles or current_groups or pre_entries or cv_datasets or deg_items:
        section_titles.append(translate("full_report_summary", language))
    if activation_bundles:
        section_titles.append("Activacion")
    if pc_bundles:
        section_titles.append("PC")
    if eis_entries or pre_entries:
        section_titles.append("EIS")
    if cv_datasets:
        section_titles.append("CV")
    if deg_items:
        section_titles.append("Deg")
    section_entries: list[tuple[str, int]] = []

    with PdfPages(output_path) as pdf:
        _emit(progress_callback, translate("full_report_generating", language), step[0], total)
        _save_figure(pdf, _title_page(input_dir, output_dir, counts, language))
        _bookmark(section_entries, translate("full_report_title", language), step[0])
        step[0] += 1
        _save_figure(pdf, _index_page(section_titles, language))
        _bookmark(section_entries, translate("full_report_index", language), step[0])
        step[0] += 1

        if activation_bundles or pc_bundles or current_groups or pre_entries or cv_datasets or deg_items:
            _bookmark(section_entries, translate("full_report_summary", language), step[0])
            _save_figure(
                pdf,
                _section_page(
                    translate("full_report_summary", language),
                    translate("full_report_summary_subtitle", language),
                    language,
                ),
            )
            step[0] += 1

        summary_indicator_figs = _draw_summary_indicators_pages(
            activation_bundles,
            pc_bundles,
            eis_entries,
            pre_entries,
            cv_datasets,
            deg_items,
            language,
        )
        if summary_indicator_figs:
            title = translate("full_report_summary_indicators_title", language)
            for page_index, summary_indicator_fig in enumerate(summary_indicator_figs):
                _emit(progress_callback, title, step[0], total)
                if page_index == 0:
                    _bookmark(section_entries, title, step[0])
                _save_figure(pdf, summary_indicator_fig)
                step[0] += 1

        for bundle in activation_bundles:
            title = translate("full_report_activation_summary_title", language, curve=bundle.label)
            _emit(
                progress_callback,
                title,
                step[0],
                total,
            )
            activation_summary_fig = _draw_activation_local_summary(bundle, font_defaults, language)
            if activation_summary_fig is not None:
                _bookmark(section_entries, title, step[0])
                _save_figure(pdf, activation_summary_fig)
                step[0] += 1

        if pc_bundles:
            title = translate("full_report_pc_summary_title", language)
            _emit(progress_callback, title, step[0], total)
            pc_summary_fig = _draw_pc_ascending_summary(pc_bundles, font_defaults, language)
            if pc_summary_fig is not None:
                _bookmark(section_entries, title, step[0])
                _save_figure(pdf, pc_summary_fig)
                step[0] += 1

        for current_label, group_entries in current_groups:
            title = translate("full_report_eis_summary_title", language, current=current_label)
            _emit(
                progress_callback,
                title,
                step[0],
                total,
            )
            summary_fig = _draw_eis_nyquist_summary(current_label, group_entries, font_defaults, language)
            if summary_fig is not None:
                _bookmark(section_entries, title, step[0])
                _save_figure(pdf, summary_fig)
                step[0] += 1

        if current_groups:
            title = translate("full_report_eis_y0_summary_title", language)
            _emit(progress_callback, title, step[0], total)
            y0_summary_fig = _draw_eis_y0_bar_summary(current_groups, font_defaults, language)
            if y0_summary_fig is not None:
                _bookmark(section_entries, title, step[0])
                _save_figure(pdf, y0_summary_fig)
                step[0] += 1

        if deg_items:
            title = translate("full_report_deg_summary_title", language)
            _emit(progress_callback, title, step[0], total)
            deg_summary_fig = _draw_deg_summary(deg_items, font_defaults, language)
            if deg_summary_fig is not None:
                _bookmark(section_entries, title, step[0])
                _save_figure(pdf, deg_summary_fig)
                step[0] += 1

        if activation_bundles:
            _bookmark(section_entries, "Activacion", step[0])
            _save_figure(
                pdf,
                _section_page(
                    "Activacion",
                    translate("full_report_individual_subtitle", language),
                    language,
                ),
            )
            step[0] += 1
            _append_activation_individual_pages(
                pdf,
                activation_bundles,
                font_defaults,
                language,
                warnings,
                progress_callback,
                step,
                total,
            )

        if pc_bundles:
            _bookmark(section_entries, "PC", step[0])
            _save_figure(
                pdf,
                _section_page(
                    "PC",
                    translate("full_report_individual_subtitle", language),
                    language,
                ),
            )
            step[0] += 1
            _append_pc_individual_pages(pdf, pc_bundles, font_defaults, language, warnings, progress_callback, step, total)

        if eis_entries or pre_entries:
            _bookmark(section_entries, "EIS", step[0])
            _save_figure(
                pdf,
                _section_page(
                    "EIS",
                    translate("full_report_individual_subtitle", language),
                    language,
                ),
            )
            step[0] += 1
            _append_eis_individual_pages(pdf, eis_entries, pre_entries, font_defaults, language, warnings, progress_callback, step, total)

        if cv_datasets:
            _bookmark(section_entries, "CV", step[0])
            _save_figure(
                pdf,
                _section_page(
                    "CV",
                    translate("full_report_individual_subtitle", language),
                    language,
                ),
            )
            step[0] += 1
            _append_cv_individual_pages(pdf, cv_datasets, font_defaults, language, warnings, progress_callback, step, total)

        if deg_items:
            _bookmark(section_entries, "Deg", step[0])
            _save_figure(
                pdf,
                _section_page(
                    "Deg",
                    translate("full_report_individual_subtitle", language),
                    language,
                ),
            )
            step[0] += 1
            _append_deg_individual_pages(pdf, deg_items, font_defaults, language, warnings, progress_callback, step, total)

        _append_warnings_page(pdf, warnings, language)

    _add_pdf_outline(output_path, section_entries)
    _emit(progress_callback, translate("full_report_completed", language), total, total)
    return output_path
