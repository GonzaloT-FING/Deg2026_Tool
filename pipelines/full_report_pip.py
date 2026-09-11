from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime, timedelta
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
from pipelines import ocp_pip as ocp
from pipelines import pol_cur_pip as pc


ProgressCallback = Callable[[str, int | None, int | None], None]
RELEVANT_SUMMARY_ROWS_PER_PAGE = 28


@dataclass
class PdfOutlineEntry:
    title: str
    page_index: int
    children: list["PdfOutlineEntry"] = field(default_factory=list)


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


def _stage_bookmark_label(prefix: str, stage_number: int | None, fallback: str, language: str) -> str:
    return f"{prefix} - {_stage_label(stage_number, fallback, language)}"


def _pc_curve_label(bundle: pc.CurveBundle, language: str) -> str:
    return f"{_stage_label(bundle.curve_id, bundle.description, language)} - {bundle.description} #{bundle.curve_id}"


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
        ("EIS", counts.get("eis", 0) + counts.get("pre_stab", 0), "EIS"),
        ("CV", counts.get("cv", 0), "CV"),
        ("Deg", counts.get("deg", 0), translate("deg_stage_count", language)),
        ("Deg OCP", counts.get("deg_ocp", 0), "OCP"),
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


def _pdf_text_string(text: str) -> bytes:
    data = b"\xfe\xff" + str(text).encode("utf-16-be")
    return b"<" + data.hex().upper().encode("ascii") + b">"


def _latest_pdf_object(data: bytes, object_id: int) -> bytes | None:
    pattern = re.compile(rb"(?m)^" + str(object_id).encode("ascii") + rb"\s+0\s+obj\s*(.*?)\s*endobj", re.DOTALL)
    matches = list(pattern.finditer(data))
    if not matches:
        return None
    return matches[-1].group(1).strip()


def _last_pdf_trailer(data: bytes) -> tuple[bytes, int] | None:
    pattern = re.compile(rb"trailer\s*<<(.*?)>>\s*startxref\s*(\d+)\s*%%EOF", re.DOTALL)
    matches = list(pattern.finditer(data))
    if not matches:
        return None
    last = matches[-1]
    return last.group(1), int(last.group(2))


def _page_references_from_pdf(data: bytes, catalog_body: bytes) -> list[int]:
    pages_match = re.search(rb"/Pages\s+(\d+)\s+\d+\s+R", catalog_body)
    if pages_match:
        pages_body = _latest_pdf_object(data, int(pages_match.group(1)))
        if pages_body is not None:
            kids_match = re.search(rb"/Kids\s*\[(.*?)\]", pages_body, re.DOTALL)
            if kids_match:
                page_ids = [
                    int(match.group(1))
                    for match in re.finditer(rb"(\d+)\s+\d+\s+R", kids_match.group(1))
                ]
                if page_ids:
                    return page_ids

    page_pattern = re.compile(rb"(?m)^(\d+)\s+0\s+obj\s*<<.*?/Type\s*/Page\b.*?endobj", re.DOTALL)
    return [int(match.group(1)) for match in page_pattern.finditer(data)]


def _catalog_with_outline(catalog_body: bytes, outline_object_id: int) -> bytes:
    body = catalog_body.strip()
    body = re.sub(rb"/Outlines\s+\d+\s+\d+\s+R", b"", body)
    body = re.sub(rb"/PageMode\s*/[A-Za-z0-9]+", b"", body)
    if body.startswith(b"<<") and body.endswith(b">>"):
        return (
            body[:-2].rstrip()
            + b"\n/Outlines "
            + str(outline_object_id).encode("ascii")
            + b" 0 R\n/PageMode /UseOutlines\n>>"
        )

    pages_match = re.search(rb"/Pages\s+(\d+)\s+\d+\s+R", body)
    if pages_match:
        pages_object_id = int(pages_match.group(1))
        return (
            b"<< /Type /Catalog /Pages "
            + str(pages_object_id).encode("ascii")
            + b" 0 R /Outlines "
            + str(outline_object_id).encode("ascii")
            + b" 0 R /PageMode /UseOutlines >>"
        )
    raise ValueError("No se pudo localizar el catalogo PDF.")


def _add_pdf_outline_incremental(output_path: Path, section_entries: list[tuple[str, int]]) -> None:
    data = output_path.read_bytes()
    trailer = _last_pdf_trailer(data)
    if trailer is None:
        raise ValueError("No se pudo localizar el trailer PDF.")
    trailer_body, previous_xref = trailer

    root_match = re.search(rb"/Root\s+(\d+)\s+\d+\s+R", trailer_body)
    size_match = re.search(rb"/Size\s+(\d+)", trailer_body)
    if root_match is None or size_match is None:
        raise ValueError("No se pudo localizar el catalogo PDF.")

    root_object_id = int(root_match.group(1))
    catalog_body = _latest_pdf_object(data, root_object_id)
    if catalog_body is None:
        raise ValueError("No se pudo leer el catalogo PDF.")

    page_object_ids = _page_references_from_pdf(data, catalog_body)
    if not page_object_ids:
        raise ValueError("No se pudieron localizar las paginas PDF.")

    entries = [
        (title, page_index)
        for title, page_index in section_entries
        if 0 <= page_index < len(page_object_ids)
    ]
    if not entries:
        return

    next_object_id = int(size_match.group(1))
    outline_object_id = next_object_id
    item_object_ids = list(range(next_object_id + 1, next_object_id + 1 + len(entries)))
    new_size = next_object_id + 1 + len(entries)

    def _reference(object_id: int) -> bytes:
        return str(object_id).encode("ascii") + b" 0 R"

    outline_body = (
        b"<< /Type /Outlines /First "
        + _reference(item_object_ids[0])
        + b" /Last "
        + _reference(item_object_ids[-1])
        + b" /Count "
        + str(len(item_object_ids)).encode("ascii")
        + b" >>"
    )
    root_body = _catalog_with_outline(catalog_body, outline_object_id)

    objects: list[tuple[int, bytes]] = [(root_object_id, root_body), (outline_object_id, outline_body)]
    for index, ((title, page_index), item_object_id) in enumerate(zip(entries, item_object_ids)):
        parts = [
            b"<< /Title ",
            _pdf_text_string(title),
            b" /Parent ",
            _reference(outline_object_id),
            b" /Dest [",
            _reference(page_object_ids[page_index]),
            b" /Fit]",
        ]
        if index > 0:
            parts.extend([b" /Prev ", _reference(item_object_ids[index - 1])])
        if index < len(item_object_ids) - 1:
            parts.extend([b" /Next ", _reference(item_object_ids[index + 1])])
        parts.append(b" >>")
        objects.append((item_object_id, b"".join(parts)))

    appended = bytearray()
    offsets: dict[int, int] = {}
    base_offset = len(data)
    for object_id, body in objects:
        offsets[object_id] = base_offset + len(appended)
        appended.extend(str(object_id).encode("ascii") + b" 0 obj\n")
        appended.extend(body)
        appended.extend(b"\nendobj\n")

    startxref = base_offset + len(appended)
    appended.extend(b"xref\n")
    appended.extend(str(root_object_id).encode("ascii") + b" 1\n")
    appended.extend(f"{offsets[root_object_id]:010d} 00000 n \n".encode("ascii"))
    appended.extend(str(outline_object_id).encode("ascii") + b" " + str(len(item_object_ids) + 1).encode("ascii") + b"\n")
    for object_id in [outline_object_id, *item_object_ids]:
        appended.extend(f"{offsets[object_id]:010d} 00000 n \n".encode("ascii"))

    trailer_parts = [
        b"trailer\n<< /Size ",
        str(new_size).encode("ascii"),
        b" /Root ",
        _reference(root_object_id),
    ]
    info_match = re.search(rb"/Info\s+(\d+)\s+\d+\s+R", trailer_body)
    if info_match:
        trailer_parts.extend([b" /Info ", _reference(int(info_match.group(1)))])
    trailer_parts.extend([b" /Prev ", str(previous_xref).encode("ascii"), b" >>\n"])
    trailer_parts.extend([b"startxref\n", str(startxref).encode("ascii"), b"\n%%EOF\n"])
    appended.extend(b"".join(trailer_parts))

    with output_path.open("ab") as handle:
        handle.write(appended)


def _flatten_outline_entries(entries: list[PdfOutlineEntry]) -> list[tuple[str, int]]:
    flattened: list[tuple[str, int]] = []
    for entry in entries:
        flattened.append((entry.title, entry.page_index))
        flattened.extend(_flatten_outline_entries(entry.children))
    return flattened


def _add_pdf_outline(output_path: Path, outline_entries: list[PdfOutlineEntry]) -> None:
    if not outline_entries:
        return
    try:
        from pypdf import PdfReader, PdfWriter

        reader = PdfReader(str(output_path))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        def _write_entries(entries: list[PdfOutlineEntry], parent=None) -> None:
            for entry in entries:
                if 0 <= entry.page_index < len(reader.pages):
                    node = writer.add_outline_item(entry.title, entry.page_index, parent=parent)
                    _write_entries(entry.children, parent=node)

        _write_entries(outline_entries)
        temp_path = output_path.with_name(f"{output_path.stem}.tmp{output_path.suffix}")
        with temp_path.open("wb") as handle:
            writer.write(handle)
        temp_path.replace(output_path)
        return
    except Exception:
        pass

    try:
        _add_pdf_outline_incremental(output_path, _flatten_outline_entries(outline_entries))
    except Exception:
        return


def _bookmark(outline_entries: list[PdfOutlineEntry], title: str, page_index: int) -> PdfOutlineEntry:
    entry = PdfOutlineEntry(title, page_index)
    outline_entries.append(entry)
    return entry


def _bookmark_child(parent: PdfOutlineEntry | None, title: str, page_index: int) -> PdfOutlineEntry | None:
    if parent is None:
        return None
    return _bookmark(parent.children, title, page_index)


def _remove_bookmark_child(parent: PdfOutlineEntry | None, entry: PdfOutlineEntry | None) -> None:
    if parent is None or entry is None:
        return
    try:
        parent.children.remove(entry)
    except ValueError:
        pass


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


def _pc_first_last_bundles(bundles: list[pc.CurveBundle]) -> tuple[pc.CurveBundle, pc.CurveBundle] | None:
    if len(bundles) < 2:
        return None
    ordered = sorted(bundles, key=lambda item: (item.curve_id, item.description.lower()))
    return ordered[0], ordered[-1]


def _pc_summary_direction_points(
    bundle: pc.CurveBundle,
    direction: str,
    use_current_density: bool,
) -> list[tuple[float, float]]:
    files = bundle.asc_files if direction == "asc" else bundle.dsc_files
    if not files:
        return []

    rows = pc.concatenate_curve_data(files)
    last_rows = pc.find_last_point_of_each_step(rows, pc.infer_current_tolerance(files))
    if not last_rows:
        return []

    area_cm2 = pc._bundle_area_cm2(bundle) if use_current_density else None
    points = [
        (pc._scaled_current(row["Corriente"], use_current_density, area_cm2), row["Voltaje"])
        for row in last_rows
    ]
    points.sort(key=lambda item: item[0])
    return points


def _unique_sorted_points(points: list[tuple[float, float]]) -> list[tuple[float, float]]:
    grouped: dict[float, list[float]] = defaultdict(list)
    for x_value, y_value in points:
        grouped[x_value].append(y_value)
    return sorted((x_value, sum(y_values) / len(y_values)) for x_value, y_values in grouped.items())


def _interpolate_y(points: list[tuple[float, float]], x_value: float) -> float | None:
    if not points:
        return None
    if x_value < points[0][0] or x_value > points[-1][0]:
        return None

    for point_x, point_y in points:
        if point_x == x_value:
            return point_y

    for (x0, y0), (x1, y1) in zip(points, points[1:]):
        if x0 <= x_value <= x1 and x1 != x0:
            fraction = (x_value - x0) / (x1 - x0)
            return y0 + fraction * (y1 - y0)
    return None


def _pc_voltage_delta_points(
    first_points: list[tuple[float, float]],
    last_points: list[tuple[float, float]],
) -> list[tuple[float, float]]:
    first_points = _unique_sorted_points(first_points)
    last_points = _unique_sorted_points(last_points)
    if len(first_points) < 2 or len(last_points) < 2:
        return []

    x_min = max(first_points[0][0], last_points[0][0])
    x_max = min(first_points[-1][0], last_points[-1][0])
    if x_max < x_min:
        return []

    common_x = {x_value for x_value, _y_value in first_points + last_points if x_min <= x_value <= x_max}
    common_x.add(x_min)
    common_x.add(x_max)

    delta_points: list[tuple[float, float]] = []
    for x_value in sorted(common_x):
        first_y = _interpolate_y(first_points, x_value)
        last_y = _interpolate_y(last_points, x_value)
        if first_y is not None and last_y is not None:
            delta = last_y - first_y
            if delta > 0.0:
                delta_points.append((x_value, delta))
    return delta_points


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


def _draw_pc_first_last_delta_summary(
    bundles: list[pc.CurveBundle],
    font_defaults: PlotFontDefaults,
    language: str,
) -> Figure | None:
    pair = _pc_first_last_bundles(bundles)
    if pair is None:
        return None
    first_bundle, last_bundle = pair
    if first_bundle is last_bundle:
        return None

    use_current_density = _pc_current_axis_mode(bundles)
    first_label = _stage_label(first_bundle.curve_id, first_bundle.description, language)
    last_label = _stage_label(last_bundle.curve_id, last_bundle.description, language)
    fig = _new_plot_figure(figsize=(10.5, 6.2), dpi=150)
    ax = fig.add_subplot(111)
    all_x: list[float] = []
    all_y: list[float] = []

    direction_configs = [
        ("asc", translate("ascending", language), "^", pc.PC_PLOT_COLORS["asc_voltage"]),
        ("dsc", translate("descending", language), "v", pc.PC_PLOT_COLORS["dsc_voltage"]),
    ]
    for direction, direction_label, marker, color in direction_configs:
        first_points = _pc_summary_direction_points(first_bundle, direction, use_current_density)
        last_points = _pc_summary_direction_points(last_bundle, direction, use_current_density)
        delta_points = _pc_voltage_delta_points(first_points, last_points)
        if len(delta_points) < 2:
            continue

        x_values = [point[0] for point in delta_points]
        y_values = [point[1] for point in delta_points]
        all_x.extend(x_values)
        all_y.extend(y_values)
        ax.plot(
            x_values,
            y_values,
            marker=marker,
            linestyle="-",
            linewidth=1.8,
            markersize=5.5,
            markerfacecolor="none",
            markeredgewidth=1.1,
            color=color,
            label=f"{direction_label}: {last_label} - {first_label}",
        )

    if not all_x or not all_y:
        return None

    ax.axhline(0.0, color="#4a5568", linewidth=1.0, linestyle=":", alpha=0.75)
    ax.set_title(
        translate(
            "full_report_pc_delta_summary_title",
            language,
            first=first_label,
            last=last_label,
        )
    )
    ax.set_xlabel(
        f"{translate('current_density', language)} (A/cm^2)"
        if use_current_density
        else f"{translate('current', language)} (A)"
    )
    ax.set_ylabel(translate("full_report_pc_delta_summary_ylabel", language))
    ax.grid(True)
    ax.xaxis.set_major_locator(MaxNLocator(nbins=6))
    ax.yaxis.set_major_locator(MaxNLocator(nbins=6))
    ax.xaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
    ax.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
    ax.set_ylim(bottom=0.0)
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


def _eis_zero_voltage_group_label(entry: eis.EISPlotEntry, language: str) -> str:
    parts = [part.strip() for part in entry.display_name.split("/") if part.strip()]
    stage_prefixes = (f"{translate('stage', language)} ", "Etapa ", "Stage ")
    parts = [part for part in parts if not any(part.startswith(prefix) for prefix in stage_prefixes)]
    if parts:
        return " / ".join(parts)
    return entry.display_name or entry.voltage_label or "0V vs OCP"


def _eis_resistance_group_key(entry: eis.EISPlotEntry, language: str) -> tuple[float | str, str]:
    if entry.current_value is not None:
        return _eis_current_group_key(entry)
    if eis._is_zero_voltage_measurement(entry.parsed, entry.voltage_label):
        label = _eis_zero_voltage_group_label(entry, language)
        return label, label
    return _eis_current_group_key(entry)


def _is_activation_ocp_bar_label(label: object) -> bool:
    normalized = re.sub(r"[^a-z0-9]+", "_", str(label).strip().lower()).strip("_")
    return normalized == "eispot_0v_1"


def _eis_group_sort_key(key: tuple[float | str, str]) -> tuple[int, float, str]:
    if _is_activation_ocp_bar_label(key[1]):
        return -1, 0.0, key[1]
    if not isinstance(key[0], float) and (
        eis._is_zero_voltage_label(key[1]) or "0v vs ocp" in str(key[1]).strip().lower()
    ):
        return 0, 0.0, key[1]
    if isinstance(key[0], float):
        return 1, key[0], key[1]
    return 2, math.inf, key[1]


def _eis_current_groups(entries: list[eis.EISPlotEntry]) -> list[tuple[str, list[eis.EISPlotEntry]]]:
    grouped: dict[tuple[float | str, str], list[eis.EISPlotEntry]] = defaultdict(list)
    for entry in entries:
        if entry.current_value is None:
            continue
        grouped[_eis_current_group_key(entry)].append(entry)

    ordered_keys = sorted(grouped, key=_eis_group_sort_key)
    return [(key[1], grouped[key]) for key in ordered_keys]


def _eis_resistance_summary_groups(
    entries: list[eis.EISPlotEntry],
    language: str,
) -> list[tuple[str, list[eis.EISPlotEntry]]]:
    grouped: dict[tuple[float | str, str], list[eis.EISPlotEntry]] = defaultdict(list)
    for entry in entries:
        if entry.current_value is None and not eis._is_zero_voltage_measurement(entry.parsed, entry.voltage_label):
            continue
        grouped[_eis_resistance_group_key(entry, language)].append(entry)

    ordered_keys = sorted(grouped, key=_eis_group_sort_key)
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


def _nyquist_area_resistance_indicator(entry: eis.EISPlotEntry) -> tuple[float, str] | None:
    y0 = _nyquist_y0_indicator(entry)
    area_cm2 = eis._metadata_area_cm2(entry.parsed)
    if y0 is None or area_cm2 is None or area_cm2 <= 0:
        return None
    value, _unit = y0
    return value * area_cm2, "ohm.cm^2"


def _pc_high_current_resistance_indicator(bundle: pc.CurveBundle, language: str) -> tuple[float, str] | None:
    try:
        rows = pc.build_pc_dv_di_report_indicators(
            bundle,
            point_fraction=1.0,
            smoothing_algorithm="Median filter",
            smoothing_window=1,
            use_current_density=True,
            language=language,
        )
    except Exception:
        return None

    ascending_label = translate("ascending", language)
    target_label = translate("high_current_dvdi", language, direction=ascending_label)
    fallback: tuple[float, str] | None = None
    for label, value, unit in rows:
        numeric = _optional_float(value)
        if numeric is None:
            continue
        if str(label) == target_label:
            return numeric, "ohm.cm^2"
        if fallback is None and str(unit) == "V/(A/cm^2)" and "dV/dI" in str(label):
            fallback = (numeric, "ohm.cm^2")
    return fallback


def _eis_stage_key(entry: eis.EISPlotEntry) -> int | str:
    return entry.stage_number if entry.stage_number is not None else entry.display_name


def _draw_eis_y0_bar_summary(
    current_groups: list[tuple[str, list[eis.EISPlotEntry]]],
    pc_bundles: list[pc.CurveBundle],
    font_defaults: PlotFontDefaults,
    language: str,
) -> Figure | None:
    if not current_groups and not pc_bundles:
        return None

    stage_keys: list[int | str] = []
    stage_labels: dict[int | str, str] = {}
    stage_colors: dict[int | str, str] = {}
    group_labels: list[str] = []
    group_values: list[dict[int | str, float]] = []
    y_unit = ""

    for current_label, entries in current_groups:
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
            resistance = _nyquist_area_resistance_indicator(entry)
            if resistance is None:
                continue
            value, unit = resistance
            if not y_unit and unit:
                y_unit = unit
            stage_key = _eis_stage_key(entry)
            if stage_key not in stage_labels:
                stage_keys.append(stage_key)
                stage_labels[stage_key] = _stage_label(entry.stage_number, entry.display_name, language)
                if entry.nyquist_color:
                    stage_colors[stage_key] = entry.nyquist_color
            values_for_group[stage_key] = value
        if values_for_group:
            group_labels.append(current_label)
            group_values.append(values_for_group)

    pc_values: dict[int | str, float] = {}
    for bundle in sorted(pc_bundles, key=lambda item: (item.curve_id, item.description)):
        resistance = _pc_high_current_resistance_indicator(bundle, language)
        if resistance is None:
            continue
        value, unit = resistance
        if not y_unit and unit:
            y_unit = unit
        stage_key = bundle.curve_id
        if stage_key not in stage_labels:
            stage_keys.append(stage_key)
            stage_labels[stage_key] = _stage_label(bundle.curve_id, bundle.description, language)
        pc_values[stage_key] = value
    if pc_values:
        group_labels.append("PC")
        group_values.append(pc_values)

    if not stage_keys or not any(group_values):
        return None

    fig = _new_plot_figure(figsize=(10.0, 6.2), dpi=150)
    ax = fig.add_subplot(111)
    x_positions = list(range(len(group_values)))
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
    ax.set_xlabel("EIS / PC")
    y_label = translate("resistance", language)
    ax.set_ylabel(f"{y_label} ({y_unit})" if y_unit else y_label)
    ax.set_xticks(x_positions)
    ax.set_xticklabels(group_labels, rotation=20, ha="right")
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
    title_detail: str | None = None,
    show_temperature: bool = True,
    stage_numbers: list[int] | None = None,
    show_fit_line: bool = False,
    fit_use_linear: bool = False,
    use_stage_colors: bool = True,
) -> dict[str, object]:
    limits = deg.compute_autofit_v_vs_t_limits(parsed_items, show_temperature=show_temperature, time_unit="h")
    plot_title = translate(title_key, language)
    if title_detail:
        plot_title = f"{plot_title} - {title_detail}"
    return {
        "show_temperature": show_temperature,
        "voltage_linestyle": "-",
        "temperature_linestyle": "--",
        "time_unit": "h",
        "x_tick_count": 6,
        "y_tick_count": 6,
        "plot_title": plot_title,
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
        "fit_use_linear": fit_use_linear,
        "show_fit_line": show_fit_line,
        "show_fit_range": False,
        "stage_numbers": stage_numbers,
        "use_stage_colors": use_stage_colors,
        "language": language,
    }


def _deg_stage_range_label(parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]], language: str) -> str:
    stages = [item[0].stage for item in _sorted_deg_items(parsed_items)]
    if not stages:
        return ""
    stage_word = translate("stage", language)
    if len(stages) == 1:
        return f"{stage_word} {stages[0]}"
    return f"{stage_word} {stages[0]}-{stages[-1]}"


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


def _deg_series_segments(
    parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
) -> tuple[list[tuple[deg.DegFile, list[float], list[float]]], bool]:
    raw_segments: list[tuple[deg.DegFile, list[tuple[float, float]], datetime | None]] = []
    can_use_absolute_time = True

    for deg_file, parsed in _sorted_deg_items(parsed_items):
        try:
            time_values, voltage_values = deg._required_numeric_series(parsed, "T", "Vf")
        except ValueError:
            continue

        pairs = sorted(zip(time_values, voltage_values), key=lambda item: item[0])
        if len(pairs) < 2:
            continue

        try:
            stage_start = deg._start_datetime(parsed, deg_file.path.name)
        except ValueError:
            stage_start = None
            can_use_absolute_time = False

        raw_segments.append((deg_file, pairs, stage_start))

    if not raw_segments:
        return [], False

    if can_use_absolute_time and all(stage_start is not None for _deg_file, _pairs, stage_start in raw_segments):
        absolute_segments: list[tuple[deg.DegFile, list[tuple[datetime, float]]]] = []
        for deg_file, pairs, stage_start in raw_segments:
            if stage_start is None:
                continue
            absolute_segments.append(
                (
                    deg_file,
                    [(stage_start + timedelta(seconds=time_value), voltage) for time_value, voltage in pairs],
                )
            )

        reference_start = min(
            absolute_time
            for _deg_file, absolute_points in absolute_segments
            for absolute_time, _voltage in absolute_points
        )
        return [
            (
                deg_file,
                [(absolute_time - reference_start).total_seconds() / deg.SECONDS_PER_HOUR for absolute_time, _voltage in absolute_points],
                [voltage for _absolute_time, voltage in absolute_points],
            )
            for deg_file, absolute_points in absolute_segments
        ], True

    cumulative_seconds = 0.0
    segments: list[tuple[deg.DegFile, list[float], list[float]]] = []
    for deg_file, pairs, _stage_start in raw_segments:
        stage_start_time = pairs[0][0]
        stage_end_time = pairs[-1][0]
        stage_duration = max(0.0, stage_end_time - stage_start_time)
        segments.append(
            (
                deg_file,
                [(time_value - stage_start_time + cumulative_seconds) / deg.SECONDS_PER_HOUR for time_value, _voltage in pairs],
                [voltage for _time_value, voltage in pairs],
            )
        )
        cumulative_seconds += stage_duration

    return segments, False


def _draw_deg_series_summary(
    parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
    font_defaults: PlotFontDefaults,
    language: str,
) -> Figure | None:
    segments, _uses_absolute_time = _deg_series_segments(parsed_items)
    if not segments:
        return None

    fig = _new_plot_figure(figsize=(10.5, 6.2), dpi=150)
    ax = fig.add_subplot(111)
    stage_numbers = [deg_file.stage for deg_file, _x_values, _y_values in segments]
    all_x: list[float] = []
    all_y: list[float] = []

    for index, (deg_file, x_values, y_values) in enumerate(segments):
        if not x_values or not y_values:
            continue
        all_x.extend(x_values)
        all_y.extend(y_values)
        color = deg._stage_plot_color(deg_file, stage_numbers, index, len(segments))
        ax.plot(
            x_values,
            y_values,
            color=color,
            linewidth=1.5,
            linestyle="-",
            label=_stage_label(deg_file.stage, deg_file.path.stem, language),
        )

    if not all_x or not all_y:
        return None

    rows = _deg_simple_slope_rows(parsed_items, language)
    if rows:
        label, value, unit = rows[0]
        ax.text(
            0.015,
            0.985,
            f"{label}: {value} {unit}".strip(),
            transform=ax.transAxes,
            fontsize=max(7.0, font_defaults.legend * 0.9),
            va="top",
            ha="left",
            bbox=dict(boxstyle="round,pad=0.25", facecolor="white", edgecolor="#cbd5e0", alpha=0.88),
        )

    ax.set_title(translate("full_report_deg_series_summary_title", language))
    ax.set_xlabel(f"{translate('time', language)} [h]")
    ax.set_ylabel(f"{translate('voltage', language)} [V]")
    ax.grid(True, alpha=0.25)
    ax.xaxis.set_major_locator(MaxNLocator(nbins=6))
    ax.yaxis.set_major_locator(MaxNLocator(nbins=6))
    ax.xaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
    ax.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
    if min(all_x) != max(all_x):
        ax.set_xlim(min(all_x), max(all_x))
    voltage_decimals = deg._auto_decimals(all_y)
    v_min, v_max = deg._tight_limits(all_y, decimals=voltage_decimals)
    if v_min is not None and v_max is not None:
        ax.set_ylim(v_min, v_max)
    legend = ax.legend(fontsize=max(6.0, font_defaults.legend * 0.85), loc="best")
    make_legend_draggable(legend)
    apply_plot_font_defaults(fig, font_defaults)
    fig.tight_layout()
    return fig


def _deg_consecutive_segments(
    parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
) -> list[tuple[deg.DegFile, list[float], list[float], list[float], list[float]]]:
    segments: list[tuple[deg.DegFile, list[float], list[float], list[float], list[float]]] = []
    cumulative_seconds = 0.0

    for deg_file, parsed in _sorted_deg_items(parsed_items):
        try:
            time_values, voltage_values = deg._required_numeric_series(parsed, "T", "Vf")
        except ValueError:
            continue

        voltage_pairs = sorted(zip(time_values, voltage_values), key=lambda item: item[0])
        if len(voltage_pairs) < 2:
            continue

        try:
            temp_time_values, temperature_values = deg._required_numeric_series(parsed, "T", "Temp")
            temperature_pairs = sorted(zip(temp_time_values, temperature_values), key=lambda item: item[0])
        except ValueError:
            temperature_pairs = []

        stage_start = min([voltage_pairs[0][0], *[point[0] for point in temperature_pairs[:1]]])
        stage_end = max([voltage_pairs[-1][0], *[point[0] for point in temperature_pairs[-1:]]])
        stage_duration = max(0.0, stage_end - stage_start)
        segments.append(
            (
                deg_file,
                [(time_value - stage_start + cumulative_seconds) / deg.SECONDS_PER_HOUR for time_value, _voltage in voltage_pairs],
                [voltage for _time_value, voltage in voltage_pairs],
                [(time_value - stage_start + cumulative_seconds) / deg.SECONDS_PER_HOUR for time_value, _temperature in temperature_pairs],
                [temperature for _time_value, temperature in temperature_pairs],
            )
        )
        cumulative_seconds += stage_duration

    return segments


def _draw_deg_consecutive_summary(
    parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
    font_defaults: PlotFontDefaults,
    language: str,
) -> Figure | None:
    segments = _deg_consecutive_segments(parsed_items)
    if not segments:
        return None

    fig = _new_plot_figure(figsize=(10.5, 7.2), dpi=150)
    ax_voltage = fig.add_subplot(211)
    ax_temp = fig.add_subplot(212, sharex=ax_voltage)
    stage_numbers = [deg_file.stage for deg_file, _x_values, _y_values, _temp_x_values, _temp_values in segments]
    all_x: list[float] = []
    all_voltage: list[float] = []
    all_temperature: list[float] = []

    for index, (deg_file, x_values, voltage_values, temp_x_values, temperature_values) in enumerate(segments):
        if not x_values or not voltage_values:
            continue

        if index > 0:
            for axis in (ax_voltage, ax_temp):
                axis.axvline(x_values[0], color="#718096", linewidth=0.8, linestyle=":", alpha=0.45, label="_nolegend_")

        all_x.extend(x_values)
        all_voltage.extend(voltage_values)
        color = deg._stage_plot_color(deg_file, stage_numbers, index, len(segments))
        stage_label = _stage_label(deg_file.stage, deg_file.path.stem, language)
        ax_voltage.plot(
            x_values,
            voltage_values,
            color=color,
            linewidth=1.5,
            linestyle="-",
            label=f"{stage_label} V",
        )

        if temp_x_values and temperature_values:
            all_x.extend(temp_x_values)
            all_temperature.extend(temperature_values)
            ax_temp.plot(
                temp_x_values,
                temperature_values,
                color=color,
                linewidth=1.3,
                linestyle="-",
                alpha=0.92,
                label=f"{stage_label} T",
            )

    if not all_x or not all_voltage:
        return None

    ax_voltage.set_title(translate("full_report_deg_consecutive_summary_title", language))
    ax_voltage.set_ylabel(f"{translate('voltage', language)} [V]")
    ax_temp.set_ylabel(f"{translate('temperature', language)} [C]")
    ax_temp.set_xlabel(f"{translate('time', language)} [h]")
    ax_voltage.grid(True, alpha=0.25)
    ax_temp.grid(True, alpha=0.25)
    ax_voltage.tick_params(axis="x", labelbottom=False)
    ax_temp.xaxis.set_major_locator(MaxNLocator(nbins=6))
    ax_voltage.yaxis.set_major_locator(MaxNLocator(nbins=6))
    ax_temp.yaxis.set_major_locator(MaxNLocator(nbins=6))
    ax_temp.xaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
    ax_voltage.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
    ax_temp.yaxis.set_major_formatter(StrMethodFormatter("{x:g}"))
    if min(all_x) != max(all_x):
        ax_voltage.set_xlim(min(all_x), max(all_x))

    voltage_decimals = deg._auto_decimals(all_voltage)
    v_min, v_max = deg._tight_limits(all_voltage, decimals=voltage_decimals)
    if v_min is not None and v_max is not None:
        ax_voltage.set_ylim(v_min, v_max)
    if all_temperature:
        temp_decimals = deg._auto_decimals(all_temperature)
        temp_min, temp_max = deg._tight_limits(all_temperature, decimals=temp_decimals)
        if temp_min is not None and temp_max is not None:
            ax_temp.set_ylim(temp_min, temp_max)

    voltage_handles, voltage_labels = ax_voltage.get_legend_handles_labels()
    if voltage_handles:
        legend_columns = 2 if len(voltage_handles) > 6 else 1
        make_legend_draggable(
            ax_voltage.legend(
                voltage_handles,
                voltage_labels,
                fontsize=max(6.0, font_defaults.legend * 0.78),
                loc="best",
                ncol=legend_columns,
            )
        )
    temp_handles, temp_labels = ax_temp.get_legend_handles_labels()
    if temp_handles:
        legend_columns = 2 if len(temp_handles) > 6 else 1
        make_legend_draggable(
            ax_temp.legend(
                temp_handles,
                temp_labels,
                fontsize=max(6.0, font_defaults.legend * 0.78),
                loc="best",
                ncol=legend_columns,
            )
        )
    apply_plot_font_defaults(fig, font_defaults)
    fig.tight_layout()
    return fig


def _draw_deg_summary(
    parsed_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
    font_defaults: PlotFontDefaults,
    language: str,
) -> Figure | None:
    parsed_items = _sorted_deg_items(parsed_items)
    if not parsed_items:
        return None
    stage_numbers = [deg_file.stage for deg_file, _parsed in parsed_items]

    fig = _new_plot_figure(figsize=(10.0, 6.4), dpi=150)
    has_plot = deg.draw_v_vs_t_on_figure(
        fig=fig,
        parsed_items=parsed_items,
        **_deg_v_vs_t_kwargs(
            parsed_items,
            font_defaults,
            language,
            title_key="full_report_deg_summary_title",
            title_detail=_deg_stage_range_label(parsed_items, language),
            show_temperature=False,
            stage_numbers=stage_numbers,
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


def _relevant_summary_indicator_rows(
    activation_bundles: list[activ.ActivationBundle],
    deg_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
    language: str,
) -> list[tuple[object, object, object]]:
    rows: list[tuple[object, object, object]] = []

    for bundle in sorted(activation_bundles, key=lambda item: item.label):
        for label, value, unit in activ.build_activation_report_indicators(bundle, language=language):
            if str(label).startswith("dV stab C#"):
                rows.append((f"Activacion - {bundle.label} - {label}", value, unit))

    slope_label = translate("deg_average_slope", language)
    for deg_file, parsed in _sorted_deg_items(deg_items):
        stage_label = _stage_label(deg_file.stage, deg_file.path.stem, language)
        for label, value, unit in deg.build_deg_report_indicators([(deg_file, parsed)], language=language):
            if str(label) == slope_label:
                rows.append((f"Deg - {stage_label} - {label}", value, unit))
                break

    return rows


def _relevant_summary_indicator_page_count(rows: list[tuple[object, object, object]]) -> int:
    if not rows:
        return 0
    return math.ceil(len(rows) / RELEVANT_SUMMARY_ROWS_PER_PAGE)


def _relevant_summary_indicator_pages(
    rows: list[tuple[object, object, object]],
    language: str,
) -> list[Figure]:
    page_count = _relevant_summary_indicator_page_count(rows)
    pages: list[Figure] = []
    title = translate("full_report_relevant_indicators_title", language)
    subtitle = translate("full_report_relevant_indicators_subtitle", language)
    for page_index in range(page_count):
        start = page_index * RELEVANT_SUMMARY_ROWS_PER_PAGE
        chunk = rows[start : start + RELEVANT_SUMMARY_ROWS_PER_PAGE]
        page_title = title if page_count == 1 else f"{title} ({page_index + 1}/{page_count})"
        pages.append(
            _table_page(
                page_title,
                subtitle,
                [(translate("indicators", language), chunk, [0.05, 0.08, 0.90, 0.78])],
                language,
            )
        )
    return pages


def _sorted_deg_ocp_items(parsed_items: list[tuple[deg.DegFile, ocp.ParsedDTA]]) -> list[tuple[deg.DegFile, ocp.ParsedDTA]]:
    return sorted(parsed_items, key=lambda item: (item[0].stage, item[0].path.name.lower()))


def _deg_ocp_plot_kwargs(
    deg_file: deg.DegFile,
    parsed: ocp.ParsedDTA,
    font_defaults: PlotFontDefaults,
    language: str,
) -> dict[str, object]:
    limits = ocp.compute_autofit_v_vs_t_limits(parsed, show_temperature=True)
    stage_label = _stage_label(deg_file.stage, deg_file.path.stem, language)
    return {
        "show_temperature": True,
        "voltage_linestyle": "-",
        "temperature_linestyle": "--",
        "tick_count": 6,
        "plot_title": f"Deg OCP - {stage_label}",
        "show_title": True,
        "title_fontsize": font_defaults.title,
        "tick_fontsize": font_defaults.tick,
        "label_fontsize": font_defaults.label,
        "legend_fontsize": max(6.0, font_defaults.legend * 0.9),
        "line_width": 1.5,
        "t_min": _optional_float(limits.get("t_min")),
        "t_max": _optional_float(limits.get("t_max")),
        "v_min": _optional_float(limits.get("v_min")),
        "v_max": _optional_float(limits.get("v_max")),
        "temp_min": _optional_float(limits.get("temp_min")),
        "temp_max": _optional_float(limits.get("temp_max")),
        "language": language,
    }


def _pc_v_vs_i_kwargs(bundle: pc.CurveBundle, font_defaults: PlotFontDefaults, language: str) -> dict[str, object]:
    use_density = _pc_bundle_use_density(bundle)
    limits = pc.compute_default_v_vs_i_limits(bundle, use_current_density=use_density)
    curve_label = _pc_curve_label(bundle, language)
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
        "plot_title": f"{translate('v_vs_i', language)} - {curve_label}",
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
    curve_label = _pc_curve_label(bundle, language)
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
        "plot_title": f"{translate('series_by_time', language)} - {curve_label}",
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
    curve_label = _pc_curve_label(bundle, language)
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
        "plot_title": f"{translate('dv_di', language)} - {curve_label}",
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
    outline_parent: PdfOutlineEntry | None = None,
) -> None:
    for bundle in bundles:
        curve_label = _pc_curve_label(bundle, language)
        use_density = _pc_bundle_use_density(bundle)
        start_page = step[0]
        outline_entry = _bookmark_child(
            outline_parent,
            _stage_bookmark_label("PC", bundle.curve_id, bundle.description, language),
            start_page,
        )
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
                    "plot_title": f"{curve_label} - {translate('voltage_step_range', language)}",
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
                    "plot_title": f"{curve_label} - {translate('temperature_step_range', language)}",
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

        if step[0] == start_page:
            _remove_bookmark_child(outline_parent, outline_entry)


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


def _eis_measurement_key(entry: eis.EISPlotEntry) -> tuple[int | str, float | str] | None:
    if entry.current_value is None:
        return None
    stage_key: int | str = entry.stage_number if entry.stage_number is not None else entry.display_name
    return stage_key, round(entry.current_value, 12)


def _eis_stage_sort_key(stage_key: int | str) -> tuple[int, float, str]:
    if isinstance(stage_key, int):
        return (0, float(stage_key), "")
    return (1, math.inf, str(stage_key).lower())


def _eis_stage_bookmark_label(
    stage_key: int | str,
    stage_entries: list[eis.EISPlotEntry],
    stage_pre_entries: list[eis.EISPlotEntry],
    language: str,
) -> str:
    for entry in [*stage_entries, *stage_pre_entries]:
        if entry.stage_number is not None:
            return _stage_bookmark_label("EIS", entry.stage_number, entry.display_name, language)
    fallback = stage_entries[0].display_name if stage_entries else None
    if fallback is None and stage_pre_entries:
        fallback = stage_pre_entries[0].display_name
    return _stage_bookmark_label("EIS", None, fallback or str(stage_key), language)


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
    outline_parent: PdfOutlineEntry | None = None,
) -> None:
    pre_by_key: dict[tuple[int | str, float | str], list[eis.EISPlotEntry]] = defaultdict(list)
    for pre_entry in pre_entries:
        key = _eis_measurement_key(pre_entry)
        if key is not None:
            pre_by_key[key].append(pre_entry)
    emitted_pre_ids: set[int] = set()
    eis_by_stage: dict[int | str, list[eis.EISPlotEntry]] = defaultdict(list)
    pre_by_stage: dict[int | str, list[eis.EISPlotEntry]] = defaultdict(list)

    for entry in entries:
        eis_by_stage[_eis_stage_key(entry)].append(entry)
    for pre_entry in pre_entries:
        pre_by_stage[_eis_stage_key(pre_entry)].append(pre_entry)

    def _append_pre_entry(entry: eis.EISPlotEntry) -> None:
        emitted_pre_ids.add(id(entry))
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

    def _append_eis_entry(entry: eis.EISPlotEntry) -> None:
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

            series_indicator_rows = eis.build_series_pt_indicator_rows(entry.parsed)
            if series_indicator_rows:
                series_table_fig = _table_page(
                    translate("eis_series_pt_report_title", language, curve=entry.display_name),
                    "",
                    [
                        (
                            translate("metadata", language),
                            eis._build_eis_report_metadata_rows(entry.parsed, language),
                            [0.05, 0.53, 0.90, 0.34],
                        ),
                        (
                            translate("series_pt_indicators", language),
                            series_indicator_rows,
                            [0.05, 0.08, 0.90, 0.36],
                        ),
                    ],
                    language,
                )
                _save_figure(pdf, series_table_fig)
                step[0] += 1
        except Exception as exc:
            warnings.append(f"EIS {entry.display_name}: {type(exc).__name__}: {exc}")

    stage_keys = sorted(set(eis_by_stage) | set(pre_by_stage), key=_eis_stage_sort_key)
    for stage_key in stage_keys:
        stage_entries = eis_by_stage.get(stage_key, [])
        stage_pre_entries = pre_by_stage.get(stage_key, [])
        start_page = step[0]
        outline_entry = _bookmark_child(
            outline_parent,
            _eis_stage_bookmark_label(stage_key, stage_entries, stage_pre_entries, language),
            start_page,
        )

        for entry in stage_entries:
            key = _eis_measurement_key(entry)
            if key is not None:
                for pre_entry in pre_by_key.get(key, []):
                    if _eis_stage_key(pre_entry) == stage_key and id(pre_entry) not in emitted_pre_ids:
                        _append_pre_entry(pre_entry)
            _append_eis_entry(entry)

        for pre_entry in stage_pre_entries:
            if id(pre_entry) not in emitted_pre_ids:
                _append_pre_entry(pre_entry)

        if step[0] == start_page:
            _remove_bookmark_child(outline_parent, outline_entry)


def _append_cv_individual_pages(
    pdf: PdfPages,
    datasets: list[cv.CVDataset],
    font_defaults: PlotFontDefaults,
    language: str,
    warnings: list[str],
    progress_callback: ProgressCallback | None,
    step: list[int],
    total: int,
    outline_parent: PdfOutlineEntry | None = None,
) -> None:
    for dataset in datasets:
        label = cv._dataset_stage_label(dataset)
        _emit(progress_callback, f"CV: {label}", step[0], total)
        start_page = step[0]
        outline_entry: PdfOutlineEntry | None = None
        try:
            visible_segment_keys = {segment.key for segment in dataset.segments}
            report_segment_keys = {
                segment.key
                for segment in dataset.segments
                if segment.key in visible_segment_keys and segment.cycle >= 2
            }
            if not report_segment_keys:
                continue

            outline_entry = _bookmark_child(
                outline_parent,
                _stage_bookmark_label("CV", dataset.stage_number, dataset.display_name, language),
                start_page,
            )
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
        if step[0] == start_page:
            _remove_bookmark_child(outline_parent, outline_entry)


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
    ocp_items: list[tuple[deg.DegFile, ocp.ParsedDTA]],
    font_defaults: PlotFontDefaults,
    language: str,
    warnings: list[str],
    progress_callback: ProgressCallback | None,
    step: list[int],
    total: int,
    outline_parent: PdfOutlineEntry | None = None,
) -> None:
    parsed_items = _sorted_deg_items(parsed_items)
    ocp_items = _sorted_deg_ocp_items(ocp_items)
    stage_numbers = [deg_file.stage for deg_file, _parsed in parsed_items]

    for deg_file, parsed in parsed_items:
        single_items = [(deg_file, parsed)]
        stage_label = _stage_label(deg_file.stage, deg_file.path.stem, language)
        source_label = f"{translate('deg_report_title', language)} - {stage_label}"
        _emit(progress_callback, source_label, step[0], total)
        start_page = step[0]
        outline_entry = _bookmark_child(
            outline_parent,
            _stage_bookmark_label("Deg Galv", deg_file.stage, deg_file.path.stem, language),
            start_page,
        )
        try:
            plot_fig = _new_plot_figure(figsize=(10.0, 6.4), dpi=150)
            if deg.draw_v_vs_t_on_figure(
                fig=plot_fig,
                parsed_items=single_items,
                **_deg_v_vs_t_kwargs(
                    single_items,
                    font_defaults,
                    language,
                    title_key="deg_report_plot_title",
                    title_detail=stage_label,
                    stage_numbers=stage_numbers,
                    show_fit_line=True,
                    fit_use_linear=True,
                    use_stage_colors=False,
                ),
            ):
                _save_figure(pdf, plot_fig)
                step[0] += 1

            table_fig = _table_page(
                source_label,
                translate("deg_report_subtitle", language),
                [
                    (
                        translate("metadata", language),
                        deg.build_deg_report_metadata(single_items, language=language),
                        [0.05, 0.53, 0.90, 0.34],
                    ),
                    (
                        translate("indicators", language),
                        deg.build_deg_report_indicators(single_items, language=language),
                        [0.05, 0.14, 0.90, 0.29],
                    ),
                ],
                language,
            )
            _save_figure(pdf, table_fig)
            step[0] += 1
        except Exception as exc:
            warnings.append(f"{source_label}: {type(exc).__name__}: {exc}")
        if step[0] == start_page:
            _remove_bookmark_child(outline_parent, outline_entry)

    for deg_file, parsed in ocp_items:
        stage_label = _stage_label(deg_file.stage, deg_file.path.stem, language)
        source_label = f"Deg OCP - {stage_label}"
        _emit(progress_callback, source_label, step[0], total)
        start_page = step[0]
        outline_entry = _bookmark_child(
            outline_parent,
            _stage_bookmark_label("Deg OCP", deg_file.stage, deg_file.path.stem, language),
            start_page,
        )
        try:
            plot_fig = _new_plot_figure(figsize=(10.0, 6.4), dpi=150)
            if ocp.draw_v_vs_t_on_figure(
                plot_fig,
                parsed=parsed,
                source_name=source_label,
                **_deg_ocp_plot_kwargs(deg_file, parsed, font_defaults, language),
            ):
                _save_figure(pdf, plot_fig)
                step[0] += 1

            table_fig = _table_page(
                translate("ocp_v_vs_t_report_title", language, file=source_label),
                translate("ocp_v_vs_t_report_subtitle", language),
                [
                    (
                        translate("metadata", language),
                        ocp.build_ocp_report_metadata(parsed, source_label, language=language),
                        [0.05, 0.53, 0.90, 0.34],
                    ),
                    (
                        translate("indicators", language),
                        ocp.build_ocp_v_vs_t_report_indicators(parsed, language=language),
                        [0.05, 0.08, 0.90, 0.36],
                    ),
                ],
                language,
            )
            _save_figure(pdf, table_fig)
            step[0] += 1
        except Exception as exc:
            warnings.append(f"{source_label}: {type(exc).__name__}: {exc}")
        if step[0] == start_page:
            _remove_bookmark_child(outline_parent, outline_entry)


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
    resistance_groups: list[tuple[str, list[eis.EISPlotEntry]]],
    deg_items: list[tuple[deg.DegFile, deg.ParsedDTA]],
    deg_ocp_items: list[tuple[deg.DegFile, ocp.ParsedDTA]],
    relevant_indicator_rows: list[tuple[object, object, object]],
) -> int:
    total = 2
    if activation_bundles or pc_bundles or resistance_groups or pre_entries or cv_datasets or deg_items or deg_ocp_items:
        total += 1
    total += len(activation_bundles)
    if pc_bundles:
        total += 1
    if len(pc_bundles) >= 2:
        total += 1
    total += len(current_groups)
    if resistance_groups or pc_bundles:
        total += 1
    if deg_items:
        total += 1
        total += 1
        total += 1
    total += _relevant_summary_indicator_page_count(relevant_indicator_rows)
    if activation_bundles:
        total += 1 + (3 * len(activation_bundles))
    if pc_bundles:
        total += 1 + (10 * len(pc_bundles))
    if eis_entries or pre_entries:
        total += 1 + (5 * len(eis_entries)) + (2 * len(pre_entries))
    if cv_datasets:
        total += 1 + (2 * len(cv_datasets))
    if deg_items or deg_ocp_items:
        total += 1
    if deg_items:
        total += 2 * len(deg_items)
    if deg_ocp_items:
        total += 2 * len(deg_ocp_items)
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
    deg_ocp_files = deg.find_deg_ocp_files(input_dir)
    deg_ocp_items = _sorted_deg_ocp_items(
        [(deg_file, ocp.parse_gamry_dta(deg_file.path)) for deg_file in deg_ocp_files]
    )
    current_groups = _eis_current_groups(eis_entries)
    resistance_groups = _eis_resistance_summary_groups(eis_entries, language)
    relevant_indicator_rows = _relevant_summary_indicator_rows(activation_bundles, deg_items, language)

    if not (activation_bundles or pc_bundles or eis_entries or pre_entries or cv_datasets or deg_items or deg_ocp_items):
        raise ValueError(translate("full_report_no_data", language))

    counts = {
        "activation": len(activation_bundles),
        "pc": len(pc_bundles),
        "eis": len(eis_entries),
        "pre_stab": len(pre_entries),
        "cv": len(cv_datasets),
        "deg": len(deg_items),
        "deg_ocp": len(deg_ocp_items),
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
        resistance_groups,
        deg_items,
        deg_ocp_items,
        relevant_indicator_rows,
    )
    step = [0]
    warnings: list[str] = []
    section_titles: list[str] = []
    if activation_bundles or pc_bundles or resistance_groups or pre_entries or cv_datasets or deg_items or deg_ocp_items:
        section_titles.append(translate("full_report_summary", language))
    if activation_bundles:
        section_titles.append("Activacion")
    if pc_bundles:
        section_titles.append("PC")
    if eis_entries or pre_entries:
        section_titles.append("EIS")
    if cv_datasets:
        section_titles.append("CV")
    if deg_items or deg_ocp_items:
        section_titles.append("Deg")
    section_entries: list[PdfOutlineEntry] = []
    summary_parent: PdfOutlineEntry | None = None

    with PdfPages(output_path) as pdf:
        _emit(progress_callback, translate("full_report_generating", language), step[0], total)
        _save_figure(pdf, _title_page(input_dir, output_dir, counts, language))
        _bookmark(section_entries, translate("full_report_title", language), step[0])
        step[0] += 1
        _save_figure(pdf, _index_page(section_titles, language))
        _bookmark(section_entries, translate("full_report_index", language), step[0])
        step[0] += 1

        if activation_bundles or pc_bundles or resistance_groups or pre_entries or cv_datasets or deg_items or deg_ocp_items:
            summary_parent = _bookmark(section_entries, translate("full_report_summary", language), step[0])
            _save_figure(
                pdf,
                _section_page(
                    translate("full_report_summary", language),
                    translate("full_report_summary_subtitle", language),
                    language,
                ),
            )
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
                _bookmark_child(summary_parent, title, step[0])
                _save_figure(pdf, activation_summary_fig)
                step[0] += 1

        if pc_bundles:
            title = translate("full_report_pc_summary_title", language)
            _emit(progress_callback, title, step[0], total)
            pc_summary_fig = _draw_pc_ascending_summary(pc_bundles, font_defaults, language)
            if pc_summary_fig is not None:
                _bookmark_child(summary_parent, title, step[0])
                _save_figure(pdf, pc_summary_fig)
                step[0] += 1

            pc_delta_pair = _pc_first_last_bundles(pc_bundles)
            if pc_delta_pair is not None:
                first_bundle, last_bundle = pc_delta_pair
                first_label = _stage_label(first_bundle.curve_id, first_bundle.description, language)
                last_label = _stage_label(last_bundle.curve_id, last_bundle.description, language)
                title = translate(
                    "full_report_pc_delta_summary_title",
                    language,
                    first=first_label,
                    last=last_label,
                )
                _emit(progress_callback, title, step[0], total)
                pc_delta_fig = _draw_pc_first_last_delta_summary(pc_bundles, font_defaults, language)
                if pc_delta_fig is not None:
                    _bookmark_child(summary_parent, title, step[0])
                    _save_figure(pdf, pc_delta_fig)
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
                _bookmark_child(summary_parent, title, step[0])
                _save_figure(pdf, summary_fig)
                step[0] += 1

        if resistance_groups or pc_bundles:
            title = translate("full_report_eis_y0_summary_title", language)
            _emit(progress_callback, title, step[0], total)
            y0_summary_fig = _draw_eis_y0_bar_summary(resistance_groups, pc_bundles, font_defaults, language)
            if y0_summary_fig is not None:
                _bookmark_child(summary_parent, title, step[0])
                _save_figure(pdf, y0_summary_fig)
                step[0] += 1

        if deg_items:
            title = translate("full_report_deg_summary_title", language)
            _emit(progress_callback, title, step[0], total)
            deg_summary_fig = _draw_deg_summary(deg_items, font_defaults, language)
            if deg_summary_fig is not None:
                _bookmark_child(summary_parent, title, step[0])
                _save_figure(pdf, deg_summary_fig)
                step[0] += 1

            title = translate("full_report_deg_series_summary_title", language)
            _emit(progress_callback, title, step[0], total)
            deg_series_summary_fig = _draw_deg_series_summary(deg_items, font_defaults, language)
            if deg_series_summary_fig is not None:
                _bookmark_child(summary_parent, title, step[0])
                _save_figure(pdf, deg_series_summary_fig)
                step[0] += 1

            title = translate("full_report_deg_consecutive_summary_title", language)
            _emit(progress_callback, title, step[0], total)
            deg_consecutive_summary_fig = _draw_deg_consecutive_summary(deg_items, font_defaults, language)
            if deg_consecutive_summary_fig is not None:
                _bookmark_child(summary_parent, title, step[0])
                _save_figure(pdf, deg_consecutive_summary_fig)
                step[0] += 1

        if relevant_indicator_rows:
            title = translate("full_report_relevant_indicators_title", language)
            for page_index, indicator_fig in enumerate(_relevant_summary_indicator_pages(relevant_indicator_rows, language)):
                _emit(progress_callback, title, step[0], total)
                if page_index == 0:
                    _bookmark_child(summary_parent, title, step[0])
                _save_figure(pdf, indicator_fig)
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
            pc_parent = _bookmark(section_entries, "PC", step[0])
            _save_figure(
                pdf,
                _section_page(
                    "PC",
                    translate("full_report_individual_subtitle", language),
                    language,
                ),
            )
            step[0] += 1
            _append_pc_individual_pages(
                pdf,
                pc_bundles,
                font_defaults,
                language,
                warnings,
                progress_callback,
                step,
                total,
                pc_parent,
            )

        if eis_entries or pre_entries:
            eis_parent = _bookmark(section_entries, "EIS", step[0])
            _save_figure(
                pdf,
                _section_page(
                    "EIS",
                    translate("full_report_individual_subtitle", language),
                    language,
                ),
            )
            step[0] += 1
            _append_eis_individual_pages(
                pdf,
                eis_entries,
                pre_entries,
                font_defaults,
                language,
                warnings,
                progress_callback,
                step,
                total,
                eis_parent,
            )

        if cv_datasets:
            cv_parent = _bookmark(section_entries, "CV", step[0])
            _save_figure(
                pdf,
                _section_page(
                    "CV",
                    translate("full_report_individual_subtitle", language),
                    language,
                ),
            )
            step[0] += 1
            _append_cv_individual_pages(
                pdf,
                cv_datasets,
                font_defaults,
                language,
                warnings,
                progress_callback,
                step,
                total,
                cv_parent,
            )

        if deg_items or deg_ocp_items:
            deg_parent = _bookmark(section_entries, "Deg", step[0])
            _save_figure(
                pdf,
                _section_page(
                    "Deg",
                    translate("full_report_individual_subtitle", language),
                    language,
                ),
            )
            step[0] += 1
            _append_deg_individual_pages(
                pdf,
                deg_items,
                deg_ocp_items,
                font_defaults,
                language,
                warnings,
                progress_callback,
                step,
                total,
                deg_parent,
            )

        _append_warnings_page(pdf, warnings, language)

    _add_pdf_outline(output_path, section_entries)
    _emit(progress_callback, translate("full_report_completed", language), total, total)
    return output_path
