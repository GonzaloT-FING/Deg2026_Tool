from __future__ import annotations

from dataclasses import dataclass

from matplotlib.figure import Figure


def _format_font_value(value: float) -> str:
    return f"{float(value):g}"


def _parse_positive_float(value: str | float, label: str) -> float:
    try:
        parsed = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be a positive number.") from exc
    if parsed <= 0:
        raise ValueError(f"{label} must be a positive number.")
    return parsed


@dataclass(frozen=True)
class PlotFontDefaults:
    title: float = 14.0
    tick: float = 10.0
    label: float = 11.0
    legend: float = 10.0

    def as_strings(self) -> dict[str, str]:
        return {
            "title": _format_font_value(self.title),
            "tick": _format_font_value(self.tick),
            "label": _format_font_value(self.label),
            "legend": _format_font_value(self.legend),
        }


DEFAULT_PLOT_FONT_DEFAULTS = PlotFontDefaults()


def resolve_plot_font_defaults(font_defaults: PlotFontDefaults | None = None) -> PlotFontDefaults:
    return font_defaults if font_defaults is not None else DEFAULT_PLOT_FONT_DEFAULTS


def parse_plot_font_defaults(
    *,
    title: str | float,
    tick: str | float,
    label: str | float,
    legend: str | float,
) -> PlotFontDefaults:
    return PlotFontDefaults(
        title=_parse_positive_float(title, "Title size"),
        tick=_parse_positive_float(tick, "Tick size"),
        label=_parse_positive_float(label, "Label size"),
        legend=_parse_positive_float(legend, "Legend size"),
    )


def x_tick_label_pad_points(tick_size: float) -> float:
    size = max(1.0, float(tick_size))
    # Keep a little extra breathing room at the bottom corners when tick fonts grow.
    return min(12.0, max(6.0, size * 0.6))


def apply_x_tick_label_padding(ax, tick_size: float) -> float:
    pad = x_tick_label_pad_points(tick_size)
    ax.tick_params(axis="x", pad=pad)
    return pad


def axis_bottom_footprint_px(fig: Figure, ax) -> float:
    renderer = fig.canvas.get_renderer()
    axis_bbox = ax.xaxis.get_tightbbox(renderer)
    axes_bbox = ax.get_window_extent(renderer)
    if axis_bbox is None:
        return 0.0
    return max(0.0, axes_bbox.y0 - axis_bbox.y0)


def ensure_axis_bottom_margin(
    fig: Figure,
    ax,
    min_bottom_margin: float,
    tick_size: float,
    *,
    max_bottom_margin: float = 0.42,
) -> float:
    fig.canvas.draw()
    fig_height_px = fig.get_size_inches()[1] * fig.dpi
    if fig_height_px <= 0:
        return min_bottom_margin

    bottom_px = axis_bottom_footprint_px(fig, ax)
    outer_pad_px = max(10.0, float(tick_size))
    dynamic_bottom_margin = (bottom_px + outer_pad_px) / fig_height_px
    return min(max_bottom_margin, max(float(min_bottom_margin), dynamic_bottom_margin))


def apply_plot_font_defaults(fig: Figure, font_defaults: PlotFontDefaults | None = None) -> PlotFontDefaults:
    defaults = resolve_plot_font_defaults(font_defaults)

    for ax in fig.axes:
        ax.tick_params(axis="both", labelsize=defaults.tick)
        apply_x_tick_label_padding(ax, defaults.tick)
        ax.xaxis.label.set_fontsize(defaults.label)
        ax.yaxis.label.set_fontsize(defaults.label)
        ax.title.set_fontsize(defaults.title)

        legend = ax.get_legend()
        if legend is not None:
            for text in legend.get_texts():
                text.set_fontsize(defaults.legend)
            legend_title = legend.get_title()
            if legend_title is not None:
                legend_title.set_fontsize(defaults.legend)

    if fig._suptitle is not None:
        fig._suptitle.set_fontsize(defaults.title)

    return defaults
