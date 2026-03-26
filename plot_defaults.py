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


def apply_plot_font_defaults(fig: Figure, font_defaults: PlotFontDefaults | None = None) -> PlotFontDefaults:
    defaults = resolve_plot_font_defaults(font_defaults)

    for ax in fig.axes:
        ax.tick_params(axis="both", labelsize=defaults.tick)
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
