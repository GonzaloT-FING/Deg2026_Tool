"""Experimental multi-EIS fitting pipeline scaffold.

This module is intentionally separate from ``eis_pip.py`` so we can develop
the fitting workflow without destabilizing the current EIS export/plot tools.
The first step is a preparation UI: choose spectra, assign the varying
parameter value for each one, validate a common frequency grid, and define the
equivalent-circuit expression that will later drive the optimizer.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from functools import lru_cache
from pathlib import Path
from typing import Iterable
import math
import re

import tkinter as tk
from tkinter import ttk, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.ticker import LogLocator, MaxNLocator

from plot_defaults import make_legend_draggable
from ui_layout import create_scrollable_controls


EIS_FILENAME_TOKEN = "EISPOT"
VARYING_PARAMETER_OPTIONS = ["Voltage"]
WEIGHTING_OPTIONS = ["Modulus", "Unit"]
PARAMETER_TREATMENT_OPTIONS = ["shared", "almost_shared", "smooth", "independent"]
ZHIT_SMOOTHING_OPTIONS = ["modsinc", "none", "lowess", "savgol", "whithend", "auto"]
ZHIT_INTERPOLATION_OPTIONS = ["makima", "akima", "cubic", "pchip", "auto"]
ZHIT_WINDOW_OPTIONS = ["auto", "boxcar", "cosine", "triang", "hann", "hamming", "bartlett", "blackman"]
CIRCUIT_TEMPLATE_MAP = {
    "Randles": "R0-p(R1,C1)",
    "Randles + CPE": "R0-p(R1,CPE1)",
    "Two time constants": "R0-p(R1,C1)-p(R2,C2)",
    "Two time constants + CPE": "R0-p(R1,CPE1)-p(R2,CPE2)",
    "Ladder with inductor": "L0-R0-p(R1,CPE1)-p(R2,CPE2)-p(R3,CPE3)",
    "Custom": "",
}


@dataclass(frozen=True)
class EisFile:
    path: Path


@dataclass
class ParsedDTA:
    meta_values: dict[str, str]
    meta_units: dict[str, str]
    header: list[str]
    units: list[str]
    rows: list[list[str]]


@dataclass
class EISSpectrum:
    source: EisFile
    frequency_hz: list[float]
    z_real_ohm: list[float]
    z_imag_ohm: list[float]
    z_mod_ohm: list[float] = field(default_factory=list)
    z_phase_deg: list[float] = field(default_factory=list)
    temperature_c: list[float] = field(default_factory=list)
    metadata: dict[str, str] = field(default_factory=dict)
    sweep_parameter_name: str = "Voltage"
    sweep_parameter_value: float | None = None


@dataclass
class MultiEISDataset:
    spectra: list[EISSpectrum]

    @property
    def count(self) -> int:
        return len(self.spectra)


@dataclass
class ParameterTreatmentConfig:
    name: str
    element_token: str
    quantity_label: str
    unit: str
    treatment_mode: str
    initial_guess: float
    lower_bound: float
    upper_bound: float
    smoothness_weight: float = 0.0
    similarity_weight: float = 0.0


@dataclass
class FitStrategyConfig:
    circuit_expression: str
    circuit_formula: str
    parameter_configs: list[ParameterTreatmentConfig] = field(default_factory=list)
    varying_parameter_name: str = "Voltage"
    weighting_mode: str = "Modulus"
    sequential_seed: bool = True
    smoothness_enabled: bool = True
    max_iterations: int = 500


@dataclass
class SpectrumFitResult:
    source_name: str
    success: bool
    message: str = ""
    sweep_parameter_value: float | None = None
    parameters: dict[str, float] = field(default_factory=dict)
    parameter_standard_errors: dict[str, float] = field(default_factory=dict)
    objective_value: float | None = None


@dataclass
class MultiEISFitResult:
    dataset_size: int
    strategy: FitStrategyConfig
    success: bool = False
    message: str = ""
    objective_value: float | None = None
    nfev: int | None = None
    parameter_trajectories: dict[str, list[float]] = field(default_factory=dict)
    parameter_standard_error_trajectories: dict[str, list[float]] = field(default_factory=dict)
    spectrum_results: list[SpectrumFitResult] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class ParameterLayoutEntry:
    config: ParameterTreatmentConfig
    start: int
    stop: int


@dataclass
class ValidationReport:
    text: str
    kk_success_count: int = 0
    zhit_success_count: int = 0


def to_float(val: str) -> float | None:
    text = val.strip()
    if not text:
        return None
    text = text.replace(",", ".")
    try:
        return float(text)
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


def _apply_bar_axis_tick_aligned_limits(ax, values: Iterable[float], tick_count: int = 6) -> None:
    finite_values: list[float] = []
    for value in values:
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            continue
        if math.isfinite(numeric):
            finite_values.append(numeric)

    if not finite_values:
        return

    target_tick_count = max(3, int(tick_count))

    if ax.get_yscale() == "log":
        positive_values = [value for value in finite_values if value > 0]
        if not positive_values:
            return

        axis_min = min(positive_values)
        axis_max = max(positive_values)
        if math.isclose(axis_min, axis_max):
            axis_min /= 10.0
            axis_max *= 10.0

        locator = LogLocator(base=10.0, numticks=target_tick_count)
    else:
        axis_min = min(finite_values)
        axis_max = max(finite_values)

        if axis_min >= 0:
            axis_min = 0.0
        elif axis_max <= 0:
            axis_max = 0.0

        if math.isclose(axis_min, axis_max):
            if math.isclose(axis_min, 0.0):
                axis_max = 1.0
            else:
                padding = max(abs(axis_max) * 0.1, 1.0)
                axis_min -= padding
                axis_max += padding

        locator = MaxNLocator(nbins=target_tick_count)

    tick_values = list(locator.tick_values(axis_min, axis_max))
    if not tick_values:
        return

    lower_candidates = [tick for tick in tick_values if tick <= axis_min + 1e-12]
    upper_candidates = [tick for tick in tick_values if tick >= axis_max - 1e-12]
    axis_bottom = lower_candidates[-1] if lower_candidates else axis_min
    axis_top = upper_candidates[0] if upper_candidates else axis_max
    visible_ticks = [
        tick
        for tick in tick_values
        if axis_bottom - 1e-12 <= tick <= axis_top + 1e-12
    ]

    if len(visible_ticks) < 2:
        return

    ax.set_ylim(bottom=axis_bottom, top=axis_top)
    ax.set_yticks(visible_ticks)


def parse_gamry_dta(path: Path) -> ParsedDTA:
    """Parse one Gamry EIS .DTA file containing a ZCURVE table."""
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
                meta_values[key] = parts[2].strip()
                description = parts[-1].strip() if len(parts) >= 4 else ""
                meta_units[key] = _extract_parenthesized_unit(description)
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
        raise ValueError(f"No data header found in {path.name} (expected 'Pt ...').")

    return ParsedDTA(
        meta_values=meta_values,
        meta_units=meta_units,
        header=header,
        units=units,
        rows=rows,
    )


def _column_index(parsed: ParsedDTA, column_name: str) -> int | None:
    try:
        return parsed.header.index(column_name)
    except ValueError:
        return None


def _extract_numeric_column(parsed: ParsedDTA, column_name: str) -> list[float]:
    idx = _column_index(parsed, column_name)
    if idx is None:
        return []

    values: list[float] = []
    for row in parsed.rows:
        if idx >= len(row):
            continue
        num = to_float(row[idx])
        if num is not None:
            values.append(num)
    return values


def build_spectrum(source: EisFile, parsed: ParsedDTA) -> EISSpectrum:
    frequency_hz = _extract_numeric_column(parsed, "Freq")
    z_real_ohm = _extract_numeric_column(parsed, "Zreal")
    z_imag_ohm = _extract_numeric_column(parsed, "Zimag")

    if not frequency_hz or not z_real_ohm or not z_imag_ohm:
        raise ValueError(f"{source.path.name}: faltan columnas EIS requeridas para fitting.")

    if not (len(frequency_hz) == len(z_real_ohm) == len(z_imag_ohm)):
        raise ValueError(f"{source.path.name}: las columnas EIS no tienen el mismo largo.")

    return EISSpectrum(
        source=source,
        frequency_hz=frequency_hz,
        z_real_ohm=z_real_ohm,
        z_imag_ohm=z_imag_ohm,
        z_mod_ohm=_extract_numeric_column(parsed, "Zmod"),
        z_phase_deg=_extract_numeric_column(parsed, "Zphz"),
        temperature_c=_extract_numeric_column(parsed, "Temp"),
        metadata=dict(parsed.meta_values),
    )


def find_eis_files(input_dir: Path) -> list[EisFile]:
    return [
        EisFile(path=path)
        for path in sorted(Path(input_dir).glob("*.DTA"))
        if path.is_file() and EIS_FILENAME_TOKEN in path.name
    ]


def load_dataset_from_files(eis_files: Iterable[Path | EisFile]) -> MultiEISDataset:
    spectra: list[EISSpectrum] = []
    for item in eis_files:
        eis_file = item if isinstance(item, EisFile) else EisFile(path=Path(item))
        parsed = parse_gamry_dta(eis_file.path)
        spectra.append(build_spectrum(eis_file, parsed))
    return MultiEISDataset(spectra=spectra)


def load_dataset(input_dir: Path) -> MultiEISDataset:
    return load_dataset_from_files(find_eis_files(input_dir))


def validate_dataset(dataset: MultiEISDataset) -> None:
    if not dataset.spectra:
        raise ValueError("No se encontraron espectros EIS para fitting.")


def validate_frequency_grid(dataset: MultiEISDataset, rel_tol: float = 1e-9, abs_tol: float = 1e-12) -> None:
    validate_dataset(dataset)

    reference = dataset.spectra[0]
    ref_freq = reference.frequency_hz

    for spectrum in dataset.spectra[1:]:
        if len(spectrum.frequency_hz) != len(ref_freq):
            raise ValueError(
                f"{spectrum.source.path.name}: la cantidad de frecuencias no coincide con "
                f"{reference.source.path.name}."
            )

        for idx, (expected, current) in enumerate(zip(ref_freq, spectrum.frequency_hz), start=1):
            tolerance = max(abs_tol, abs(expected) * rel_tol)
            if abs(expected - current) > tolerance:
                raise ValueError(
                    f"{spectrum.source.path.name}: la frecuencia #{idx} no coincide con "
                    f"{reference.source.path.name} ({expected:g} Hz vs {current:g} Hz)."
                )


def _format_frequency(value: float) -> str:
    if value == 0:
        return "0 Hz"
    if abs(value) >= 1000 or abs(value) < 0.1:
        return f"{value:.3g} Hz"
    return f"{value:.4f}".rstrip("0").rstrip(".") + " Hz"


def _frequency_grid_summary(dataset: MultiEISDataset) -> str:
    validate_dataset(dataset)
    freq = dataset.spectra[0].frequency_hz
    return f"{len(freq)} pts | {_format_frequency(max(freq))} -> {_format_frequency(min(freq))}"


def _default_voltage_text(spectrum: EISSpectrum) -> str:
    for key in ("VDC", "Vdc"):
        raw = spectrum.metadata.get(key, "")
        value = to_float(raw)
        if value is not None:
            return f"{value:.6g}"
    return ""


def _tokenize_circuit_expression(expression: str) -> list[str]:
    cleaned = re.sub(r"\s+", "", expression)
    if not cleaned:
        return []

    token_re = re.compile(r"CPE\d+|[LRCW]\d+|p|[(),-]")
    tokens: list[str] = []
    pos = 0
    while pos < len(cleaned):
        match = token_re.match(cleaned, pos)
        if match is None:
            raise ValueError(f"Circuito invalido cerca de: {cleaned[pos:]}")
        tokens.append(match.group(0))
        pos = match.end()
    return tokens


def _component_formula(token: str) -> str:
    if token.startswith("CPE"):
        suffix = token[3:]
        return f"1/(Q{suffix}*(j*w)^n{suffix})"

    kind = token[0]
    suffix = token[1:]
    if kind == "R":
        return f"R{suffix}"
    if kind == "L":
        return f"j*w*L{suffix}"
    if kind == "C":
        return f"1/(j*w*C{suffix})"
    if kind == "W":
        return f"sigma{suffix}/sqrt(j*w)"
    raise ValueError(f"Elemento no soportado: {token}")


def _parse_circuit_expression(tokens: list[str]) -> tuple[object, int]:
    def parse_series(pos: int, stop_tokens: set[str]) -> tuple[object, int]:
        items: list[object] = []
        node, pos = parse_term(pos)
        items.append(node)

        while pos < len(tokens) and tokens[pos] not in stop_tokens:
            if tokens[pos] != "-":
                raise ValueError(f"Se esperaba '-' y se encontro '{tokens[pos]}'.")
            node, pos = parse_term(pos + 1)
            items.append(node)

        if len(items) == 1:
            return items[0], pos
        return ("series", items), pos

    def parse_term(pos: int) -> tuple[object, int]:
        if pos >= len(tokens):
            raise ValueError("Expresion de circuito incompleta.")

        token = tokens[pos]
        if token == "p":
            if pos + 1 >= len(tokens) or tokens[pos + 1] != "(":
                raise ValueError("Despues de 'p' debe venir '('.")
            pos += 2
            items: list[object] = []
            while True:
                item, pos = parse_series(pos, {",", ")"})
                items.append(item)
                if pos >= len(tokens):
                    raise ValueError("Falta ')' en un bloque paralelo.")
                if tokens[pos] == ")":
                    pos += 1
                    break
                if tokens[pos] != ",":
                    raise ValueError(f"Se esperaba ',' y se encontro '{tokens[pos]}'.")
                pos += 1

            if len(items) < 2:
                raise ValueError("Un bloque paralelo debe tener al menos dos ramas.")
            return ("parallel", items), pos

        if re.fullmatch(r"CPE\d+|[LRCW]\d+", token):
            return ("component", token), pos + 1

        raise ValueError(f"Token no soportado en el circuito: {token}")

    tree, end_pos = parse_series(0, set())
    if end_pos != len(tokens):
        raise ValueError(f"Token inesperado al final del circuito: {tokens[end_pos]}")
    return tree, end_pos


def _formula_from_circuit_tree(node: object) -> str:
    kind, payload = node
    if kind == "component":
        return _component_formula(payload)
    if kind == "series":
        return "(" + " + ".join(_formula_from_circuit_tree(item) for item in payload) + ")"
    if kind == "parallel":
        return "1/(" + " + ".join(f"1/({_formula_from_circuit_tree(item)})" for item in payload) + ")"
    raise ValueError(f"Nodo de circuito no soportado: {kind}")


def _circuit_formula_preview(expression: str) -> str:
    tokens = _tokenize_circuit_expression(expression)
    if not tokens:
        raise ValueError("Debe definir una expresion para el circuito equivalente.")
    tree, _ = _parse_circuit_expression(tokens)
    return _formula_from_circuit_tree(tree)


def _validate_circuit_expression(expression: str) -> str:
    cleaned = expression.strip()
    if not cleaned:
        raise ValueError("Debe definir una expresion para el circuito equivalente.")
    _circuit_formula_preview(cleaned)
    return cleaned


def _parse_circuit_tree_from_expression(expression: str) -> object:
    tokens = _tokenize_circuit_expression(expression)
    if not tokens:
        raise ValueError("Debe definir una expresion para el circuito equivalente.")
    tree, _ = _parse_circuit_expression(tokens)
    return tree


def _component_tokens_in_order(node: object) -> list[str]:
    kind, payload = node
    if kind == "component":
        return [payload]

    tokens: list[str] = []
    seen: set[str] = set()
    for child in payload:
        for token in _component_tokens_in_order(child):
            if token in seen:
                continue
            seen.add(token)
            tokens.append(token)
    return tokens


def _default_parameter_configs(expression: str) -> list[ParameterTreatmentConfig]:
    tree = _parse_circuit_tree_from_expression(expression)
    configs: list[ParameterTreatmentConfig] = []

    for token in _component_tokens_in_order(tree):
        kind = token[0]
        suffix = token[3:] if token.startswith("CPE") else token[1:]

        if token.startswith("CPE"):
            configs.append(
                ParameterTreatmentConfig(
                    name=f"Q{suffix}",
                    element_token=token,
                    quantity_label="Q",
                    unit="S*s^n",
                    treatment_mode="smooth",
                    initial_guess=1e-4,
                    lower_bound=1e-12,
                    upper_bound=1e3,
                    smoothness_weight=1.0,
                    similarity_weight=0.0,
                )
            )
            configs.append(
                ParameterTreatmentConfig(
                    name=f"n{suffix}",
                    element_token=token,
                    quantity_label="n",
                    unit="-",
                    treatment_mode="smooth",
                    initial_guess=0.9,
                    lower_bound=0.05,
                    upper_bound=1.0,
                    smoothness_weight=0.2,
                    similarity_weight=0.0,
                )
            )
            continue

        if kind == "L":
            configs.append(
                ParameterTreatmentConfig(
                    name=token,
                    element_token=token,
                    quantity_label="L",
                    unit="H",
                    treatment_mode="shared",
                    initial_guess=1e-6,
                    lower_bound=1e-12,
                    upper_bound=1.0,
                )
            )
            continue

        if kind == "R":
            is_series_resistance = token == "R0"
            configs.append(
                ParameterTreatmentConfig(
                    name=token,
                    element_token=token,
                    quantity_label="R",
                    unit="ohm",
                    treatment_mode="almost_shared" if is_series_resistance else "smooth",
                    initial_guess=0.1,
                    lower_bound=1e-9,
                    upper_bound=1e9,
                    smoothness_weight=0.5 if is_series_resistance else 1.0,
                    similarity_weight=20.0 if is_series_resistance else 0.0,
                )
            )
            continue

        if kind == "C":
            configs.append(
                ParameterTreatmentConfig(
                    name=token,
                    element_token=token,
                    quantity_label="C",
                    unit="F",
                    treatment_mode="smooth",
                    initial_guess=1e-3,
                    lower_bound=1e-12,
                    upper_bound=1e3,
                    smoothness_weight=1.0,
                    similarity_weight=0.0,
                )
            )
            continue

        if kind == "W":
            configs.append(
                ParameterTreatmentConfig(
                    name=f"sigma{suffix}",
                    element_token=token,
                    quantity_label="sigma",
                    unit="ohm*s^-0.5",
                    treatment_mode="smooth",
                    initial_guess=1.0,
                    lower_bound=1e-12,
                    upper_bound=1e12,
                    smoothness_weight=1.0,
                    similarity_weight=0.0,
                )
            )
            continue

        raise ValueError(f"Elemento no soportado para defaults de fitting: {token}")

    return configs


def _positive_int(text: str, field_name: str) -> int:
    cleaned = text.strip()
    if not cleaned:
        raise ValueError(f"{field_name} no puede estar vacio.")
    value = int(cleaned)
    if value <= 0:
        raise ValueError(f"{field_name} debe ser mayor que 0.")
    return value


def _positive_float(text: str, field_name: str) -> float:
    cleaned = text.strip().replace(",", ".")
    if not cleaned:
        raise ValueError(f"{field_name} no puede estar vacio.")
    value = float(cleaned)
    if value <= 0:
        raise ValueError(f"{field_name} debe ser mayor que 0.")
    return value


def _nonnegative_float(text: str, field_name: str) -> float:
    cleaned = text.strip().replace(",", ".")
    if not cleaned:
        raise ValueError(f"{field_name} no puede estar vacio.")
    value = float(cleaned)
    if value < 0:
        raise ValueError(f"{field_name} no puede ser negativo.")
    return value


def _required_float(text: str, field_name: str) -> float:
    cleaned = text.strip().replace(",", ".")
    if not cleaned:
        raise ValueError(f"{field_name} no puede estar vacio.")
    return float(cleaned)


def _import_pyimpspec():
    try:
        import numpy as np
        from pyimpspec import DataSet, perform_kramers_kronig_test, perform_zhit
    except Exception as exc:
        raise ValueError(f"No se pudo importar pyimpspec: {exc}") from exc

    return np, DataSet, perform_kramers_kronig_test, perform_zhit


def _build_pyimpspec_dataset(spectrum: EISSpectrum):
    np, DataSet, _perform_kramers_kronig_test, _perform_zhit = _import_pyimpspec()
    frequencies = np.asarray(spectrum.frequency_hz, dtype=float)
    impedances = np.asarray(spectrum.z_real_ohm, dtype=float) + 1j * np.asarray(spectrum.z_imag_ohm, dtype=float)
    return DataSet(
        frequencies=frequencies,
        impedances=impedances,
        path=str(spectrum.source.path),
        label=spectrum.source.path.name,
    )


def _pyimpspec_window_registry_count() -> int | None:
    try:
        from pyimpspec.analysis.zhit.weights import _WINDOW_FUNCTIONS, _initialize_window_functions
    except Exception:
        return None

    try:
        _initialize_window_functions()
    except Exception:
        return None

    return len(_WINDOW_FUNCTIONS)


def _zhit_window_candidates(window_name: str) -> list[str]:
    if window_name == "auto":
        return [name for name in ZHIT_WINDOW_OPTIONS if name != "auto"]
    return [window_name]


def _zhit_window_bounds_hz(center: float, width: float) -> tuple[float, float]:
    min_log_f = center - width / 2.0
    max_log_f = center + width / 2.0
    return 10 ** min_log_f, 10 ** max_log_f


def _generate_custom_zhit_weights(
    log_f,
    window_name: str,
    center: float,
    width: float,
):
    import math

    import numpy as np
    from scipy.interpolate import Akima1DInterpolator
    from scipy.signal import windows as scipy_windows

    window_func = getattr(scipy_windows, window_name, None)
    if not callable(window_func):
        raise ValueError(f"Ventana Z-HIT no soportada por SciPy: {window_name}")

    weights = np.zeros(log_f.shape, dtype=float)
    min_log_f = center - width / 2.0
    max_log_f = center + width / 2.0
    num_points = 10 * int(math.ceil(max_log_f) - math.floor(min_log_f)) + 1

    x_values = [
        value
        for value in np.log10(
            np.logspace(
                math.floor(min_log_f),
                math.ceil(max_log_f),
                num=num_points,
            )
        )
        if min_log_f <= value <= max_log_f
    ]
    if min_log_f not in x_values:
        x_values.insert(0, min_log_f)
    if max_log_f not in x_values:
        x_values.append(max_log_f)

    interpolator = Akima1DInterpolator(
        x_values,
        window_func(M=len(x_values)),
    )
    for index, log_freq in enumerate(log_f):
        if min_log_f <= log_freq <= max_log_f:
            weights[index] = float(interpolator(log_freq))

    weights = np.clip(weights, 0.0, 1.0)
    return weights


def _zhit_spectrum_debug_lines(
    spectrum: EISSpectrum,
    zhit_config: dict[str, object],
) -> list[str]:
    import math

    frequencies = [freq for freq in spectrum.frequency_hz if freq > 0]
    if not frequencies:
        return ["  Z-HIT data: no hay frecuencias positivas disponibles."]

    f_min = min(frequencies)
    f_max = max(frequencies)
    decades = math.log10(f_max) - math.log10(f_min)
    z_imag_min = min(spectrum.z_imag_ohm)
    z_imag_max = max(spectrum.z_imag_ohm)
    center = float(zhit_config.get("center", 1.5))
    width = float(zhit_config.get("width", 3.0))
    lower_hz, upper_hz = _zhit_window_bounds_hz(center, width)
    points_in_window = sum(1 for freq in frequencies if lower_hz <= freq <= upper_hz)

    suggested_center = (math.log10(f_min) + math.log10(f_max)) / 2.0
    suggested_width = max(math.log10(f_max) - math.log10(f_min), 0.5)

    lines = [
        f"  Z-HIT data: {len(frequencies)} pts, {f_min:.6g} to {f_max:.6g} Hz ({decades:.3f} decades).",
        f"  Im(Z) range sent to pyimpspec: {z_imag_min:.6g} to {z_imag_max:.6g} ohm.",
        (
            "  Offset window from center/width: "
            f"{lower_hz:.6g} to {upper_hz:.6g} Hz, "
            f"points inside={points_in_window}."
        ),
    ]
    if points_in_window == 0:
        lines.append(
            "  Hint: ningun punto cae dentro de la ventana de offset. "
            f"Pruebe center={suggested_center:.3f}, width={suggested_width:.3f}."
        )
    elif points_in_window < 3:
        lines.append(
            "  Hint: hay muy pocos puntos dentro de la ventana de offset. "
            f"Puede ser mejor probar center={suggested_center:.3f}, width={suggested_width:.3f}."
        )
    return lines


def _perform_zhit_with_workaround(data, zhit_config: dict[str, object]):
    np, _DataSet, _perform_kramers_kronig_test, perform_zhit = _import_pyimpspec()

    config = dict(zhit_config)
    selected_window = str(config.get("window", "auto"))
    center = float(config.get("center", 1.5))
    width = float(config.get("width", 3.0))
    log_f = np.log10(np.asarray(data.get_frequencies(), dtype=float))
    candidate_windows = _zhit_window_candidates(selected_window)
    successes: list[tuple[float, str, int, object]] = []
    failures: list[str] = []

    for window_name in candidate_windows:
        try:
            weights = _generate_custom_zhit_weights(
                log_f=log_f,
                window_name=window_name,
                center=center,
                width=width,
            )
            weighted_points = int(np.count_nonzero(weights > 0.0))
            if weighted_points == 0:
                failures.append(f"{window_name}: 0 puntos dentro de la ventana.")
                continue

            attempt_config = dict(config)
            attempt_config["window"] = window_name
            attempt_config["weights"] = weights
            result = perform_zhit(data, **attempt_config)
            successes.append((float(result.pseudo_chisqr), window_name, weighted_points, result))
        except Exception as exc:
            failures.append(f"{window_name}: {_friendly_pyimpspec_error('Z-HIT', exc)}")

    if not successes:
        detail = " | ".join(failures[:4])
        if len(failures) > 4:
            detail += " | ..."
        raise ValueError(
            "No se obtuvo ninguna reconstruccion Z-HIT valida con las ventanas "
            f"probadas. Detalles: {detail}"
        )

    successes.sort(key=lambda item: item[0])
    best_pseudo_chisqr, best_window, weighted_points, best_result = successes[0]

    debug_lines = [
        (
            "  Z-HIT workaround: pesos custom aplicados desde la ventana "
            f"'{best_window}' con {weighted_points} puntos ponderados."
        ),
        f"  Best log pseudo chi-squared candidate: {best_pseudo_chisqr:.6g}",
    ]
    if selected_window == "auto" and len(successes) > 1:
        top_candidates = ", ".join(
            f"{window_name}={pseudo_chisqr:.4g}"
            for pseudo_chisqr, window_name, _weighted_points, _result in successes[:3]
        )
        debug_lines.append(f"  Best window candidates: {top_candidates}")

    return best_result, debug_lines


def _statistics_dataframe_to_text(statistics_df, title: str) -> list[str]:
    lines = [title]
    try:
        for _, row in statistics_df.iterrows():
            lines.append(f"  {row['Label']}: {row['Value']}")
    except Exception:
        lines.append(str(statistics_df))
    return lines


def _friendly_pyimpspec_error(kind: str, exc: Exception) -> str:
    if kind == "Z-HIT" and isinstance(exc, IndexError):
        return (
            "pyimpspec no pudo construir ninguna reconstruccion Z-HIT valida con "
            "los datos o ajustes actuales. Normalmente esto apunta a que no hubo "
            "ninguna combinacion util de smoothing/interpolation/window para la "
            "etapa de ajuste del offset."
        )
    return f"{type(exc).__name__}: {exc}"


def run_pyimpspec_validation(dataset: MultiEISDataset) -> ValidationReport:
    return run_pyimpspec_validation_with_config(dataset, zhit_config=None)


def run_pyimpspec_validation_with_config(
    dataset: MultiEISDataset,
    zhit_config: dict[str, object] | None = None,
) -> ValidationReport:
    validate_dataset(dataset)
    validate_frequency_grid(dataset)

    _np, _DataSet, perform_kramers_kronig_test, _perform_zhit = _import_pyimpspec()
    zhit_config = dict(zhit_config or {})
    native_window_count = _pyimpspec_window_registry_count()

    lines = [
        "pyimpspec validation report",
        "",
        f"Selected spectra: {dataset.count}",
        f"Frequency grid: {_frequency_grid_summary(dataset)}",
        (
            "pyimpspec native Z-HIT windows detected: "
            f"{native_window_count if native_window_count is not None else 'unknown'}"
        ),
        "Pipeline workaround: custom Z-HIT weights are enabled.",
        "Z-HIT settings:",
        f"  smoothing={zhit_config.get('smoothing', 'modsinc')}",
        f"  interpolation={zhit_config.get('interpolation', 'makima')}",
        f"  window={zhit_config.get('window', 'auto')}",
        f"  num_points={zhit_config.get('num_points', 3)}",
        f"  polynomial_order={zhit_config.get('polynomial_order', 2)}",
        f"  num_iterations={zhit_config.get('num_iterations', 3)}",
        f"  center={zhit_config.get('center', 1.5)}",
        f"  width={zhit_config.get('width', 3.0)}",
        "",
    ]

    kk_success_count = 0
    zhit_success_count = 0

    for spectrum in dataset.spectra:
        data = _build_pyimpspec_dataset(spectrum)
        param_text = (
            f"{spectrum.sweep_parameter_value:.6g} {spectrum.sweep_parameter_name}"
            if spectrum.sweep_parameter_value is not None
            else spectrum.sweep_parameter_name
        )
        lines.append(f"{spectrum.source.path.name} | {param_text}")
        lines.extend(_zhit_spectrum_debug_lines(spectrum, zhit_config))

        try:
            kk_result = perform_kramers_kronig_test(data)
            kk_success_count += 1
            lines.append("  KK: OK")
            lines.extend(_statistics_dataframe_to_text(kk_result.to_statistics_dataframe(), "  KK statistics:"))
        except Exception as exc:
            lines.append(f"  KK: ERROR - {_friendly_pyimpspec_error('KK', exc)}")

        try:
            zhit_result, zhit_debug_lines = _perform_zhit_with_workaround(data, zhit_config)
            zhit_success_count += 1
            lines.append("  Z-HIT: OK")
            lines.extend(zhit_debug_lines)
            lines.extend(_statistics_dataframe_to_text(zhit_result.to_statistics_dataframe(), "  Z-HIT statistics:"))
        except Exception as exc:
            lines.append(f"  Z-HIT: ERROR - {_friendly_pyimpspec_error('Z-HIT', exc)}")

        lines.append("")

    lines.append(
        f"Summary: KK passed on {kk_success_count}/{dataset.count} spectra, "
        f"Z-HIT passed on {zhit_success_count}/{dataset.count} spectra."
    )

    return ValidationReport(
        text="\n".join(lines),
        kk_success_count=kk_success_count,
        zhit_success_count=zhit_success_count,
    )


def build_selected_dataset(
    dataset: MultiEISDataset,
    selections: list[tuple[int, float]],
    varying_parameter_name: str = "Voltage",
) -> MultiEISDataset:
    if len(selections) < 2:
        raise ValueError("Debe seleccionar al menos dos archivos EIS para multifit.")

    selected_spectra = [
        replace(
            dataset.spectra[idx],
            sweep_parameter_name=varying_parameter_name,
            sweep_parameter_value=value,
        )
        for idx, value in selections
    ]
    selected_spectra.sort(
        key=lambda spectrum: (
            float("inf") if spectrum.sweep_parameter_value is None else spectrum.sweep_parameter_value,
            spectrum.source.path.name.lower(),
        )
    )
    return MultiEISDataset(spectra=selected_spectra)


def _parameter_variable_count(config: ParameterTreatmentConfig, spectrum_count: int) -> int:
    if config.treatment_mode == "shared":
        return 1
    return spectrum_count


def _parameter_treatment_summary_lines(parameter_configs: list[ParameterTreatmentConfig]) -> list[str]:
    lines: list[str] = []
    for config in parameter_configs:
        lines.append(
            (
                f"  {config.name}: mode={config.treatment_mode}, "
                f"guess={config.initial_guess:.6g}, "
                f"bounds=[{config.lower_bound:.6g}, {config.upper_bound:.6g}], "
                f"smooth={config.smoothness_weight:.6g}, "
                f"similarity={config.similarity_weight:.6g}"
            )
        )
    return lines


@lru_cache(maxsize=1)
def _import_fit_backend():
    try:
        import numpy as np
        from scipy.optimize import least_squares
    except Exception as exc:
        raise ValueError(f"No se pudo importar el backend de fitting (NumPy/SciPy): {exc}") from exc

    return np, least_squares


def _build_parameter_layout(
    parameter_configs: list[ParameterTreatmentConfig],
    spectrum_count: int,
) -> list[ParameterLayoutEntry]:
    layout: list[ParameterLayoutEntry] = []
    cursor = 0
    for config in parameter_configs:
        width = _parameter_variable_count(config, spectrum_count)
        layout.append(ParameterLayoutEntry(config=config, start=cursor, stop=cursor + width))
        cursor += width
    return layout


def _clip_parameter_value(config: ParameterTreatmentConfig, value: float) -> float:
    if value < config.lower_bound:
        return config.lower_bound
    if value > config.upper_bound:
        return config.upper_bound
    return value


def _positive_axis_from_dataset(dataset: MultiEISDataset):
    np, _least_squares = _import_fit_backend()

    axis_values = np.asarray(
        [
            float(spectrum.sweep_parameter_value if spectrum.sweep_parameter_value is not None else index)
            for index, spectrum in enumerate(dataset.spectra)
        ],
        dtype=float,
    )
    if axis_values.size > 1 and np.any(np.diff(axis_values) <= 0.0):
        axis_values = np.arange(dataset.count, dtype=float)
    return axis_values


def _prepare_spectrum_arrays(dataset: MultiEISDataset):
    np, _least_squares = _import_fit_backend()

    prepared = []
    for spectrum in dataset.spectra:
        frequencies = np.asarray(spectrum.frequency_hz, dtype=float)
        omega = 2.0 * np.pi * frequencies
        z_exp = np.asarray(spectrum.z_real_ohm, dtype=float) + 1j * np.asarray(spectrum.z_imag_ohm, dtype=float)
        prepared.append(
            {
                "spectrum": spectrum,
                "frequencies": frequencies,
                "omega": omega,
                "z_exp": z_exp,
            }
        )
    return prepared


def _residual_weight_array(z_exp, weighting_mode: str):
    np, _least_squares = _import_fit_backend()

    if weighting_mode == "Modulus":
        return np.maximum(np.abs(z_exp), 1e-12)
    return np.ones_like(z_exp.real, dtype=float)


def _component_impedance(token: str, parameters: dict[str, float], omega):
    np, _least_squares = _import_fit_backend()
    jw = 1j * omega

    if token.startswith("CPE"):
        suffix = token[3:]
        q_value = parameters[f"Q{suffix}"]
        n_value = parameters[f"n{suffix}"]
        return 1.0 / (q_value * np.power(jw, n_value))

    kind = token[0]
    suffix = token[1:]
    if kind == "R":
        return np.full(omega.shape, parameters[token], dtype=complex)
    if kind == "L":
        return 1j * omega * parameters[token]
    if kind == "C":
        return 1.0 / (1j * omega * parameters[token])
    if kind == "W":
        return parameters[f"sigma{suffix}"] / np.sqrt(jw)
    raise ValueError(f"Elemento no soportado para evaluacion: {token}")


def _evaluate_circuit_impedance(node: object, parameters: dict[str, float], omega):
    np, _least_squares = _import_fit_backend()

    kind, payload = node
    if kind == "component":
        return _component_impedance(payload, parameters, omega)
    if kind == "series":
        total = np.zeros(omega.shape, dtype=complex)
        for child in payload:
            total = total + _evaluate_circuit_impedance(child, parameters, omega)
        return total
    if kind == "parallel":
        total_admittance = np.zeros(omega.shape, dtype=complex)
        for child in payload:
            total_admittance = total_admittance + 1.0 / _evaluate_circuit_impedance(child, parameters, omega)
        return 1.0 / total_admittance
    raise ValueError(f"Nodo de circuito no soportado: {kind}")


def _unpack_optimizer_vector(
    log_vector,
    layout: list[ParameterLayoutEntry],
    spectrum_count: int,
):
    np, _least_squares = _import_fit_backend()

    log_vector = np.asarray(log_vector, dtype=float)
    actual_series: dict[str, object] = {}
    log_series: dict[str, object] = {}
    for entry in layout:
        config = entry.config
        values = log_vector[entry.start:entry.stop]
        if config.treatment_mode == "shared":
            scalar_log = float(values[0])
            log_array = np.full(spectrum_count, scalar_log, dtype=float)
            actual_array = np.full(spectrum_count, float(np.exp(scalar_log)), dtype=float)
        else:
            log_array = np.asarray(values, dtype=float)
            actual_array = np.exp(log_array)
        log_series[config.name] = log_array
        actual_series[config.name] = actual_array
    return actual_series, log_series


def _single_spectrum_seed(
    prepared_spectrum: dict[str, object],
    parameter_configs: list[ParameterTreatmentConfig],
    circuit_tree: object,
    weighting_mode: str,
    max_iterations: int,
    initial_values: list[float],
):
    np, least_squares = _import_fit_backend()

    x0 = np.asarray(
        [
            np.log(_clip_parameter_value(config, value))
            for config, value in zip(parameter_configs, initial_values)
        ],
        dtype=float,
    )
    lower_bounds = np.asarray([np.log(config.lower_bound) for config in parameter_configs], dtype=float)
    upper_bounds = np.asarray([np.log(config.upper_bound) for config in parameter_configs], dtype=float)

    z_exp = prepared_spectrum["z_exp"]
    weights = _residual_weight_array(z_exp, weighting_mode)
    omega = prepared_spectrum["omega"]

    def _objective(log_values):
        actual_values = np.exp(np.asarray(log_values, dtype=float))
        parameters = {
            config.name: float(value)
            for config, value in zip(parameter_configs, actual_values)
        }
        with np.errstate(all="ignore"):
            z_model = _evaluate_circuit_impedance(circuit_tree, parameters, omega)
            residual = np.concatenate(
                (
                    (z_model.real - z_exp.real) / weights,
                    (z_model.imag - z_exp.imag) / weights,
                )
            )
        return np.nan_to_num(residual, nan=1e12, posinf=1e12, neginf=-1e12)

    result = least_squares(
        _objective,
        x0=x0,
        bounds=(lower_bounds, upper_bounds),
        method="trf",
        max_nfev=max_iterations,
    )
    actual = np.exp(result.x)
    return [float(value) for value in actual], result


def _build_initial_guess_vector(
    dataset: MultiEISDataset,
    prepared_spectra: list[dict[str, object]],
    strategy: FitStrategyConfig,
    layout: list[ParameterLayoutEntry],
    circuit_tree: object,
    seed_trajectories: dict[str, list[float]] | None = None,
):
    np, _least_squares = _import_fit_backend()

    seeded_values: dict[str, list[float]] = {}
    for config in strategy.parameter_configs:
        seeded_series = None
        if seed_trajectories is not None:
            raw_series = seed_trajectories.get(config.name)
            if raw_series is not None and len(raw_series) == dataset.count:
                seeded_series = [
                    _clip_parameter_value(config, float(value))
                    for value in raw_series
                ]
        if seeded_series is None:
            seeded_series = [_clip_parameter_value(config, config.initial_guess) for _ in range(dataset.count)]
        seeded_values[config.name] = seeded_series

    notes: list[str] = []

    if seed_trajectories is not None:
        notes.append(
            "Manual seed table: using the current editable parameter table as the simultaneous-fit initial seed."
        )
        notes.append("Sequential seeding: skipped because an explicit per-spectrum seed table was supplied.")
    elif strategy.sequential_seed:
        current_values = [
            _clip_parameter_value(config, config.initial_guess)
            for config in strategy.parameter_configs
        ]
        seed_successes = 0
        for index, prepared_spectrum in enumerate(prepared_spectra):
            fitted_values, seed_result = _single_spectrum_seed(
                prepared_spectrum=prepared_spectrum,
                parameter_configs=strategy.parameter_configs,
                circuit_tree=circuit_tree,
                weighting_mode=strategy.weighting_mode,
                max_iterations=max(50, min(strategy.max_iterations, 300)),
                initial_values=current_values,
            )
            if seed_result.success:
                current_values = fitted_values
                seed_successes += 1
            for config, value in zip(strategy.parameter_configs, current_values):
                seeded_values[config.name][index] = _clip_parameter_value(config, value)
        notes.append(f"Sequential seeding: {seed_successes}/{dataset.count} single-spectrum seed fits converged.")
    else:
        notes.append("Sequential seeding: disabled, using the user guesses for every spectrum.")

    x0: list[float] = []
    lower_bounds: list[float] = []
    upper_bounds: list[float] = []
    for entry in layout:
        config = entry.config
        series = [
            _clip_parameter_value(config, value)
            for value in seeded_values[config.name]
        ]
        if config.treatment_mode == "shared":
            representative = float(np.exp(np.mean(np.log(np.asarray(series, dtype=float)))))
            x0.append(np.log(_clip_parameter_value(config, representative)))
            lower_bounds.append(np.log(config.lower_bound))
            upper_bounds.append(np.log(config.upper_bound))
        else:
            for value in series:
                x0.append(np.log(_clip_parameter_value(config, value)))
                lower_bounds.append(np.log(config.lower_bound))
                upper_bounds.append(np.log(config.upper_bound))

    return np.asarray(x0, dtype=float), np.asarray(lower_bounds, dtype=float), np.asarray(upper_bounds, dtype=float), notes


def _trajectory_curvature_residual(log_values, axis_values):
    np, _least_squares = _import_fit_backend()

    log_values = np.asarray(log_values, dtype=float)
    axis_values = np.asarray(axis_values, dtype=float)
    if log_values.size <= 1:
        return np.asarray([], dtype=float)
    if log_values.size == 2:
        delta_x = max(float(axis_values[1] - axis_values[0]), 1e-12)
        return np.asarray([(log_values[1] - log_values[0]) / delta_x], dtype=float)

    residuals: list[float] = []
    for index in range(1, log_values.size - 1):
        left_dx = max(float(axis_values[index] - axis_values[index - 1]), 1e-12)
        right_dx = max(float(axis_values[index + 1] - axis_values[index]), 1e-12)
        left_slope = (log_values[index] - log_values[index - 1]) / left_dx
        right_slope = (log_values[index + 1] - log_values[index]) / right_dx
        curvature = 2.0 * (right_slope - left_slope) / (left_dx + right_dx)
        residuals.append(float(curvature))
    return np.asarray(residuals, dtype=float)


def _simultaneous_objective(
    log_vector,
    prepared_spectra: list[dict[str, object]],
    strategy: FitStrategyConfig,
    layout: list[ParameterLayoutEntry],
    circuit_tree: object,
    axis_values,
):
    np, _least_squares = _import_fit_backend()

    actual_series, log_series = _unpack_optimizer_vector(
        log_vector=log_vector,
        layout=layout,
        spectrum_count=len(prepared_spectra),
    )
    residual_blocks: list[object] = []

    for spectrum_index, prepared in enumerate(prepared_spectra):
        parameters = {
            config.name: float(actual_series[config.name][spectrum_index])
            for config in strategy.parameter_configs
        }
        z_exp = prepared["z_exp"]
        weights = _residual_weight_array(z_exp, strategy.weighting_mode)
        with np.errstate(all="ignore"):
            z_model = _evaluate_circuit_impedance(
                circuit_tree,
                parameters,
                prepared["omega"],
            )
            spectrum_residual = np.concatenate(
                (
                    (z_model.real - z_exp.real) / weights,
                    (z_model.imag - z_exp.imag) / weights,
                )
            )
        residual_blocks.append(np.nan_to_num(spectrum_residual, nan=1e12, posinf=1e12, neginf=-1e12))

    if strategy.smoothness_enabled:
        for entry in layout:
            config = entry.config
            parameter_log_series = log_series[config.name]
            if config.treatment_mode in ("smooth", "almost_shared") and config.smoothness_weight > 0.0:
                curvature = _trajectory_curvature_residual(parameter_log_series, axis_values)
                if curvature.size > 0:
                    residual_blocks.append(np.sqrt(config.smoothness_weight) * curvature)
            if config.treatment_mode == "almost_shared" and config.similarity_weight > 0.0:
                centered = parameter_log_series - np.mean(parameter_log_series)
                residual_blocks.append(np.sqrt(config.similarity_weight) * centered)

    return np.concatenate(residual_blocks)


def _format_parameter_trajectory_lines(
    parameter_trajectories: dict[str, list[float]],
) -> list[str]:
    lines: list[str] = []
    for name, values in parameter_trajectories.items():
        joined = ", ".join(f"{value:.6g}" for value in values)
        lines.append(f"  {name}: {joined}")
    return lines


def _data_residual_degrees_of_freedom(
    prepared_spectra: list[dict[str, object]],
    variable_count: int,
) -> int:
    data_points = sum(2 * int(prepared["z_exp"].size) for prepared in prepared_spectra)
    return max(1, data_points - int(variable_count))


def _simultaneous_scalar_objective(
    log_vector,
    prepared_spectra: list[dict[str, object]],
    strategy: FitStrategyConfig,
    layout: list[ParameterLayoutEntry],
    circuit_tree: object,
    axis_values,
) -> float:
    np, _least_squares = _import_fit_backend()

    residual = np.asarray(
        _simultaneous_objective(
            log_vector,
            prepared_spectra,
            strategy,
            layout,
            circuit_tree,
            axis_values,
        ),
        dtype=float,
    )
    return 0.5 * float(np.dot(residual, residual))


def _finite_difference_hessian(objective, x0):
    np, _least_squares = _import_fit_backend()

    x0 = np.asarray(x0, dtype=float)
    if x0.size == 0:
        return np.zeros((0, 0), dtype=float)

    steps = 1e-4 * np.maximum(1.0, np.abs(x0))
    f0 = float(objective(x0))
    hessian = np.zeros((x0.size, x0.size), dtype=float)

    for i in range(x0.size):
        step_i = steps[i]
        offset_i = np.zeros_like(x0)
        offset_i[i] = step_i

        f_plus = float(objective(x0 + offset_i))
        f_minus = float(objective(x0 - offset_i))
        hessian[i, i] = (f_plus - 2.0 * f0 + f_minus) / (step_i ** 2)

        for j in range(i + 1, x0.size):
            step_j = steps[j]
            offset_j = np.zeros_like(x0)
            offset_j[j] = step_j

            f_pp = float(objective(x0 + offset_i + offset_j))
            f_pm = float(objective(x0 + offset_i - offset_j))
            f_mp = float(objective(x0 - offset_i + offset_j))
            f_mm = float(objective(x0 - offset_i - offset_j))
            mixed_value = (f_pp - f_pm - f_mp + f_mm) / (4.0 * step_i * step_j)
            hessian[i, j] = mixed_value
            hessian[j, i] = mixed_value

    return np.nan_to_num(hessian, nan=0.0, posinf=0.0, neginf=0.0)


def _parameter_standard_error_trajectories_from_log_std(
    std_log_vector,
    layout: list[ParameterLayoutEntry],
    actual_series: dict[str, object],
    spectrum_count: int,
) -> dict[str, list[float]]:
    np, _least_squares = _import_fit_backend()

    std_log_vector = np.asarray(std_log_vector, dtype=float)
    error_trajectories: dict[str, list[float]] = {}
    for entry in layout:
        config = entry.config
        actual_values = np.asarray(actual_series[config.name], dtype=float)
        log_std_values = std_log_vector[entry.start:entry.stop]

        if config.treatment_mode == "shared":
            scalar_std = float(abs(log_std_values[0])) if log_std_values.size else float("nan")
            error_values = actual_values * scalar_std
        else:
            expanded_log_std = np.asarray(log_std_values, dtype=float)
            if expanded_log_std.size != spectrum_count:
                expanded_log_std = np.resize(expanded_log_std, spectrum_count)
            error_values = actual_values * np.abs(expanded_log_std)

        error_trajectories[config.name] = [
            float(value) if math.isfinite(float(value)) else float("nan")
            for value in np.asarray(error_values, dtype=float)
        ]
    return error_trajectories


def _compute_hessian_inverse_standard_errors(
    optimized_log_vector,
    prepared_spectra: list[dict[str, object]],
    strategy: FitStrategyConfig,
    layout: list[ParameterLayoutEntry],
    circuit_tree: object,
    axis_values,
    actual_series: dict[str, object],
) -> tuple[dict[str, list[float]], list[str]]:
    np, _least_squares = _import_fit_backend()

    optimized_log_vector = np.asarray(optimized_log_vector, dtype=float)
    if optimized_log_vector.size == 0:
        return {}, ["Parameter standard errors: unavailable because no fit variables were present."]

    objective = lambda values: _simultaneous_scalar_objective(
        values,
        prepared_spectra,
        strategy,
        layout,
        circuit_tree,
        axis_values,
    )
    hessian = _finite_difference_hessian(objective, optimized_log_vector)
    hessian = 0.5 * (hessian + hessian.T)

    inverse_kind = "inverse"
    try:
        covariance_log = np.linalg.inv(hessian)
    except np.linalg.LinAlgError:
        covariance_log = np.linalg.pinv(hessian)
        inverse_kind = "pseudoinverse"

    residual = np.asarray(
        _simultaneous_objective(
            optimized_log_vector,
            prepared_spectra,
            strategy,
            layout,
            circuit_tree,
            axis_values,
        ),
        dtype=float,
    )
    dof = _data_residual_degrees_of_freedom(prepared_spectra, optimized_log_vector.size)
    wrms = float(np.dot(residual, residual)) / float(dof)
    covariance_log = np.asarray(covariance_log, dtype=float) * wrms
    covariance_log = np.nan_to_num(covariance_log, nan=0.0, posinf=0.0, neginf=0.0)

    covariance_diag = np.asarray(np.diag(covariance_log), dtype=float)
    negative_variances = int(np.count_nonzero(covariance_diag < 0.0))
    std_log = np.sqrt(np.clip(covariance_diag, 0.0, None))
    error_trajectories = _parameter_standard_error_trajectories_from_log_std(
        std_log,
        layout,
        actual_series,
        len(prepared_spectra),
    )

    notes = [
        f"Parameter standard errors: Hessian {inverse_kind} in log-parameter space.",
        f"Covariance scaling: wrms={wrms:.6g} with dof={dof}.",
    ]
    if negative_variances > 0:
        notes.append(
            f"Variance clipping: {negative_variances} negative Hessian-derived variances were clipped to zero."
        )
    return error_trajectories, notes


def fit_dataset(
    dataset: MultiEISDataset,
    strategy: FitStrategyConfig,
    seed_trajectories: dict[str, list[float]] | None = None,
) -> MultiEISFitResult:
    """Run a simultaneous multi-spectrum fit in log-parameter space."""
    validate_dataset(dataset)
    validate_frequency_grid(dataset)
    if not strategy.parameter_configs:
        raise ValueError("No hay parametros configurados para el fitting simultaneo.")

    total_variables = sum(
        _parameter_variable_count(config, dataset.count)
        for config in strategy.parameter_configs
    )
    circuit_tree = _parse_circuit_tree_from_expression(strategy.circuit_expression)
    axis_values = _positive_axis_from_dataset(dataset)
    prepared_spectra = _prepare_spectrum_arrays(dataset)
    layout = _build_parameter_layout(strategy.parameter_configs, dataset.count)
    x0, lower_bounds, upper_bounds, seed_notes = _build_initial_guess_vector(
        dataset=dataset,
        prepared_spectra=prepared_spectra,
        strategy=strategy,
        layout=layout,
        circuit_tree=circuit_tree,
        seed_trajectories=seed_trajectories,
    )
    np, least_squares = _import_fit_backend()

    notes = [
        f"Detected {dataset.count} spectrum/s ordered by {strategy.varying_parameter_name}.",
        f"Circuit expression: {strategy.circuit_expression}",
        f"Circuit formula: {strategy.circuit_formula}",
        f"Weighting mode: {strategy.weighting_mode}",
        f"Parameter count: {len(strategy.parameter_configs)} logical parameters.",
        f"Simultaneous variable count: {total_variables}.",
        "Optimizer space: all parameters are fitted in log space.",
        "Regularization operator: D2 is the second-difference / discrete curvature of",
        "each parameter trajectory versus voltage, so the user does not enter a D2 value.",
        f"Smoothness regularization: {'enabled' if strategy.smoothness_enabled else 'disabled'}.",
    ]
    notes.extend(seed_notes)
    notes.extend(_parameter_treatment_summary_lines(strategy.parameter_configs))

    result = least_squares(
        _simultaneous_objective,
        x0=x0,
        bounds=(lower_bounds, upper_bounds),
        method="trf",
        max_nfev=strategy.max_iterations,
        args=(
            prepared_spectra,
            strategy,
            layout,
            circuit_tree,
            axis_values,
        ),
    )

    actual_series, _log_series = _unpack_optimizer_vector(
        log_vector=result.x,
        layout=layout,
        spectrum_count=dataset.count,
    )
    parameter_trajectories = {
        config.name: [float(value) for value in actual_series[config.name]]
        for config in strategy.parameter_configs
    }
    parameter_standard_error_trajectories, standard_error_notes = _compute_hessian_inverse_standard_errors(
        optimized_log_vector=result.x,
        prepared_spectra=prepared_spectra,
        strategy=strategy,
        layout=layout,
        circuit_tree=circuit_tree,
        axis_values=axis_values,
        actual_series=actual_series,
    )

    spectrum_results: list[SpectrumFitResult] = []
    for spectrum_index, prepared in enumerate(prepared_spectra):
        parameters = {
            config.name: float(actual_series[config.name][spectrum_index])
            for config in strategy.parameter_configs
        }
        parameter_standard_errors = {
            config.name: float(
                parameter_standard_error_trajectories.get(config.name, [float("nan")] * dataset.count)[spectrum_index]
            )
            for config in strategy.parameter_configs
        }
        z_exp = prepared["z_exp"]
        weights = _residual_weight_array(z_exp, strategy.weighting_mode)
        with np.errstate(all="ignore"):
            z_model = _evaluate_circuit_impedance(circuit_tree, parameters, prepared["omega"])
            local_residual = np.concatenate(
                (
                    (z_model.real - z_exp.real) / weights,
                    (z_model.imag - z_exp.imag) / weights,
                )
            )
        local_residual = np.nan_to_num(local_residual, nan=1e12, posinf=1e12, neginf=-1e12)
        spectrum_results.append(
            SpectrumFitResult(
                source_name=prepared["spectrum"].source.path.name,
                success=bool(result.success),
                message=result.message,
                sweep_parameter_value=prepared["spectrum"].sweep_parameter_value,
                parameters=parameters,
                parameter_standard_errors=parameter_standard_errors,
                objective_value=float(np.mean(local_residual**2)),
            )
        )

    notes.append(f"Global objective (sum of squared residuals): {2.0 * result.cost:.6g}")
    notes.extend(standard_error_notes)
    notes.extend(["Parameter trajectories:"])
    notes.extend(_format_parameter_trajectory_lines(parameter_trajectories))
    notes.extend(["Parameter standard errors:"])
    notes.extend(_format_parameter_trajectory_lines(parameter_standard_error_trajectories))

    return MultiEISFitResult(
        dataset_size=dataset.count,
        strategy=strategy,
        success=bool(result.success),
        message=result.message,
        objective_value=float(2.0 * result.cost),
        nfev=getattr(result, "nfev", None),
        parameter_trajectories=parameter_trajectories,
        parameter_standard_error_trajectories=parameter_standard_error_trajectories,
        spectrum_results=spectrum_results,
        notes=notes,
    )


def _build_scrollable_controls(parent, width: int = 560) -> tuple[ttk.Frame, ttk.Frame]:
    outer, inner = create_scrollable_controls(
        parent,
        outer_padding=0,
        fixed_width=width,
        pack=False,
        inner_padding=10,
    )
    return outer, inner


def _set_text(widget: tk.Text, content: str) -> None:
    widget.configure(state="normal")
    widget.delete("1.0", tk.END)
    widget.insert("1.0", content)
    widget.configure(state="disabled")


def open_multifit_window(
    input_dir: Path | None = None,
    eis_files: Iterable[Path] | None = None,
) -> None:
    if eis_files is not None:
        dataset = load_dataset_from_files(eis_files)
    elif input_dir is not None:
        dataset = load_dataset(Path(input_dir))
    else:
        raise ValueError("Se requiere input_dir o eis_files para abrir MultiFit.")

    validate_dataset(dataset)

    created_root = False
    root = tk._default_root
    if root is None:
        root = tk.Tk()
        root.withdraw()
        created_root = True

    win = tk.Toplevel(root)
    win.title("EIS - MultiFit")
    win.geometry("1480x820")
    win.minsize(1180, 720)

    main_pane = ttk.Panedwindow(win, orient="horizontal")
    main_pane.pack(fill="both", expand=True)

    controls_outer, controls_frame = _build_scrollable_controls(main_pane, width=560)
    right_frame = ttk.Frame(main_pane, padding=10)
    right_pane = ttk.Panedwindow(right_frame, orient="vertical")
    right_pane.pack(fill="both", expand=True)

    main_pane.add(controls_outer, weight=0)
    main_pane.add(right_frame, weight=1)

    preview_box = ttk.LabelFrame(right_pane, text="Nyquist preview", padding=10)

    preview_fig = Figure(figsize=(7.2, 4.6), dpi=100)
    preview_canvas = FigureCanvasTkAgg(preview_fig, master=preview_box)
    preview_canvas.draw()
    preview_canvas.get_tk_widget().pack(fill="both", expand=True)

    details_notebook = ttk.Notebook(right_pane)

    right_pane.add(preview_box, weight=3)
    right_pane.add(details_notebook, weight=2)

    summary_tab = ttk.Frame(details_notebook, padding=10)
    validation_tab = ttk.Frame(details_notebook, padding=10)
    notes_tab = ttk.Frame(details_notebook, padding=10)
    parameters_tab = ttk.Frame(details_notebook, padding=0)
    details_notebook.add(summary_tab, text="Summary")
    details_notebook.add(validation_tab, text="Validation")
    details_notebook.add(notes_tab, text="Notes")
    details_notebook.add(parameters_tab, text="Parameters")

    summary_text = tk.Text(summary_tab, height=14, wrap="word", state="disabled")
    summary_text.pack(fill="both", expand=True)

    validation_text = tk.Text(validation_tab, wrap="word", state="disabled")
    validation_text.pack(fill="both", expand=True)

    notes_text = tk.Text(notes_tab, wrap="word", state="disabled")
    notes_text.pack(fill="both", expand=True)

    ttk.Label(
        parameters_tab,
        text="Run the simultaneous fit to populate the fitted-parameter charts grouped by voltage.",
        wraplength=780,
        justify="left",
        foreground="#555555",
    ).pack(anchor="w", fill="x", padx=10, pady=(10, 8))

    parameter_plots_expanded_var = tk.BooleanVar(value=True)
    parameter_plots_toggle_text = tk.StringVar(value="Hide parameter plots")

    parameters_plot_toggle_row = ttk.Frame(parameters_tab)
    parameters_plot_toggle_row.pack(fill="x", padx=10, pady=(0, 8))
    ttk.Button(
        parameters_plot_toggle_row,
        textvariable=parameter_plots_toggle_text,
        command=lambda: _toggle_parameter_plot_visibility(),
    ).pack(side="left")

    parameters_plot_box = ttk.LabelFrame(parameters_tab, text="Fitted parameters vs varying parameter", padding=10)
    parameters_plot_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    parameters_plot_frame = ttk.Frame(parameters_plot_box)
    parameters_plot_frame.pack(fill="both", expand=True)

    parameters_plot_view = tk.Canvas(parameters_plot_frame, highlightthickness=0)
    parameters_plot_scrollbar = ttk.Scrollbar(
        parameters_plot_frame,
        orient="horizontal",
        command=parameters_plot_view.xview,
    )
    parameters_plot_view.configure(xscrollcommand=parameters_plot_scrollbar.set)

    parameters_plot_scrollbar.pack(side="bottom", fill="x")
    parameters_plot_view.pack(side="top", fill="both", expand=True)

    parameters_plot_inner = ttk.Frame(parameters_plot_view)
    parameters_plot_inner.pack_propagate(False)
    parameters_plot_window = parameters_plot_view.create_window((0, 0), window=parameters_plot_inner, anchor="nw")

    parameters_fig = Figure(figsize=(8.4, 4.6), dpi=100)
    parameters_canvas = FigureCanvasTkAgg(parameters_fig, master=parameters_plot_inner)
    parameters_canvas.draw()
    parameters_canvas_widget = parameters_canvas.get_tk_widget()
    parameters_canvas_widget.pack(fill="both", expand=True)

    parameters_table_box = ttk.LabelFrame(parameters_tab, text="Grouped parameter table", padding=10)
    parameters_table_box.pack(fill="x", expand=False, padx=10, pady=(0, 10))

    ttk.Label(
        parameters_table_box,
        text=(
            "Edit a fitted value and press Enter or leave the field to update the Nyquist preview. "
            "The gray SE values come from the last simultaneous fit and are not recomputed for manual edits."
        ),
        wraplength=820,
        justify="left",
        foreground="#555555",
    ).pack(anchor="w", fill="x", pady=(0, 8))

    parameters_table_frame = ttk.Frame(parameters_table_box)
    parameters_table_frame.pack(fill="both", expand=True)

    parameters_table_view = tk.Canvas(parameters_table_frame, highlightthickness=0, height=170)
    parameters_table_y = ttk.Scrollbar(parameters_table_frame, orient="vertical", command=parameters_table_view.yview)
    parameters_table_x = ttk.Scrollbar(parameters_table_frame, orient="horizontal", command=parameters_table_view.xview)
    parameters_table_view.configure(
        yscrollcommand=parameters_table_y.set,
        xscrollcommand=parameters_table_x.set,
    )

    parameters_table_y.pack(side="right", fill="y")
    parameters_table_x.pack(side="bottom", fill="x")
    parameters_table_view.pack(side="left", fill="both", expand=True)

    parameters_table_inner = ttk.Frame(parameters_table_view)
    parameters_table_window = parameters_table_view.create_window((0, 0), window=parameters_table_inner, anchor="nw")

    status_var = tk.StringVar(value="Seleccione los archivos EIS y asigne un voltaje a cada uno.")

    varying_parameter_var = tk.StringVar(value="Voltage")
    weighting_mode_var = tk.StringVar(value="Modulus")
    sequential_seed_var = tk.BooleanVar(value=True)
    smoothness_enabled_var = tk.BooleanVar(value=True)
    max_iterations_var = tk.StringVar(value="500")
    zhit_expanded_var = tk.BooleanVar(value=False)
    zhit_smoothing_var = tk.StringVar(value="modsinc")
    zhit_interpolation_var = tk.StringVar(value="makima")
    zhit_window_var = tk.StringVar(value="auto")
    zhit_num_points_var = tk.StringVar(value="3")
    zhit_polynomial_order_var = tk.StringVar(value="2")
    zhit_num_iterations_var = tk.StringVar(value="3")
    zhit_center_var = tk.StringVar(value="1.5")
    zhit_width_var = tk.StringVar(value="3.0")
    circuit_template_var = tk.StringVar(value="Randles")
    circuit_expression_var = tk.StringVar(value=CIRCUIT_TEMPLATE_MAP["Randles"])

    file_rows: list[dict[str, object]] = []
    parameter_rows: list[dict[str, object]] = []
    latest_fit_result: MultiEISFitResult | None = None
    latest_fit_signature: tuple | None = None
    parameter_plot_refresh_after_id: str | None = None
    parameter_table_cells: dict[tuple[int, str], dict[str, object]] = {}
    parameter_table_content_width_px = 0

    ttk.Label(
        controls_frame,
        text=f"Detected EIS files: {dataset.count}",
        justify="left",
        wraplength=320,
    ).pack(anchor="w", pady=(0, 10))

    files_box = ttk.LabelFrame(controls_frame, text="Available files")
    files_box.pack(fill="x", pady=5)

    files_button_row = ttk.Frame(files_box)
    files_button_row.pack(fill="x", padx=8, pady=(8, 4))

    files_inner = ttk.Frame(files_box)
    files_inner.pack(fill="x", padx=8, pady=(0, 8))

    for idx, spectrum in enumerate(dataset.spectra):
        selected_var = tk.BooleanVar(value=True)
        voltage_var = tk.StringVar(value=_default_voltage_text(spectrum))

        row = ttk.Frame(files_inner)
        row.pack(fill="x", pady=4)

        ttk.Checkbutton(
            row,
            variable=selected_var,
            command=lambda: _update_nyquist_preview(),
        ).grid(row=0, column=0, rowspan=2, sticky="nw", padx=(0, 6))
        ttk.Label(row, text=spectrum.source.path.name, wraplength=240, justify="left").grid(
            row=0, column=1, sticky="w"
        )
        ttk.Label(
            row,
            text=f"{len(spectrum.frequency_hz)} pts | {_format_frequency(max(spectrum.frequency_hz))} -> {_format_frequency(min(spectrum.frequency_hz))}",
            foreground="#555555",
        ).grid(row=1, column=1, sticky="w")
        ttk.Label(row, text="Voltage").grid(row=0, column=2, sticky="e", padx=(12, 4))
        voltage_entry = ttk.Entry(row, textvariable=voltage_var, width=10)
        voltage_entry.grid(row=0, column=3, sticky="w")
        ttk.Label(row, text="V").grid(row=0, column=4, sticky="w", padx=(4, 0))
        voltage_entry.bind("<Return>", lambda _event: _update_nyquist_preview())
        voltage_entry.bind("<FocusOut>", lambda _event: _update_nyquist_preview())

        file_rows.append(
            {
                "index": idx,
                "selected_var": selected_var,
                "voltage_var": voltage_var,
                "spectrum": spectrum,
            }
        )

    parameter_box = ttk.LabelFrame(controls_frame, text="Varying parameter")
    parameter_box.pack(fill="x", pady=5)

    ttk.Label(parameter_box, text="Parameter").grid(row=0, column=0, sticky="w", padx=8, pady=6)
    parameter_combo = ttk.Combobox(
        parameter_box,
        textvariable=varying_parameter_var,
        values=VARYING_PARAMETER_OPTIONS,
        state="readonly",
        width=12,
    )
    parameter_combo.grid(row=0, column=1, sticky="w", padx=8, pady=6)

    circuit_box = ttk.LabelFrame(controls_frame, text="Equivalent circuit")
    circuit_box.pack(fill="x", pady=5)

    ttk.Label(circuit_box, text="Template").grid(row=0, column=0, sticky="w", padx=8, pady=6)
    template_combo = ttk.Combobox(
        circuit_box,
        textvariable=circuit_template_var,
        values=list(CIRCUIT_TEMPLATE_MAP.keys()),
        state="readonly",
        width=22,
    )
    template_combo.grid(row=0, column=1, sticky="w", padx=8, pady=6)

    ttk.Label(circuit_box, text="Expression").grid(row=1, column=0, sticky="nw", padx=8, pady=6)
    circuit_entry = ttk.Entry(circuit_box, textvariable=circuit_expression_var, width=32)
    circuit_entry.grid(row=1, column=1, sticky="we", padx=8, pady=6)

    ttk.Label(
        circuit_box,
        text="Use L, R, C, CPE and W tokens. Use '-' for series and p(...) for parallel, e.g. L0-R0-p(R1,CPE1).",
        wraplength=280,
        foreground="#555555",
        justify="left",
    ).grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 8))

    parameter_treatment_box = ttk.LabelFrame(controls_frame, text="Parameter treatment")
    parameter_treatment_box.pack(fill="x", pady=5)

    ttk.Label(
        parameter_treatment_box,
        text=(
            "shared = one value for all spectra, almost_shared = nearly constant, "
            "smooth = one value per spectrum with smooth evolution, independent = free per spectrum."
        ),
        wraplength=300,
        foreground="#555555",
        justify="left",
    ).pack(anchor="w", padx=8, pady=(8, 6))

    ttk.Button(parameter_treatment_box, text="Refresh parameters", command=lambda: _sync_parameter_rows()).pack(
        anchor="w",
        padx=8,
        pady=(0, 6),
    )

    parameter_rows_frame = ttk.Frame(parameter_treatment_box)
    parameter_rows_frame.pack(fill="x", padx=8, pady=(0, 8))

    strategy_box = ttk.LabelFrame(controls_frame, text="Fit strategy")
    strategy_box.pack(fill="x", pady=5)

    ttk.Label(strategy_box, text="Weighting").grid(row=0, column=0, sticky="w", padx=8, pady=6)
    weighting_combo = ttk.Combobox(
        strategy_box,
        textvariable=weighting_mode_var,
        values=WEIGHTING_OPTIONS,
        state="readonly",
        width=12,
    )
    weighting_combo.grid(row=0, column=1, sticky="w", padx=8, pady=6)

    ttk.Label(strategy_box, text="Max iterations").grid(row=1, column=0, sticky="w", padx=8, pady=6)
    max_iterations_entry = ttk.Entry(strategy_box, textvariable=max_iterations_var, width=10)
    max_iterations_entry.grid(row=1, column=1, sticky="w", padx=8, pady=6)

    ttk.Checkbutton(
        strategy_box,
        text="Use sequential seeding",
        variable=sequential_seed_var,
    ).grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=4)
    ttk.Checkbutton(
        strategy_box,
        text="Enable smoothness regularization",
        variable=smoothness_enabled_var,
    ).grid(row=3, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 8))

    validation_box = ttk.LabelFrame(controls_frame, text="Validation backend")
    validation_box.pack(fill="x", pady=5)

    try:
        _import_pyimpspec()
        pyimpspec_status = "pyimpspec backend available"
    except ValueError as exc:
        pyimpspec_status = str(exc)

    ttk.Label(
        validation_box,
        text=pyimpspec_status,
        wraplength=300,
        justify="left",
        foreground="#1f5f99" if pyimpspec_status == "pyimpspec backend available" else "#b33a3a",
    ).pack(anchor="w", padx=8, pady=8)

    zhit_toggle_button = ttk.Button(validation_box, text="Show Z-HIT settings")
    zhit_toggle_button.pack(anchor="w", padx=8, pady=(0, 8))

    zhit_settings_frame = ttk.Frame(validation_box)

    ttk.Label(zhit_settings_frame, text="Smoothing").grid(row=0, column=0, sticky="w", padx=8, pady=3)
    zhit_smoothing_combo = ttk.Combobox(
        zhit_settings_frame,
        textvariable=zhit_smoothing_var,
        values=ZHIT_SMOOTHING_OPTIONS,
        state="readonly",
        width=12,
    )
    zhit_smoothing_combo.grid(row=0, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(zhit_settings_frame, text="Interpolation").grid(row=1, column=0, sticky="w", padx=8, pady=3)
    zhit_interpolation_combo = ttk.Combobox(
        zhit_settings_frame,
        textvariable=zhit_interpolation_var,
        values=ZHIT_INTERPOLATION_OPTIONS,
        state="readonly",
        width=12,
    )
    zhit_interpolation_combo.grid(row=1, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(zhit_settings_frame, text="Window").grid(row=2, column=0, sticky="w", padx=8, pady=3)
    zhit_window_combo = ttk.Combobox(
        zhit_settings_frame,
        textvariable=zhit_window_var,
        values=ZHIT_WINDOW_OPTIONS,
        state="readonly",
        width=12,
    )
    zhit_window_combo.grid(row=2, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(zhit_settings_frame, text="Num points").grid(row=3, column=0, sticky="w", padx=8, pady=3)
    zhit_num_points_entry = ttk.Entry(zhit_settings_frame, textvariable=zhit_num_points_var, width=10)
    zhit_num_points_entry.grid(row=3, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(zhit_settings_frame, text="Poly order").grid(row=4, column=0, sticky="w", padx=8, pady=3)
    zhit_polynomial_order_entry = ttk.Entry(zhit_settings_frame, textvariable=zhit_polynomial_order_var, width=10)
    zhit_polynomial_order_entry.grid(row=4, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(zhit_settings_frame, text="Iterations").grid(row=5, column=0, sticky="w", padx=8, pady=3)
    zhit_num_iterations_entry = ttk.Entry(zhit_settings_frame, textvariable=zhit_num_iterations_var, width=10)
    zhit_num_iterations_entry.grid(row=5, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(zhit_settings_frame, text="Center").grid(row=6, column=0, sticky="w", padx=8, pady=3)
    zhit_center_entry = ttk.Entry(zhit_settings_frame, textvariable=zhit_center_var, width=10)
    zhit_center_entry.grid(row=6, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(zhit_settings_frame, text="Width").grid(row=7, column=0, sticky="w", padx=8, pady=3)
    zhit_width_entry = ttk.Entry(zhit_settings_frame, textvariable=zhit_width_var, width=10)
    zhit_width_entry.grid(row=7, column=1, sticky="w", padx=8, pady=3)

    ttk.Label(
        zhit_settings_frame,
        text="Center and width are on the log10 frequency scale used by pyimpspec.",
        wraplength=280,
        justify="left",
        foreground="#555555",
    ).grid(row=8, column=0, columnspan=2, sticky="w", padx=8, pady=(2, 8))

    def _selected_spectra_for_preview() -> list[EISSpectrum]:
        spectra: list[EISSpectrum] = []
        for row in file_rows:
            if row["selected_var"].get():
                spectra.append(row["spectrum"])
        return spectra

    def _preview_label(row_info: dict[str, object]) -> str:
        spectrum = row_info["spectrum"]
        voltage = to_float(row_info["voltage_var"].get())
        if voltage is None:
            return spectrum.source.path.name
        return f"{voltage:.6g} V | {spectrum.source.path.name}"

    def _show_ui_error(title: str, message: str) -> None:
        status_var.set(f"Error: {message}")
        try:
            messagebox.showerror(title, message, parent=win)
        except Exception:
            pass

    def _show_ui_warning(title: str, message: str) -> None:
        status_var.set(message)
        try:
            messagebox.showwarning(title, message, parent=win)
        except Exception:
            pass

    def _parameter_plot_view_size() -> tuple[int, int]:
        parameters_plot_view.update_idletasks()

        width = parameters_plot_view.winfo_width()
        height = parameters_plot_view.winfo_height()

        if width <= 1:
            width = parameters_plot_box.winfo_width() - 24
        if height <= 1:
            height = parameters_plot_box.winfo_height() - parameters_plot_scrollbar.winfo_reqheight() - 24

        if width <= 1:
            width = 840
        if height <= 1:
            height = 420

        return width, height

    def _set_parameter_plot_widget_size(width_px: int, height_px: int) -> None:
        width_px = max(1, int(round(width_px)))
        height_px = max(1, int(round(height_px)))

        parameters_canvas_widget.configure(width=width_px, height=height_px)
        parameters_plot_inner.configure(width=width_px, height=height_px)
        parameters_plot_view.itemconfigure(parameters_plot_window, width=width_px, height=height_px)
        parameters_plot_view.configure(scrollregion=(0, 0, width_px, height_px))

    def _sync_parameter_plot_scrollregion(_event=None) -> None:
        parameters_plot_view.update_idletasks()
        _set_parameter_plot_widget_size(
            parameters_canvas_widget.winfo_width(),
            parameters_canvas_widget.winfo_height(),
        )

    def _queue_parameter_plot_refresh(_event=None) -> None:
        nonlocal parameter_plot_refresh_after_id

        if parameter_plot_refresh_after_id is not None:
            try:
                win.after_cancel(parameter_plot_refresh_after_id)
            except Exception:
                pass

        parameter_plot_refresh_after_id = win.after(120, _refresh_parameter_plot_layout)

    def _toggle_parameter_plot_visibility() -> None:
        expanded = not bool(parameter_plots_expanded_var.get())
        parameter_plots_expanded_var.set(expanded)
        if expanded:
            parameter_plots_toggle_text.set("Hide parameter plots")
            parameters_plot_box.pack(fill="both", expand=True, padx=10, pady=(0, 10), after=parameters_plot_toggle_row)
            _queue_parameter_plot_refresh()
        else:
            parameter_plots_toggle_text.set("Show parameter plots")
            parameters_plot_box.pack_forget()

    def _parameter_plot_configs(result: MultiEISFitResult) -> list[ParameterTreatmentConfig]:
        if result.strategy.parameter_configs:
            return result.strategy.parameter_configs
        return [
            ParameterTreatmentConfig(
                name=name,
                element_token="",
                quantity_label=name,
                unit="-",
                treatment_mode="independent",
                initial_guess=1.0,
                lower_bound=1e-12,
                upper_bound=1e12,
            )
            for name in result.parameter_trajectories.keys()
        ]

    def _sync_parameter_table_scrollregion(_event=None) -> None:
        nonlocal parameter_table_content_width_px

        parameters_table_view.update_idletasks()
        requested_width = max(parameters_table_inner.winfo_reqwidth(), parameter_table_content_width_px)
        viewport_width = parameters_table_view.winfo_width()
        table_width = max(1, requested_width, viewport_width)
        parameters_table_view.itemconfigure(parameters_table_window, width=table_width)

        table_height = max(1, parameters_table_inner.winfo_reqheight(), parameters_table_view.winfo_height())
        parameters_table_view.configure(scrollregion=(0, 0, table_width, table_height))

    def _clear_parameter_table(message: str) -> None:
        nonlocal parameter_table_content_width_px

        parameter_table_content_width_px = 0
        parameter_table_cells.clear()
        for child in parameters_table_inner.winfo_children():
            child.destroy()
        parameters_table_inner.grid_columnconfigure(0, weight=1)
        ttk.Label(
            parameters_table_inner,
            text=message,
            wraplength=820,
            justify="left",
            foreground="#555555",
        ).grid(row=0, column=0, sticky="w")
        _sync_parameter_table_scrollregion()

    def _parameter_row_axis_text(spectrum_result: SpectrumFitResult, spectrum_index: int) -> str:
        if spectrum_result.sweep_parameter_value is not None:
            return f"{spectrum_result.sweep_parameter_value:.6g}"
        return f"S{spectrum_index + 1}"

    def _rebuild_parameter_trajectories_from_spectrum_results(result: MultiEISFitResult) -> None:
        result.parameter_trajectories = {
            config.name: [
                float(spectrum_result.parameters.get(config.name, float("nan")))
                for spectrum_result in result.spectrum_results
            ]
            for config in _parameter_plot_configs(result)
        }

    def _clear_parameter_results(message: str, reset_scroll: bool = True) -> None:
        scroll_start = 0.0 if reset_scroll else parameters_plot_view.xview()[0]
        viewport_width_px, viewport_height_px = _parameter_plot_view_size()

        parameters_fig.clear()
        parameters_fig.set_size_inches(
            viewport_width_px / parameters_fig.dpi,
            viewport_height_px / parameters_fig.dpi,
        )
        _set_parameter_plot_widget_size(viewport_width_px, viewport_height_px)
        ax = parameters_fig.add_subplot(111)
        ax.axis("off")
        ax.text(0.5, 0.5, message, ha="center", va="center", wrap=True)
        parameters_fig.tight_layout()
        parameters_canvas.draw()
        _sync_parameter_plot_scrollregion()
        parameters_plot_view.xview_moveto(scroll_start)
        _clear_parameter_table(message)

    def _render_parameter_plots(result: MultiEISFitResult, reset_scroll: bool = True) -> None:
        if not parameter_plots_expanded_var.get():
            return

        configs = _parameter_plot_configs(result)
        axis_label = result.strategy.varying_parameter_name or "Parameter"
        axis_values = [
            spectrum_result.sweep_parameter_value
            for spectrum_result in result.spectrum_results
        ]
        x_labels = [
            f"{value:.6g}" if value is not None else f"S{index + 1}"
            for index, value in enumerate(axis_values)
        ]
        x_positions = list(range(len(x_labels)))
        palette = [
            "#4c78a8",
            "#f58518",
            "#54a24b",
            "#e45756",
            "#72b7b2",
            "#b279a2",
            "#ff9da6",
            "#9d755d",
            "#bab0ab",
        ]

        scroll_start = 0.0 if reset_scroll else parameters_plot_view.xview()[0]
        viewport_width_px, viewport_height_px = _parameter_plot_view_size()
        subplot_count = len(configs)
        min_subplot_width_px = 300
        figure_width_px = max(viewport_width_px, subplot_count * min_subplot_width_px)
        figure_height_px = viewport_height_px

        parameters_fig.clear()
        parameters_fig.set_size_inches(
            figure_width_px / parameters_fig.dpi,
            figure_height_px / parameters_fig.dpi,
        )
        _set_parameter_plot_widget_size(figure_width_px, figure_height_px)
        axes = parameters_fig.subplots(nrows=1, ncols=subplot_count, squeeze=False).ravel()

        for axis_index, config in enumerate(configs):
            ax = axes[axis_index]
            values = result.parameter_trajectories.get(config.name)
            if not values:
                values = [
                    float(spectrum_result.parameters.get(config.name, float("nan")))
                    for spectrum_result in result.spectrum_results
                ]
            color = palette[axis_index % len(palette)]
            bars = ax.bar(
                x_positions,
                values,
                width=0.68,
                color=color,
                edgecolor="#333333",
                linewidth=0.6,
            )
            positive_values = [value for value in values if value > 0]
            if len(positive_values) >= 2:
                spread = max(positive_values) / min(positive_values)
                if spread >= 100.0:
                    ax.set_yscale("log")
            _apply_bar_axis_tick_aligned_limits(ax, values)
            ax.set_ylabel(config.unit if config.unit != "-" else config.name)
            ax.set_title(
                f"{config.name} ({config.quantity_label})",
                loc="left",
                fontsize=10,
                pad=6,
            )
            ax.grid(True, axis="y", alpha=0.25)
            ax.set_xticks(x_positions)
            ax.set_xticklabels(x_labels, rotation=35, ha="right")
            ax.set_xlabel(f"{axis_label} (V)" if axis_label == "Voltage" else axis_label)
            for bar, value in zip(bars, values):
                if value <= 0:
                    label_y = bar.get_height()
                else:
                    label_y = value
                ax.annotate(
                    f"{value:.3g}",
                    xy=(bar.get_x() + bar.get_width() / 2.0, label_y),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha="center",
                    va="bottom",
                    fontsize=8,
                    rotation=0,
                )

        parameters_fig.suptitle(
            f"Fitted parameters grouped by {axis_label}",
            fontsize=12,
        )
        parameters_fig.tight_layout(rect=(0.0, 0.02, 1.0, 0.96), w_pad=2.0)
        parameters_canvas.draw()
        _sync_parameter_plot_scrollregion()
        parameters_plot_view.xview_moveto(scroll_start)

    def _apply_parameter_table_value(
        *,
        spectrum_index: int,
        config: ParameterTreatmentConfig,
        value_var: tk.StringVar,
    ) -> None:
        nonlocal latest_fit_result

        if latest_fit_result is None or spectrum_index >= len(latest_fit_result.spectrum_results):
            return

        spectrum_result = latest_fit_result.spectrum_results[spectrum_index]
        previous_value = float(spectrum_result.parameters.get(config.name, float("nan")))
        raw_value = value_var.get().strip()
        parsed_value = to_float(raw_value)

        if parsed_value is None or parsed_value <= 0.0:
            value_var.set("" if not math.isfinite(previous_value) else f"{previous_value:.6g}")
            axis_label = latest_fit_result.strategy.varying_parameter_name or "Parameter"
            status_var.set(
                f"{config.name} for {axis_label} {_parameter_row_axis_text(spectrum_result, spectrum_index)} "
                "must be a positive number."
            )
            return

        spectrum_result.parameters[config.name] = float(parsed_value)
        value_var.set(f"{parsed_value:.6g}")
        _rebuild_parameter_trajectories_from_spectrum_results(latest_fit_result)
        if parameter_plots_expanded_var.get():
            _render_parameter_plots(latest_fit_result, reset_scroll=False)
        _update_nyquist_preview()

        axis_label = latest_fit_result.strategy.varying_parameter_name or "Parameter"
        status_var.set(
            f"Updated {config.name} for {axis_label} {_parameter_row_axis_text(spectrum_result, spectrum_index)}."
        )

    def _populate_editable_parameter_table(result: MultiEISFitResult) -> None:
        nonlocal parameter_table_content_width_px

        configs = _parameter_plot_configs(result)
        if not configs or not result.spectrum_results:
            _clear_parameter_table("No fitted parameter values are available yet.")
            return

        parameter_table_cells.clear()
        for child in parameters_table_inner.winfo_children():
            child.destroy()

        axis_label = result.strategy.varying_parameter_name or "Parameter"
        axis_unit = "V" if axis_label == "Voltage" else ""
        row_label_width_px = 110
        parameter_column_width_px = 160
        parameter_table_content_width_px = row_label_width_px + len(configs) * parameter_column_width_px
        parameters_table_inner.grid_columnconfigure(0, weight=0, minsize=row_label_width_px)

        ttk.Label(
            parameters_table_inner,
            text=axis_label,
            font=("", 9, "bold"),
            justify="center",
        ).grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 2))
        ttk.Label(
            parameters_table_inner,
            text=axis_unit,
            foreground="#555555",
            justify="center",
        ).grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=(0, 6))

        for col_index, config in enumerate(configs, start=1):
            parameters_table_inner.grid_columnconfigure(col_index, weight=0, minsize=parameter_column_width_px)
            ttk.Label(
                parameters_table_inner,
                text=config.name,
                font=("", 9, "bold"),
                justify="center",
            ).grid(row=0, column=col_index, sticky="nsew", padx=(0, 6), pady=(0, 2))
            ttk.Label(
                parameters_table_inner,
                text=config.unit,
                foreground="#555555",
                justify="center",
            ).grid(row=1, column=col_index, sticky="nsew", padx=(0, 6), pady=(0, 6))

        for spectrum_index, spectrum_result in enumerate(result.spectrum_results):
            grid_row = spectrum_index + 2
            ttk.Label(
                parameters_table_inner,
                text=_parameter_row_axis_text(spectrum_result, spectrum_index),
                justify="center",
            ).grid(row=grid_row, column=0, sticky="nsew", padx=(0, 6), pady=(0, 4))

            for col_index, config in enumerate(configs, start=1):
                current_value = float(spectrum_result.parameters.get(config.name, float("nan")))
                current_error = float(spectrum_result.parameter_standard_errors.get(config.name, float("nan")))
                value_var = tk.StringVar(value="" if not math.isfinite(current_value) else f"{current_value:.6g}")
                cell_frame = ttk.Frame(parameters_table_inner)
                cell_frame.grid(row=grid_row, column=col_index, sticky="ew", padx=(0, 6), pady=(0, 4))
                entry = ttk.Entry(
                    cell_frame,
                    textvariable=value_var,
                    width=max(12, len(config.name) + 2),
                )
                entry.pack(fill="x", expand=True)
                ttk.Label(
                    cell_frame,
                    text="SE +- n/a" if not math.isfinite(current_error) else f"SE +- {current_error:.3g}",
                    foreground="#666666",
                    font=("", 8),
                    justify="left",
                ).pack(anchor="w", pady=(1, 0))
                entry.bind(
                    "<Return>",
                    lambda _event, spectrum_index=spectrum_index, config=config, value_var=value_var: (
                        _apply_parameter_table_value(
                            spectrum_index=spectrum_index,
                            config=config,
                            value_var=value_var,
                        )
                    ),
                )
                entry.bind(
                    "<FocusOut>",
                    lambda _event, spectrum_index=spectrum_index, config=config, value_var=value_var: (
                        _apply_parameter_table_value(
                            spectrum_index=spectrum_index,
                            config=config,
                            value_var=value_var,
                        )
                    ),
                )
                parameter_table_cells[(spectrum_index, config.name)] = {
                    "var": value_var,
                    "entry": entry,
                    "error": current_error,
                }

        _sync_parameter_table_scrollregion()

    def _update_parameter_results(result: MultiEISFitResult | None, reset_scroll: bool = True) -> None:
        if result is None:
            _clear_parameter_results(
                "Run the simultaneous fit to populate the fitted-parameter charts.",
                reset_scroll=reset_scroll,
            )
            return

        configs = _parameter_plot_configs(result)
        if not configs or not result.spectrum_results:
            _clear_parameter_results(
                "The fit finished without fitted parameter trajectories to display.",
                reset_scroll=reset_scroll,
            )
            return

        _rebuild_parameter_trajectories_from_spectrum_results(result)
        if parameter_plots_expanded_var.get():
            _render_parameter_plots(result, reset_scroll=reset_scroll)
        _populate_editable_parameter_table(result)

    def _refresh_parameter_plot_layout() -> None:
        nonlocal parameter_plot_refresh_after_id

        parameter_plot_refresh_after_id = None
        if latest_fit_result is None:
            _clear_parameter_results(
                "Run the simultaneous fit to populate the fitted-parameter charts.",
                reset_scroll=False,
            )
            return

        if not parameter_plots_expanded_var.get():
            return

        _render_parameter_plots(latest_fit_result, reset_scroll=False)

    def _on_parameter_plot_viewport_configure(_event=None) -> None:
        if not parameter_plots_expanded_var.get():
            return
        _sync_parameter_plot_scrollregion()
        _queue_parameter_plot_refresh()

    def _on_details_notebook_tab_changed(_event=None) -> None:
        selected_tab = details_notebook.select()
        if selected_tab == str(parameters_tab) and parameter_plots_expanded_var.get():
            _queue_parameter_plot_refresh()

    def _fit_result_key(source_name: str, sweep_value: float | None) -> tuple[str, float | None]:
        if sweep_value is None:
            return source_name, None
        return source_name, round(float(sweep_value), 12)

    def _seed_trajectories_from_current_table(
        selected_dataset: MultiEISDataset,
        strategy: FitStrategyConfig,
    ) -> dict[str, list[float]]:
        if latest_fit_result is None or not latest_fit_result.spectrum_results:
            raise ValueError("Run the simultaneous fit once before using Re-fit from current table.")

        if latest_fit_result.strategy.circuit_expression != strategy.circuit_expression:
            raise ValueError(
                "The current table belongs to a different circuit expression. Run a fresh fit first."
            )

        fit_lookup = {
            _fit_result_key(spectrum_result.source_name, spectrum_result.sweep_parameter_value): spectrum_result
            for spectrum_result in latest_fit_result.spectrum_results
        }
        seed_trajectories = {config.name: [] for config in strategy.parameter_configs}

        for spectrum in selected_dataset.spectra:
            fit_key = _fit_result_key(spectrum.source.path.name, spectrum.sweep_parameter_value)
            spectrum_result = fit_lookup.get(fit_key)
            if spectrum_result is None:
                raise ValueError(
                    "The current table no longer matches the selected spectra/voltages. Run a fresh fit first."
                )

            for config in strategy.parameter_configs:
                if config.name not in spectrum_result.parameters:
                    raise ValueError(
                        f"The current table does not contain a seed value for {config.name}. Run a fresh fit first."
                    )
                value = float(spectrum_result.parameters[config.name])
                if not math.isfinite(value) or value <= 0.0:
                    raise ValueError(
                        f"{config.name} for {spectrum.source.path.name} is not a valid positive seed value."
                    )
                seed_trajectories[config.name].append(_clip_parameter_value(config, value))

        return seed_trajectories

    def _current_fit_signature() -> tuple:
        file_signature = []
        for row in file_rows:
            voltage = to_float(row["voltage_var"].get())
            file_signature.append(
                (
                    row["spectrum"].source.path.name,
                    bool(row["selected_var"].get()),
                    voltage if voltage is not None else row["voltage_var"].get().strip(),
                )
            )

        parameter_signature = []
        for row in parameter_rows:
            parameter_signature.append(
                (
                    row["name"],
                    row["treatment_var"].get(),
                    row["guess_var"].get().strip(),
                    row["lower_var"].get().strip(),
                    row["upper_var"].get().strip(),
                    row["smooth_var"].get().strip(),
                    row["similarity_var"].get().strip(),
                )
            )

        return (
            tuple(file_signature),
            varying_parameter_var.get(),
            circuit_expression_var.get().strip(),
            weighting_mode_var.get(),
            bool(sequential_seed_var.get()),
            bool(smoothness_enabled_var.get()),
            max_iterations_var.get().strip(),
            tuple(parameter_signature),
        )

    def _update_nyquist_preview() -> None:
        nonlocal latest_fit_result, latest_fit_signature
        preview_fig.clear()
        ax = preview_fig.add_subplot(111)

        selected_rows = [row for row in file_rows if row["selected_var"].get()]
        if not selected_rows:
            ax.text(0.5, 0.5, "No selected spectra.", ha="center", va="center", transform=ax.transAxes)
            ax.set_axis_off()
            preview_canvas.draw_idle()
            return

        fit_lookup: dict[tuple[str, float | None], SpectrumFitResult] = {}
        circuit_tree = None
        show_fit_overlay = latest_fit_result is not None and latest_fit_signature == _current_fit_signature()
        if show_fit_overlay:
            try:
                circuit_tree = _parse_circuit_tree_from_expression(circuit_expression_var.get().strip())
                fit_lookup = {
                    _fit_result_key(result.source_name, result.sweep_parameter_value): result
                    for result in latest_fit_result.spectrum_results
                }
            except Exception:
                show_fit_overlay = False

        np, _least_squares = _import_fit_backend()
        data_x_values: list[float] = []
        data_y_values: list[float] = []
        fit_overlay_specs: list[tuple[object, object, str, str]] = []

        for plot_index, row in enumerate(selected_rows):
            spectrum = row["spectrum"]
            color = f"C{plot_index}"
            data_x_values.extend(float(value) for value in spectrum.z_real_ohm)
            data_y_values.extend(float(-value) for value in spectrum.z_imag_ohm)
            ax.plot(
                spectrum.z_real_ohm,
                [-value for value in spectrum.z_imag_ohm],
                linestyle="None",
                marker="o",
                markersize=3,
                markerfacecolor="none",
                markeredgecolor=color,
                color=color,
                label=f"{_preview_label(row)} data",
            )

            if show_fit_overlay:
                voltage = to_float(row["voltage_var"].get())
                fit_result = fit_lookup.get(_fit_result_key(spectrum.source.path.name, voltage))
                if fit_result is not None and fit_result.parameters:
                    try:
                        omega = 2.0 * np.pi * np.asarray(spectrum.frequency_hz, dtype=float)
                        z_model = _evaluate_circuit_impedance(circuit_tree, fit_result.parameters, omega)
                    except Exception as exc:
                        status_var.set(
                            f"No se pudo actualizar el overlay para {spectrum.source.path.name}: {exc}"
                        )
                        continue

                    fit_overlay_specs.append(
                        (
                            z_model.real,
                            -z_model.imag,
                            f"{_preview_label(row)} fit",
                            color,
                        )
                    )

        if data_x_values and data_y_values:
            finite_x = [value for value in data_x_values if math.isfinite(value)]
            finite_y = [value for value in data_y_values if math.isfinite(value)]
            if finite_x and finite_y:
                x_min = min(finite_x)
                x_max = max(finite_x)
                y_min = min(finite_y)
                y_max = max(finite_y)

                x_padding = max((x_max - x_min) * 0.05, 1e-9)
                y_padding = max((y_max - y_min) * 0.05, 1e-9)

                if math.isclose(x_min, x_max):
                    x_min -= x_padding
                    x_max += x_padding
                else:
                    x_min -= x_padding
                    x_max += x_padding

                if math.isclose(y_min, y_max):
                    y_min -= y_padding
                    y_max += y_padding
                else:
                    y_min -= y_padding
                    y_max += y_padding

                ax.set_xlim(x_min, x_max)
                ax.set_ylim(y_min, y_max)

        for x_fit, y_fit, fit_label, color in fit_overlay_specs:
            ax.plot(
                x_fit,
                y_fit,
                linestyle="-",
                marker=None,
                linewidth=1.5,
                color=color,
                label=fit_label,
                scalex=False,
                scaley=False,
            )

        ax.set_title("Nyquist")
        ax.set_xlabel("Zreal")
        ax.set_ylabel("-Zimag")
        ax.grid(True, alpha=0.25)
        ax.set_aspect("equal", adjustable="box")
        handles, labels = ax.get_legend_handles_labels()
        if handles:
            make_legend_draggable(
                ax.legend(
                    handles,
                    labels,
                    fontsize=8,
                    loc="upper left",
                    bbox_to_anchor=(1.02, 1.0),
                    borderaxespad=0.0,
                )
            )
            preview_fig.tight_layout(rect=(0.0, 0.0, 0.78, 1.0))
        else:
            preview_fig.tight_layout()
        preview_canvas.draw_idle()

    def _sync_zhit_visibility() -> None:
        if zhit_expanded_var.get():
            zhit_settings_frame.pack(fill="x", padx=0, pady=(0, 8))
            zhit_toggle_button.configure(text="Hide Z-HIT settings")
        else:
            zhit_settings_frame.pack_forget()
            zhit_toggle_button.configure(text="Show Z-HIT settings")

    def _toggle_zhit_visibility() -> None:
        zhit_expanded_var.set(not zhit_expanded_var.get())
        _sync_zhit_visibility()

    def _collect_zhit_config() -> dict[str, object]:
        return {
            "smoothing": zhit_smoothing_var.get(),
            "interpolation": zhit_interpolation_var.get(),
            "window": zhit_window_var.get(),
            "num_points": _positive_int(zhit_num_points_var.get(), "Z-HIT num points"),
            "polynomial_order": _positive_int(zhit_polynomial_order_var.get(), "Z-HIT polynomial order"),
            "num_iterations": _positive_int(zhit_num_iterations_var.get(), "Z-HIT iterations"),
            "center": _required_float(zhit_center_var.get(), "Z-HIT center"),
            "width": _positive_float(zhit_width_var.get(), "Z-HIT width"),
        }

    def _sync_parameter_rows() -> None:
        existing_values = {
            row["name"]: {
                "treatment_mode": row["treatment_var"].get(),
                "initial_guess": row["guess_var"].get(),
                "lower_bound": row["lower_var"].get(),
                "upper_bound": row["upper_var"].get(),
                "smoothness_weight": row["smooth_var"].get(),
                "similarity_weight": row["similarity_var"].get(),
            }
            for row in parameter_rows
        }

        for child in parameter_rows_frame.winfo_children():
            child.destroy()
        parameter_rows.clear()

        expression = circuit_expression_var.get().strip()
        if not expression:
            ttk.Label(
                parameter_rows_frame,
                text="Ingrese una expresion de circuito para definir los parametros.",
                foreground="#555555",
                justify="left",
                wraplength=300,
            ).pack(anchor="w")
            return

        try:
            defaults = _default_parameter_configs(expression)
        except ValueError as exc:
            ttk.Label(
                parameter_rows_frame,
                text=f"Circuito invalido: {exc}",
                foreground="#b33a3a",
                justify="left",
                wraplength=300,
            ).pack(anchor="w")
            return

        for config in defaults:
            saved = existing_values.get(config.name)
            treatment_var = tk.StringVar(value=saved["treatment_mode"] if saved else config.treatment_mode)
            guess_var = tk.StringVar(value=saved["initial_guess"] if saved else f"{config.initial_guess:.6g}")
            lower_var = tk.StringVar(value=saved["lower_bound"] if saved else f"{config.lower_bound:.6g}")
            upper_var = tk.StringVar(value=saved["upper_bound"] if saved else f"{config.upper_bound:.6g}")
            smooth_var = tk.StringVar(
                value=saved["smoothness_weight"] if saved else f"{config.smoothness_weight:.6g}"
            )
            similarity_var = tk.StringVar(
                value=saved["similarity_weight"] if saved else f"{config.similarity_weight:.6g}"
            )

            row_frame = ttk.Frame(parameter_rows_frame)
            row_frame.pack(fill="x", pady=(0, 8))

            ttk.Label(
                row_frame,
                text=f"{config.name} ({config.element_token}, {config.unit})",
                font=("", 9, "bold"),
                justify="left",
            ).grid(row=0, column=0, sticky="w")
            ttk.Label(row_frame, text="Mode").grid(row=0, column=1, sticky="e", padx=(10, 4))
            mode_combo = ttk.Combobox(
                row_frame,
                textvariable=treatment_var,
                values=PARAMETER_TREATMENT_OPTIONS,
                state="readonly",
                width=14,
            )
            mode_combo.grid(row=0, column=2, sticky="w")

            ttk.Label(row_frame, text="Guess").grid(row=1, column=0, sticky="w", pady=(4, 0))
            guess_entry = ttk.Entry(row_frame, textvariable=guess_var, width=10)
            guess_entry.grid(row=1, column=1, sticky="w", pady=(4, 0))
            ttk.Label(row_frame, text="Lower").grid(row=1, column=2, sticky="e", padx=(10, 4), pady=(4, 0))
            lower_entry = ttk.Entry(row_frame, textvariable=lower_var, width=10)
            lower_entry.grid(row=1, column=3, sticky="w", pady=(4, 0))
            ttk.Label(row_frame, text="Upper").grid(row=1, column=4, sticky="e", padx=(10, 4), pady=(4, 0))
            upper_entry = ttk.Entry(row_frame, textvariable=upper_var, width=10)
            upper_entry.grid(row=1, column=5, sticky="w", pady=(4, 0))

            ttk.Label(row_frame, text="Smooth").grid(row=2, column=0, sticky="w", pady=(4, 0))
            smooth_entry = ttk.Entry(row_frame, textvariable=smooth_var, width=10)
            smooth_entry.grid(row=2, column=1, sticky="w", pady=(4, 0))
            ttk.Label(row_frame, text="Similarity").grid(row=2, column=2, sticky="e", padx=(10, 4), pady=(4, 0))
            similarity_entry = ttk.Entry(row_frame, textvariable=similarity_var, width=10)
            similarity_entry.grid(row=2, column=3, sticky="w", pady=(4, 0))

            ttk.Separator(parameter_rows_frame, orient="horizontal").pack(fill="x", pady=(0, 6))

            def _sync_row_state(
                _event=None,
                *,
                mode_var=treatment_var,
                smooth_widget=smooth_entry,
                similarity_widget=similarity_entry,
            ) -> None:
                mode = mode_var.get()
                smooth_widget.configure(state="normal" if mode in ("smooth", "almost_shared") else "disabled")
                similarity_widget.configure(state="normal" if mode == "almost_shared" else "disabled")

            mode_combo.bind("<<ComboboxSelected>>", _sync_row_state)
            _sync_row_state()

            parameter_rows.append(
                {
                    "name": config.name,
                    "element_token": config.element_token,
                    "quantity_label": config.quantity_label,
                    "unit": config.unit,
                    "treatment_var": treatment_var,
                    "guess_var": guess_var,
                    "lower_var": lower_var,
                    "upper_var": upper_var,
                    "smooth_var": smooth_var,
                    "similarity_var": similarity_var,
                }
            )

    def _collect_parameter_configs() -> list[ParameterTreatmentConfig]:
        if not parameter_rows:
            raise ValueError("No se detectaron parametros de circuito para el fitting.")

        configs: list[ParameterTreatmentConfig] = []
        for row in parameter_rows:
            name = row["name"]
            treatment_mode = row["treatment_var"].get()
            if treatment_mode not in PARAMETER_TREATMENT_OPTIONS:
                raise ValueError(f"{name}: modo de tratamiento no soportado.")

            lower_bound = _positive_float(row["lower_var"].get(), f"{name} lower bound")
            upper_bound = _positive_float(row["upper_var"].get(), f"{name} upper bound")
            initial_guess = _positive_float(row["guess_var"].get(), f"{name} initial guess")

            if upper_bound <= lower_bound:
                raise ValueError(f"{name}: upper bound debe ser mayor que lower bound.")
            if not (lower_bound <= initial_guess <= upper_bound):
                raise ValueError(f"{name}: initial guess debe quedar dentro de los bounds.")

            smoothness_weight = 0.0
            similarity_weight = 0.0
            if treatment_mode in ("smooth", "almost_shared"):
                smoothness_weight = _nonnegative_float(row["smooth_var"].get(), f"{name} smoothness weight")
            if treatment_mode == "almost_shared":
                similarity_weight = _nonnegative_float(row["similarity_var"].get(), f"{name} similarity weight")

            configs.append(
                ParameterTreatmentConfig(
                    name=name,
                    element_token=row["element_token"],
                    quantity_label=row["quantity_label"],
                    unit=row["unit"],
                    treatment_mode=treatment_mode,
                    initial_guess=initial_guess,
                    lower_bound=lower_bound,
                    upper_bound=upper_bound,
                    smoothness_weight=smoothness_weight,
                    similarity_weight=similarity_weight,
                )
            )
        return configs

    def _select_all():
        for row in file_rows:
            row["selected_var"].set(True)
        _update_nyquist_preview()

    def _clear_all():
        for row in file_rows:
            row["selected_var"].set(False)
        _update_nyquist_preview()

    def _restore_detected_voltage():
        for row in file_rows:
            spectrum = row["spectrum"]
            row["voltage_var"].set(_default_voltage_text(spectrum))
        _update_nyquist_preview()

    ttk.Button(files_button_row, text="Select all", command=_select_all).pack(side="left", padx=(0, 6))
    ttk.Button(files_button_row, text="Clear all", command=_clear_all).pack(side="left", padx=(0, 6))
    ttk.Button(files_button_row, text="Use detected Vdc", command=_restore_detected_voltage).pack(side="left")

    def _selected_voltage_pairs() -> list[tuple[int, float]]:
        selections: list[tuple[int, float]] = []
        for row in file_rows:
            if not row["selected_var"].get():
                continue
            raw_voltage = row["voltage_var"].get()
            voltage = to_float(raw_voltage)
            spectrum = row["spectrum"]
            if voltage is None:
                raise ValueError(f"{spectrum.source.path.name}: ingrese un valor de voltaje valido.")
            selections.append((row["index"], voltage))
        if len(selections) < 2:
            raise ValueError("Debe seleccionar al menos dos archivos EIS para multifit.")
        return selections

    def _validate_configuration() -> tuple[MultiEISDataset, FitStrategyConfig]:
        selected_dataset = build_selected_dataset(
            dataset=dataset,
            selections=_selected_voltage_pairs(),
            varying_parameter_name=varying_parameter_var.get(),
        )
        validate_frequency_grid(selected_dataset)
        circuit_expression = _validate_circuit_expression(circuit_expression_var.get())
        circuit_formula = _circuit_formula_preview(circuit_expression)
        parameter_configs = _collect_parameter_configs()

        strategy = FitStrategyConfig(
            circuit_expression=circuit_expression,
            circuit_formula=circuit_formula,
            parameter_configs=parameter_configs,
            varying_parameter_name=varying_parameter_var.get(),
            weighting_mode=weighting_mode_var.get(),
            sequential_seed=sequential_seed_var.get(),
            smoothness_enabled=smoothness_enabled_var.get(),
            max_iterations=_positive_int(max_iterations_var.get(), "Max iterations"),
        )
        return selected_dataset, strategy

    def _selection_summary(selected_dataset: MultiEISDataset, strategy: FitStrategyConfig) -> str:
        lines = [
            f"Selected spectra: {selected_dataset.count}",
            f"Varying parameter: {strategy.varying_parameter_name}",
            f"Frequency grid: {_frequency_grid_summary(selected_dataset)}",
            f"Circuit: {strategy.circuit_expression}",
            f"Formula: {strategy.circuit_formula}",
            f"Weighting: {strategy.weighting_mode}",
            f"Sequential seed: {'Yes' if strategy.sequential_seed else 'No'}",
            f"Smoothness regularization: {'Yes' if strategy.smoothness_enabled else 'No'}",
            f"Max iterations: {strategy.max_iterations}",
            f"Logical parameters: {len(strategy.parameter_configs)}",
            "",
            "Parameter treatment:",
        ]
        lines.extend(_parameter_treatment_summary_lines(strategy.parameter_configs))
        lines.extend(
            [
                "",
            "Selected files ordered by parameter value:",
            ]
        )
        for spectrum in selected_dataset.spectra:
            lines.append(
                f"  {spectrum.sweep_parameter_value:.6g} V  |  {spectrum.source.path.name}"
            )
        return "\n".join(lines)

    def _validate_and_preview():
        try:
            status_var.set("Validando seleccion...")
            win.update_idletasks()
            selected_dataset, strategy = _validate_configuration()
        except ValueError as exc:
            _show_ui_error("MultiFit - Validation", str(exc))
            return
        except Exception as exc:
            _show_ui_error(
                "MultiFit - Validation",
                f"Error inesperado durante la validacion: {type(exc).__name__}: {exc}",
            )
            return

        _set_text(summary_text, _selection_summary(selected_dataset, strategy))
        _set_text(
            notes_text,
            "\n".join(
                [
                    "Validation passed.",
                    "",
                    "The circuit expression is now parsed and expanded into a symbolic impedance formula.",
                    "Per-parameter treatment modes are captured and ready for the simultaneous optimizer.",
                    "",
                    "You can now run the simultaneous fit directly from this window.",
                    "The optimizer works in log space and can combine shared, almost_shared,",
                    "smooth and independent parameter trajectories across the voltage sequence.",
                ]
            ),
        )
        details_notebook.select(summary_tab)
        status_var.set("Seleccion validada. La grilla de frecuencias coincide en todos los archivos seleccionados.")

    def _run_pyimpspec_validation():
        try:
            status_var.set("Ejecutando validacion pyimpspec...")
            win.update_idletasks()
            selected_dataset, strategy = _validate_configuration()
            report = run_pyimpspec_validation_with_config(
                selected_dataset,
                zhit_config=_collect_zhit_config(),
            )
        except ValueError as exc:
            _show_ui_error("MultiFit - pyimpspec validation", str(exc))
            return
        except Exception as exc:
            _show_ui_error(
                "MultiFit - pyimpspec validation",
                f"Error inesperado durante la validacion: {type(exc).__name__}: {exc}",
            )
            return

        _set_text(summary_text, _selection_summary(selected_dataset, strategy))
        _set_text(validation_text, report.text)
        details_notebook.select(validation_tab)
        status_var.set(
            "Validacion pyimpspec finalizada. "
            f"KK OK: {report.kk_success_count}/{selected_dataset.count}, "
            f"Z-HIT OK: {report.zhit_success_count}/{selected_dataset.count}."
        )

    def _prepare_fit_session(use_current_table_seed: bool = False):
        nonlocal latest_fit_result, latest_fit_signature
        try:
            selected_dataset, strategy = _validate_configuration()
            seed_trajectories = None
            if use_current_table_seed:
                seed_trajectories = _seed_trajectories_from_current_table(selected_dataset, strategy)
                strategy = replace(strategy, sequential_seed=False)
                status_var.set("Re-ejecutando fitting desde la tabla actual...")
            else:
                status_var.set("Ejecutando fitting simultaneo...")
            win.update_idletasks()
            result = fit_dataset(selected_dataset, strategy, seed_trajectories=seed_trajectories)
        except ValueError as exc:
            _show_ui_error("MultiFit - Simultaneous fit", str(exc))
            return
        except Exception as exc:
            _show_ui_error(
                "MultiFit - Simultaneous fit",
                f"Error inesperado durante el fitting: {type(exc).__name__}: {exc}",
            )
            return

        summary_lines = [
            _selection_summary(selected_dataset, strategy),
            "",
            f"Fit success: {'Yes' if result.success else 'No'}",
            f"Message: {result.message}",
            f"Objective: {result.objective_value:.6g}" if result.objective_value is not None else "Objective: n/a",
            f"Function evaluations: {result.nfev}" if result.nfev is not None else "Function evaluations: n/a",
            "Parameter SE: Hessian inverse in log-parameter space.",
            (
                "Initial seed source: current editable parameter table."
                if use_current_table_seed
                else "Initial seed source: standard fit setup."
            ),
            "",
            "Per-spectrum fit quality:",
        ]
        for spectrum_result in result.spectrum_results:
            summary_lines.append(
                (
                    f"  {spectrum_result.sweep_parameter_value:.6g} V | "
                    f"{spectrum_result.source_name} | "
                    f"mse={spectrum_result.objective_value:.6g}"
                )
            )
        _set_text(summary_text, "\n".join(summary_lines))

        notes_lines = [
            "Simultaneous fit notes:",
            "",
        ]
        notes_lines.extend(result.notes)
        _set_text(notes_text, "\n".join(notes_lines))
        latest_fit_result = result
        latest_fit_signature = _current_fit_signature()
        details_notebook.select(parameters_tab)
        win.update_idletasks()
        _update_parameter_results(result)
        _update_nyquist_preview()
        if result.success:
            if use_current_table_seed:
                status_var.set("Re-fit desde la tabla actual finalizado. Resultado: OK.")
            else:
                status_var.set("Fitting simultaneo finalizado. Resultado: OK.")
        else:
            _show_ui_warning(
                "MultiFit - Simultaneous fit",
                f"El fitting termino con advertencias:\n\n{result.message}",
            )

    def _reset_form():
        nonlocal latest_fit_result, latest_fit_signature
        varying_parameter_var.set("Voltage")
        weighting_mode_var.set("Modulus")
        sequential_seed_var.set(True)
        smoothness_enabled_var.set(True)
        max_iterations_var.set("500")
        zhit_expanded_var.set(False)
        zhit_smoothing_var.set("modsinc")
        zhit_interpolation_var.set("makima")
        zhit_window_var.set("auto")
        zhit_num_points_var.set("3")
        zhit_polynomial_order_var.set("2")
        zhit_num_iterations_var.set("3")
        zhit_center_var.set("1.5")
        zhit_width_var.set("3.0")
        circuit_template_var.set("Randles")
        circuit_expression_var.set(CIRCUIT_TEMPLATE_MAP["Randles"])
        latest_fit_result = None
        latest_fit_signature = None
        _sync_parameter_rows()
        _select_all()
        _restore_detected_voltage()
        _sync_zhit_visibility()
        _set_text(summary_text, "No validation has been run yet.")
        _set_text(validation_text, "No pyimpspec validation has been run yet.")
        _set_text(notes_text, "Use Validate selection or run pyimpspec validation before starting the simultaneous fit.")
        _clear_parameter_results("Run the simultaneous fit to populate the fitted-parameter charts.")
        status_var.set("Valores restaurados.")

    def _apply_template(*_args):
        template = circuit_template_var.get()
        expression = CIRCUIT_TEMPLATE_MAP.get(template, "")
        if expression:
            circuit_expression_var.set(expression)
            _sync_parameter_rows()

    template_combo.bind("<<ComboboxSelected>>", _apply_template)
    circuit_entry.bind("<Return>", lambda _event: _sync_parameter_rows())
    circuit_entry.bind("<FocusOut>", lambda _event: _sync_parameter_rows())
    zhit_toggle_button.configure(command=_toggle_zhit_visibility)
    details_notebook.bind("<<NotebookTabChanged>>", _on_details_notebook_tab_changed)
    parameters_plot_inner.bind("<Configure>", _sync_parameter_plot_scrollregion)
    parameters_plot_view.bind("<Configure>", _on_parameter_plot_viewport_configure)
    parameters_table_inner.bind("<Configure>", _sync_parameter_table_scrollregion)
    parameters_table_view.bind("<Configure>", _sync_parameter_table_scrollregion)

    buttons_row = ttk.Frame(controls_frame)
    buttons_row.pack(fill="x", pady=(12, 0))
    ttk.Button(buttons_row, text="Validate selection", command=_validate_and_preview).pack(side="left", padx=(0, 6))
    ttk.Button(buttons_row, text="Run pyimpspec validation", command=_run_pyimpspec_validation).pack(side="left", padx=(0, 6))
    ttk.Button(buttons_row, text="Run simultaneous fit", command=_prepare_fit_session).pack(side="left", padx=(0, 6))
    ttk.Button(
        buttons_row,
        text="Re-fit from current table",
        command=lambda: _prepare_fit_session(use_current_table_seed=True),
    ).pack(side="left", padx=(0, 6))
    ttk.Button(buttons_row, text="Reset", command=_reset_form).pack(side="left")

    ttk.Label(
        controls_frame,
        textvariable=status_var,
        wraplength=520,
        justify="left",
    ).pack(anchor="w", fill="x", pady=(12, 0))

    _set_text(summary_text, "No validation has been run yet.")
    _set_text(validation_text, "No pyimpspec validation has been run yet.")
    _set_text(notes_text, "Use Validate selection or run pyimpspec validation before starting the simultaneous fit.")
    _clear_parameter_results("Run the simultaneous fit to populate the fitted-parameter charts.")
    _sync_parameter_rows()
    _sync_zhit_visibility()
    _update_nyquist_preview()

    if created_root:
        win.mainloop()


def export_folder(input_dir: Path, output_dir: Path) -> list[Path]:
    """Reserved for future fit reports/results exports."""
    _ = output_dir
    dataset = load_dataset(input_dir)
    validate_dataset(dataset)
    return []


def run_pipeline(
    input_dir: Path,
    output_dir: Path,
    selected_options: Iterable[str] | None = None,
) -> list[Path]:
    _ = output_dir
    _ = selected_options

    open_multifit_window(input_dir=Path(input_dir))
    return []
