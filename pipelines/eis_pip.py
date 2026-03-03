"""EIS (.DTA) -> Excel (.xlsx) exporter for Gamry Potentiostatic EIS.

What this version does:
  - Finds all .DTA files whose filename contains 'EISPOT'
  - Parses selected metadata fields
  - Parses the ZCURVE table
  - Exports ONE .xlsx per input file with two sheets:
        1) Metadata  -> Campo / Valor / Unidad
        2) Data      -> headers row, units row, then numeric data

This version is written to match the real structure of the uploaded Gamry files.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

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
    """Gamry ZCURVE rows usually start with a leading tab.

    Example after split:
        ['', 'Pt', 'Time', 'Freq', ...]
    We remove the first empty item if present.
    """
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
    """Extract a clean metadata unit from the descriptive text.

    Real examples from the uploaded files:
      VDC       ...   DC Voltage (V)
      FREQINIT  ...   Initial Freq. (Hz)
      VAC       ...   AC Voltage (mV rms)
      AREA      ...   Sample Area (cm^2)
    """
    unit = _extract_parenthesized_unit(description)
    if unit:
        return unit

    if key == "PTSPERDEC":
        return "points/decade"

    # Non-physical fields stay blank.
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

        # Table section ------------------------------------------------------
        if not line.strip():
            continue

        parts = _drop_leading_blank(line.rstrip("\r\n").split("\t"))
        if not parts:
            continue

        # Header row
        if header is None:
            if parts[0] == "Pt":
                header = parts
            continue

        # Units row
        if parts[0] == "#":
            units = parts
            continue

        # Data rows start with the integer point index
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
# Export
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
# Folder export
# ---------------------------------------------------------------------------

def export_folder(input_dir: Path, output_dir: Path) -> list[Path]:
    """Find all EISPOT .DTA files in input_dir and export them to output_dir."""
    output_dir.mkdir(parents=True, exist_ok=True)

    dta_files = sorted(
        [
            p for p in input_dir.iterdir()
            if p.is_file() and p.suffix.lower() == ".dta" and "EISPOT" in p.name
        ]
    )

    if not dta_files:
        return []

    exported: list[Path] = []
    for dta_file in dta_files:
        parsed = parse_gamry_dta(dta_file)
        out_path = output_dir / f"{dta_file.stem}.xlsx"
        export_to_xlsx(parsed, out_path)
        exported.append(out_path)

    return exported


def main() -> None:
    """Manual standalone test."""
    input_dir = Path(r"C:\\path\\to\\your\\input")
    output_dir = Path(r"C:\\path\\to\\your\\output")

    exported = export_folder(input_dir, output_dir)
    print(f"Exported {len(exported)} file(s)")
    for path in exported:
        print(" -", path)


if __name__ == "__main__":
    main()
