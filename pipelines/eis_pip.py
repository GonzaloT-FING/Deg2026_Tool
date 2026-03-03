"""EIS (.DTA) -> Excel (.xlsx) exporter (Gamry Potentiostatic EIS)

Version 1 goals (as discussed):
  - Find all EIS files in a folder by filename containing 'EISPOT'
  - Parse required metadata fields
  - Parse the ZCURVE table
  - Export ONE .xlsx per input file with two sheets:
        1) Metadata
        2) Data

This script is intentionally simple and heavily commented for learning.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


# --- What we want to extract (your spec) ------------------------------------

META_FIELDS = [
    ("TITLE", "Title"),
    ("DATE", "Date"),
    ("TIME", "Time"),
    ("VDC", "Vdc"),
    ("FREQINIT", "freqinit"),
    ("FREQFINAL", "freqfinal"),
    ("PTSPERDEC", "ptsperdec"),
    ("VAC", "vac"),
    ("AREA", "area"),
]

DATA_MAP = {
    "Pt": "Pt",
    "Freq": "Frequency",
    "Zreal": "Zreal",
    "Zimag": "Zimag",
    "Zsig": "Zsig",
    "Zmod": "Zmod",
    "Zphz": "Zphz",
    "Idc": "Idc",
    "Vdc": "Vdc",
    "Temp": "Temperature",
}


def to_float(val: str) -> float | None:
    """Convert Gamry-style numbers (decimal comma, scientific notation) to float."""
    s = val.strip()
    if not s:
        return None
    # Gamry exports commonly use ',' as decimal separator.
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


@dataclass
class ParsedDTA:
    meta: dict[str, str]
    header: list[str]
    rows: list[list[str]]


def parse_gamry_dta(path: Path) -> ParsedDTA:
    """Parse a Gamry .DTA file that contains a ZCURVE table.

    Strategy:
      1) Read the file as text (latin-1 is safe for the ° symbol in headers)
      2) Collect metadata key/value pairs until we reach 'ZCURVE\tTABLE'
      3) Read the column header line (starts with 'Pt')
      4) Skip the units line (starts with '#')
      5) Read data rows (first token is an integer Pt)
    """

    text = path.read_text(encoding="latin-1", errors="replace")
    lines = text.splitlines()

    meta: dict[str, str] = {}
    table_started = False
    header: list[str] | None = None
    data_rows: list[list[str]] = []

    for line in lines:
        if not table_started:
            if line.startswith("ZCURVE") and "TABLE" in line:
                table_started = True
                continue

            parts = line.split("\t")
            if len(parts) >= 3 and parts[0].strip():
                key = parts[0].strip()
                val = parts[2].strip()
                meta[key] = val

        else:
            if not line.strip():
                continue

            parts = line.strip().split("\t")

            if header is None:
                if parts and parts[0] == "Pt":
                    header = parts
                continue

            # Units row starts with '#'
            if parts and parts[0] == "#":
                continue

            # Data rows start with an integer Pt index
            if parts and re.fullmatch(r"-?\d+", parts[0]):
                data_rows.append(parts)

    if header is None:
        raise ValueError(f"No data header found in {path.name} (expected 'Pt ...')")

    return ParsedDTA(meta=meta, header=header, rows=data_rows)


def export_to_xlsx(parsed: ParsedDTA, out_path: Path) -> None:
    """Create an .xlsx with Metadata + Data sheets."""

    wb = Workbook()
    wb.remove(wb.active)  # remove the default sheet

    ws_meta = wb.create_sheet("Metadata")
    ws_data = wb.create_sheet("Data")

    # --- Metadata sheet ------------------------------------------------------
    ws_meta["A1"] = "Field"
    ws_meta["B1"] = "Value"
    ws_meta["A1"].font = ws_meta["B1"].font = Font(bold=True)
    ws_meta.freeze_panes = "A2"

    row = 2
    numeric_keys = {"VDC", "FREQINIT", "FREQFINAL", "PTSPERDEC", "VAC", "AREA"}
    for key, label in META_FIELDS:
        ws_meta.cell(row=row, column=1, value=label)
        raw_val = parsed.meta.get(key, "")

        if key in numeric_keys:
            num = to_float(raw_val)
            ws_meta.cell(row=row, column=2, value=num if num is not None else raw_val)
        else:
            ws_meta.cell(row=row, column=2, value=raw_val)
        row += 1

    # --- Data sheet ----------------------------------------------------------
    col_idx = {name: i for i, name in enumerate(parsed.header)}
    selected_cols = [c for c in DATA_MAP if c in col_idx]
    out_headers = [DATA_MAP[c] for c in selected_cols]

    for j, h in enumerate(out_headers, start=1):
        cell = ws_data.cell(row=1, column=j, value=h)
        cell.font = Font(bold=True)
    ws_data.freeze_panes = "A2"

    for r_i, parts in enumerate(parsed.rows, start=2):
        for c_i, col_name in enumerate(selected_cols, start=1):
            raw = parts[col_idx[col_name]] if col_idx[col_name] < len(parts) else ""
            num = to_float(raw)
            ws_data.cell(row=r_i, column=c_i, value=num if num is not None else raw)

    # --- Light formatting: widths + vertical alignment ----------------------
    for ws in (ws_meta, ws_data):
        for col in range(1, ws.max_column + 1):
            maxlen = 0
            for r in range(1, min(ws.max_row, 50) + 1):  # sample first 50 rows
                v = ws.cell(row=r, column=col).value
                if v is None:
                    continue
                maxlen = max(maxlen, len(str(v)))
            ws.column_dimensions[get_column_letter(col)].width = min(max(10, maxlen + 2), 45)

        for row_cells in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                      min_col=1, max_col=ws.max_column):
            for cell in row_cells:
                cell.alignment = Alignment(vertical="top")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def export_folder(input_dir: Path, output_dir: Path) -> list[Path]:
    """Find EISPOT*.DTA in input_dir and export each to output_dir."""
    output_dir.mkdir(parents=True, exist_ok=True)

    dta_files = sorted([p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() == ".dta" and "EISPOT" in p.name])
    if not dta_files:
        return []

    exported: list[Path] = []
    for f in dta_files:
        parsed = parse_gamry_dta(f)
        out_path = output_dir / f"{f.stem}.xlsx"
        export_to_xlsx(parsed, out_path)
        exported.append(out_path)

    return exported


def main() -> None:
    # Example CLI usage (edit these paths):
    input_dir = Path(r"C:\\path\\to\\your\\input")
    output_dir = Path(r"C:\\path\\to\\your\\output")

    exported = export_folder(input_dir, output_dir)
    print(f"Exported {len(exported)} file(s)")
    for p in exported:
        print(" -", p)


if __name__ == "__main__":
    main()
