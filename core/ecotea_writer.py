"""
ecotea_writer.py — Writes PowerRecord lists into an EcoTEA Excel template.

If the template format changes, update POWER_COL_ORDER and the header rows
preserved at the top of write_power_sheet().
"""

import io

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from core.base_model import PowerRecord, MISSING


# ── Column mapping: PowerRecord field → Excel column index (1-based) ─────────
# Based on EcoTEA Power sheet structure confirmed from reference file.

POWER_COL_ORDER = [
    'wp6_title',           # A  (1)
    'data_owner',          # B  (2)
    'data_provider',       # C  (3)
    'data_source',         # D  (4)
    'data_source_desc',    # E  (5)
    'data_user',           # F  (6)
    'usage_purpose',       # G  (7)
    'process_code',        # H  (8)
    'description',         # I  (9)
    'geography',           # J  (10)
    'year',                # K  (11)
    'start_year',          # L  (12)
    'lifetime',            # M  (13)
    'grade',               # N  (14)
    'ef',                  # O  (15)
    'ef_unit',             # P  (16)
    'currency',            # Q  (17)
    'capex',               # R  (18)
    'capex_unit',          # S  (19)
    'fixed_opex',          # T  (20)
    'fixed_opex_unit',     # U  (21)
    'variable_opex',       # V  (22)
    'variable_opex_unit',  # W  (23)
    'tax_cost',            # X  (24)
    'sub_cost',            # Y  (25)
    'efficiency',          # Z  (26)
    'tech_efficiency',     # AA (27)
    'commodity_share',     # AB (28)
    'commodity',           # AC (29)
    'commodity_demand',    # AD (30)
    'interpolation_rule',  # AE (31)
    'afa',                 # AF (32)
    'heat_rate',           # AG (33)
    'capacity',            # AH (34)
    'capacity_type',       # AI (35)
    'constraint',          # AJ (36)
    'uc_rhsrt',            # AK (37)
    'uc_rhsrt_note',       # AL (38)
]

# First data row in the Power sheet (1-based Excel row)
# Rows 1-9 are header/metadata rows preserved from template
POWER_DATA_START_ROW = 10


def write_output(records: list[PowerRecord],
                 template_source,
                 sheet_name: str = 'Power') -> io.BytesIO:
    """
    Write records into the EcoTEA template, preserving all header rows.

    Args:
        records:         List of PowerRecord objects.
        template_source: File path (str/Path) or BytesIO of the blank EcoTEA template.
        sheet_name:      Which sheet to fill (default 'Power').

    Returns:
        BytesIO containing the filled workbook.
    """
    wb = openpyxl.load_workbook(template_source)
    ws = wb[sheet_name]

    # Detect existing formatting from the last header row (row 9 in template)
    # so new data rows look consistent
    ref_row = POWER_DATA_START_ROW - 1  # last header row

    # Clear any existing data rows (safety: avoid duplicates on re-run)
    max_row = ws.max_row
    if max_row >= POWER_DATA_START_ROW:
        for row in ws.iter_rows(min_row=POWER_DATA_START_ROW, max_row=max_row):
            for cell in row:
                cell.value = None

    # Write records
    for row_idx, rec in enumerate(records, start=POWER_DATA_START_ROW):
        for col_idx, field_name in enumerate(POWER_COL_ORDER, start=1):
            val = getattr(rec, field_name, MISSING)

            # Normalise value: MISSING → keep as '-', NaN → '-'
            if val is None or (isinstance(val, float) and str(val) == 'nan'):
                val = MISSING

            cell = ws.cell(row=row_idx, column=col_idx)
            cell.value = val

            # Copy basic alignment from reference row
            ref_cell = ws.cell(row=ref_row, column=col_idx)
            if ref_cell.font:
                cell.font = Font(name=ref_cell.font.name,
                                 size=ref_cell.font.size)
            cell.alignment = Alignment(horizontal='left', vertical='center',
                                       wrap_text=False)

    # Auto-fit column widths (approximate)
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
