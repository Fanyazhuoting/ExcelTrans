"""
vt_sg_ind.py — Converter for VT_SG_IND_GMIT → WP12345 EcoTEA Endo (Sheet 1).

IND (Industry) covers six sub-sectors, each stored in a separate Data_BY sheet:

  Data_BY_GHG — Non-CO2 GHG tracking (7 processes)    ← special: ef=1, no STOCK
  Data_BY_IRF — Refinery / Petrochemical (12 processes)
  Data_BY_ICH — Industrial Chemicals (9 processes)
  Data_BY_IBM — Biomedical (12 processes)
  Data_BY_IEL — Electronics / Semiconductors (11 processes)
  Data_BY_IOM — Other Manufacturing (21 processes)

All sheets share the same column layout:
  Col 0  Type ("Demand" / "Process" / 0)
  Col 1  Code (TechName)
  Col 2  Description
  Col 3  Fuel Type
  Col 4  Grade (always '00')
  Col 5  Efficiency (text or numeric)
  Col 6  Technology Efficiency
  Col 7  AFA (= 1 for all)
  Col 8  Lifetime
  Col 9  Unit
  Cols 10-62  Annual demand 2018-2070
  Cols 63-69  INVCOST 7 key years
  Cols 70-76  FIXOM 7 key years
  Cols 77-83  VAROM 7 key years (all NaN)
  Col  84     STOCK~2018 (energy sub-sectors only; GHG sheet has only 84 cols)

Key IND-specific rules:
  - Commodity (Comm-IN) = first 6 chars of the process code.
      e.g. IRFELECAE00 → "IRFELE", IOMNGAPHF00 → "IOMNGA", IND2MMPF200 → "IND2MM"
  - GHG sub-sector (Data_BY_GHG):
      ef = 1 (virtual NRG commodity → 1 ktCO2e/kt)
      capex_unit = 'kt'   (not 'GW')
      No STOCK column (sheet is 84 cols wide)
  - Energy sub-sectors (IRF/ICH/IBM/IEL/IOM):
      ef looked up by Fuel Type: ELE=0, HET=0, NGA=56.1, RGA=57.567,
        HFO=77.4, DSL=74.1, LPG=63.1, TGA=55.733, PCK=100.833
      capex_unit = 'GW'
      STOCK~2018 at col 84
  - type=0 process rows appear in IBM, IEL, IOM (TGA/DSL "A=0" processes)
    and must be included (same _is_process_row logic as BLD/HHD).
  - VAROM is NaN for all processes.

72 processes × 27 even years (2018-2070) = 1944 EndoRecord rows.

Mapping reference: VT_IND_GMIT_to_EcoTEA_Mapping.xlsx
"""

import re

import numpy as np

from core.base_model import BaseConverter, EndoRecord, EMPTY, safe_int, safe_float


# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

TRACEABILITY = dict(
    data_owner='WP1',
    data_provider='WP1',
    data_source='GMIT/GREF Model',
    data_source_desc='SG GREF v8.14; VT_SG_IND_GMIT',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

# Output years — 27 rows per process (even years only)
YEARS = list(range(2018, 2072, 2))

# Emission factors by Fuel Type (kt CO2/PJ for energy; kt CO2e/kt for GHG)
EF_BY_FUEL = {
    'ELE': 0,       'HET': 0,
    'NGA': 56.1,    'RGA': 57.567,  'HFO': 77.4,
    'DSL': 74.1,    'LPG': 63.1,    'TGA': 55.733,  'PCK': 100.833,
    # GHG virtual-NRG commodities — ef handled separately (ef=1)
    '2MM': 1,  '3MM': 1,  'CFO': 1,
    'WAG': 1,  'WSG': 1,  'AGG': 1,  'LUF': 1,
}

# GHG fuel types — used to flag a process as GHG-type
GHG_FUELS = {'2MM', '3MM', 'CFO', 'WAG', 'WSG', 'AGG', 'LUF'}

# Sub-sector sheets: (sheet_name, is_ghg)
SUBSECTORS = [
    ('Data_BY_GHG', True),
    ('Data_BY_IRF', False),
    ('Data_BY_ICH', False),
    ('Data_BY_IBM', False),
    ('Data_BY_IEL', False),
    ('Data_BY_IOM', False),
]

# Data_BY column indices (0-based) — identical across all sub-sectors
COL_TYPE        = 0
COL_CODE        = 1
COL_DESC        = 2
COL_FUEL_TYPE   = 3
COL_GRADE       = 4
COL_EFF         = 5
COL_TECH_EFF    = 6
COL_AFA         = 7
COL_LIFETIME    = 8

COL_DEMAND_START  = 10
COL_INVCOST_START = 63
COL_FIXOM_START   = 70
COL_VAROM_START   = 77   # all NaN
COL_STOCK_2018    = 84   # only in energy sheets (91-col sheets)

COST_KEY_YEARS = [2018, 2020, 2030, 2040, 2050, 2060, 2070]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _is_nan(v) -> bool:
    return isinstance(v, float) and np.isnan(v)


def _parse_eff(v) -> object:
    """
    Parse an Efficiency value that may be a number or descriptive text.

    Examples:
        "70% eff"       → 0.70
        "62.3 lm/W"     → 62.3
        "COP = 4.0"     → 4.0
        "COP = 4"       → 4.0
        "14.49 MJ/Nm3"  → 14.49
        0.88            → 0.88
        0               → 0.0

    Returns EMPTY when no numeric value can be extracted.
    """
    if v is None or _is_nan(v):
        return EMPTY
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    m = re.search(r'\d+\.?\d*', s)
    if not m:
        return EMPTY
    num = float(m.group())
    if '%' in s:
        num /= 100.0
    return num


def _is_process_row(row_type) -> bool:
    """
    Return True for Process-type rows.
    Some rows use type=0 (numeric) instead of "Process" (string):
      IBM: IBMTGAPSF00
      IEL: IELTGAPSF00
      IOM: IOMDSLCAF00, IOMDSLPSF00, IOMDSLMOF00
    """
    if isinstance(row_type, str):
        return row_type.strip().lower() == 'process'
    if not _is_nan(row_type) and row_type == 0:
        return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# Converter
# ─────────────────────────────────────────────────────────────────────────────

class VTSGINDConverter(BaseConverter):
    """
    Converts VT_SG_IND_GMIT → WP12345 EcoTEA Endo format.

    Iterates over all 6 sub-sector Data_BY sheets and collects all processes
    into a single flat list before writing.

    TARGET_WRITER = 'endo' signals engine.py to use wp_endo_writer.
    """

    TARGET_WRITER = 'endo'

    def extract_power_records(self) -> list[EndoRecord]:
        self._load_sheets()
        all_processes = []
        for sheet_name, is_ghg in SUBSECTORS:
            procs = self._parse_one_sheet(sheet_name, is_ghg)
            all_processes.extend(procs)

        records = []
        for proc in all_processes:
            records.extend(self._build_rows(proc))
        return records

    # ── Sheet parser ──────────────────────────────────────────────────────────

    def _parse_one_sheet(self, sheet_name: str, is_ghg: bool) -> list[dict]:
        """
        Parse one Data_BY_* sheet and return a list of process dicts.

        The Demand/Process row scanning pattern is the same across all sheets.
        GHG sheet: 84 cols (no STOCK column); energy sheets: 91 cols (STOCK at col 84).
        """
        df = self._sheets[sheet_name]
        n_cols = df.shape[1]
        has_stock = (not is_ghg) and (n_cols > COL_STOCK_2018)
        processes = []

        # ── Build year→col_index for annual demand columns (cols 10-62 only)
        year_row = df.iloc[1]
        demand_col_map: dict[int, int] = {}
        for col in range(COL_DEMAND_START, COL_INVCOST_START):
            v = year_row.iloc[col]
            if not _is_nan(v):
                demand_col_map[int(float(v))] = col

        # ── Scan rows 2+, tracking the most-recent Demand row
        current_demand_by_year: dict[int, object] = {}

        for i in range(2, len(df)):
            row_type = df.iloc[i, COL_TYPE]

            # Demand row: snapshot year-keyed demand values
            if isinstance(row_type, str) and row_type.strip().lower() == 'demand':
                current_demand_by_year = {}
                for yr, col in demand_col_map.items():
                    v = df.iloc[i, col]
                    if not _is_nan(v):
                        current_demand_by_year[yr] = float(v)
                continue

            if not _is_process_row(row_type):
                continue

            code = df.iloc[i, COL_CODE]
            if not isinstance(code, str) or not code.strip():
                continue

            def _val(col, row=i):
                v = df.iloc[row, col]
                return None if _is_nan(v) else v

            invcost = self._read_cost_cols(df, i, COL_INVCOST_START, n_cols)
            fixom   = self._read_cost_cols(df, i, COL_FIXOM_START,   n_cols)

            stock = None
            if has_stock:
                stock_raw = _val(COL_STOCK_2018)
                stock = float(stock_raw) if stock_raw is not None else None

            fuel_raw  = _val(COL_FUEL_TYPE)
            fuel_type = str(fuel_raw).strip().upper() if fuel_raw is not None else ''
            grade_raw = _val(COL_GRADE)
            grade     = str(grade_raw).strip() if grade_raw is not None else '00'

            processes.append({
                'code':        code.strip(),
                'description': str(df.iloc[i, COL_DESC]).strip(),
                'fuel_type':   fuel_type,
                'grade':       grade,
                'efficiency':  _val(COL_EFF),
                'tech_eff':    _val(COL_TECH_EFF),
                'afa':         _val(COL_AFA),
                'lifetime':    _val(COL_LIFETIME),
                'invcost':     invcost,
                'fixom':       fixom,
                'stock':       stock,
                'is_ghg':      is_ghg,
                'demand_by_year': dict(current_demand_by_year),
            })

        return processes

    def _read_cost_cols(self, df, row_idx: int, start_col: int, n_cols: int) -> dict:
        """Read 7 key-year cost values starting at start_col."""
        result = {}
        for j, yr in enumerate(COST_KEY_YEARS):
            col = start_col + j
            if col < n_cols:
                v = df.iloc[row_idx, col]
                if not _is_nan(v):
                    result[yr] = float(v)
        return result

    # ── Row builder ───────────────────────────────────────────────────────────

    def _build_rows(self, proc: dict) -> list[EndoRecord]:
        """Build 27 EndoRecord rows (one per even year 2018-2070)."""
        code      = proc['code']
        fuel_type = proc['fuel_type']
        is_ghg    = proc['is_ghg']

        # Commodity = first 6 chars of process code.
        # Works uniformly across all sub-sectors:
        #   IRFELECAE00 → "IRFELE", IND2MMPF200 → "IND2MM", WATWAGCDE00 → "WATWAG"
        commodity = code[:6] if len(code) >= 6 else EMPTY

        # Static fields
        lifetime  = safe_int(proc['lifetime'])
        grade     = proc['grade']
        eff       = _parse_eff(proc['efficiency'])
        tech_eff  = proc['tech_eff'] if proc['tech_eff'] is not None else EMPTY
        afa_val   = safe_float(proc['afa'])

        # Emission factor:
        #   GHG processes → ef = 1 (virtual NRG commodity, 1 ktCO2e/kt)
        #   Energy processes → ef from Fuel Type lookup table
        ef_val = 1 if is_ghg else EF_BY_FUEL.get(fuel_type, 0)

        # Units differ between GHG and energy processes
        capex_unit = 'kt' if is_ghg else 'GW'

        # Initial capacity: energy processes may have STOCK~2018; GHG never does
        stock = proc['stock']
        capacity      = stock if stock is not None else EMPTY
        capacity_type = 'FX'  if stock is not None else EMPTY

        rows = []
        for year in YEARS:
            capex      = self._pick_cost(proc['invcost'], year)
            fixed_opex = self._pick_cost(proc['fixom'],   year)
            # VAROM is always NaN → EMPTY

            demand = proc['demand_by_year'].get(year, EMPTY)

            rows.append(EndoRecord(
                **TRACEABILITY,
                process_code=code,
                description=proc['description'],
                year=year,
                start_year=2018,
                lifetime=lifetime,
                grade=grade,
                ef=ef_val,
                ef_unit='PJ',
                currency='MSGD2016',
                capex=capex,
                capex_unit=capex_unit,
                fixed_opex=fixed_opex,
                fixed_opex_unit='GW*yr',
                variable_opex=EMPTY,
                variable_opex_unit='PJ (2018)',
                efficiency=eff,
                tech_efficiency=tech_eff,
                commodity_share=100,
                commodity=commodity,
                commodity_demand=demand,
                afa=afa_val,
                capacity=capacity,
                capacity_type=capacity_type,
            ))
        return rows

    def _pick_cost(self, cost_dict: dict, year: int):
        """
        Return the cost value for the given year.
        Uses exact match first, then falls back to the most recent prior year.
        Returns EMPTY if no valid value is found.
        """
        if year in cost_dict:
            return cost_dict[year]
        candidates = [(yr, v) for yr, v in cost_dict.items() if yr <= year]
        if candidates:
            return max(candidates, key=lambda x: x[0])[1]
        return EMPTY
