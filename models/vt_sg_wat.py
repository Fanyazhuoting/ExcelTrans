"""
vt_sg_wat.py — Converter for VT_SG_WAT_GMIT → WP12345 EcoTEA Endo (Sheet 1).

WAT (Water) is the simplest energy module in SG-TIMES: a single Data_BY
sheet with 5 End-uses, 5 Process rows, and 5 Demand rows.  All processes
share the same key parameters:

  Fuel: ELE only → ef = 0
  Grade: '00' for all
  Lifetime: 25 years (uniform across all processes)
  VAROM: all NaN
  STOCK~2018: present for all 5 processes (no NaN case)
  Comm-IN: WATELE (first 6 chars of process code, includes WAT prefix)

The 5 water utility processes:
  WATELEDSA00  Desalination standard           EFF = 4.1  kWh/m3
  WATELERCM00  Water reclamation standard      EFF = 0.74 kWh/m3
  WATELENWT00  NEWater standard                EFF = 1.06 kWh/m3
  WATELEOTD00  Other water facilities standard EFF = 0.49 kWh/m3
  WATELEWWK00  Waterworks standard             EFF = 0.49 kWh/m3

Data_BY column layout (0-based), identical to standard modules (HHD, IFC):
  Col 0   Type  ("Demand" / "Process")
  Col 1   Code  (TechName)
  Col 2   Description
  Col 3   Fuel Type  (ELE for all)
  Col 4   Grade  ('00' for all)
  Col 5   Efficiency  (text "X kWh/m3")
  Col 6   Technology Efficiency (1 for all)
  Col 7   AFA (1 for all)
  Col 8   Lifetime (25 for all)
  Cols 10-62  Annual demand 2018-2070
  Cols 63-69  INVCOST: 2018,2020,2030,2040,2050,2060,2070
  Cols 70-76  FIXOM:   2018,2020,2030,2040,2050,2060,2070
  Cols 77-83  VAROM:   all NaN
  Col  84     STOCK~2018 (all 5 processes have a value)

5 processes × 27 even years (2018-2070) = 135 EndoRecord rows.

Mapping reference: VT_WAT_GMIT_to_EcoTEA_Mapping.xlsx
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
    data_source_desc='SG GREF v8.14; VT_SG_WAT_GMIT',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

# Output years — 27 rows per process (even years only)
YEARS = list(range(2018, 2072, 2))

# Data_BY column indices (0-based)
COL_TYPE          = 0
COL_CODE          = 1
COL_DESC          = 2
COL_FUEL_TYPE     = 3
COL_GRADE         = 4
COL_EFF           = 5
COL_TECH_EFF      = 6
COL_AFA           = 7
COL_LIFETIME      = 8

# Year-keyed column blocks (0-based start indices) — standard layout
COL_DEMAND_START  = 10    # Annual demand 2018-2070 (cols 10-62)
COL_INVCOST_START = 63    # INVCOST: 2018,2020,2030,2040,2050,2060,2070
COL_FIXOM_START   = 70    # FIXOM:   2018,2020,2030,2040,2050,2060,2070
COL_VAROM_START   = 77    # VAROM:   all NaN — not read
COL_STOCK_2018    = 84    # STOCK~2018 (CG column, 1-based col 85)

COST_KEY_YEARS = [2018, 2020, 2030, 2040, 2050, 2060, 2070]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _is_nan(v) -> bool:
    return isinstance(v, float) and np.isnan(v)


def _parse_eff(v) -> object:
    """
    Extract numeric value from Efficiency string.

    Examples:
        "4.1 kWh/m3"  → 4.1
        "0.74 kWh/m3" → 0.74
        "1.06 kWh/m3" → 1.06
        "0.49 kWh/m3" → 0.49
        4.1           → 4.1

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


# ─────────────────────────────────────────────────────────────────────────────
# Converter
# ─────────────────────────────────────────────────────────────────────────────

class VTSGWATConverter(BaseConverter):
    """
    Converts VT_SG_WAT_GMIT → WP12345 EcoTEA Endo format.

    Uses the standard single-sheet Demand-row tracking pattern.
    All 5 WAT processes: ef=0, LIFE=25, Comm-IN=WATELE, VAROM=EMPTY.

    TARGET_WRITER = 'endo' signals engine.py to use wp_endo_writer.
    """

    TARGET_WRITER = 'endo'

    def extract_power_records(self) -> list[EndoRecord]:
        self._load_sheets()
        self._parse_data_by()
        records = []
        for proc in self._processes:
            records.extend(self._build_rows(proc))
        return records

    # ── Sheet parser ──────────────────────────────────────────────────────────

    def _parse_data_by(self):
        """
        Parse Data_BY sheet.

        Row 0: column headers
        Row 1: year numbers (demand at cols 10-62; cost blocks at cols 63+)
        Row 2+: alternating Demand and Process rows (with NaN gap rows).

        All WAT Process rows use the string "Process" as type (no type=0 rows).
        """
        df = self._sheets['Data_BY']
        self._processes = []

        # ── Build year→col_index for annual demand columns (cols 10-62 only)
        year_row = df.iloc[1]
        demand_col_map: dict[int, int] = {}
        for col in range(COL_DEMAND_START, COL_INVCOST_START):
            v = year_row.iloc[col]
            if not _is_nan(v):
                demand_col_map[int(float(v))] = col

        # ── Scan rows 2+, tracking most-recent Demand row
        current_demand_by_year: dict[int, object] = {}

        for i in range(2, len(df)):
            row_type = df.iloc[i, COL_TYPE]

            if not isinstance(row_type, str):
                continue

            row_type_lower = row_type.strip().lower()

            # Demand row: snapshot year-keyed demand values
            if row_type_lower == 'demand':
                current_demand_by_year = {}
                for yr, col in demand_col_map.items():
                    v = df.iloc[i, col]
                    if not _is_nan(v):
                        current_demand_by_year[yr] = float(v)
                continue

            if row_type_lower != 'process':
                continue

            code = df.iloc[i, COL_CODE]
            if not isinstance(code, str) or not code.strip():
                continue

            def _val(col, row=i):
                v = df.iloc[row, col]
                return None if _is_nan(v) else v

            # STOCK~2018: all WAT processes have a value
            stock_raw = _val(COL_STOCK_2018)
            stock = float(stock_raw) if stock_raw is not None else None

            invcost = self._read_cost_cols(df, i, COL_INVCOST_START)
            fixom   = self._read_cost_cols(df, i, COL_FIXOM_START)
            # VAROM: all NaN — skip reading

            grade_raw = _val(COL_GRADE)

            self._processes.append({
                'code':        code.strip(),
                'description': str(df.iloc[i, COL_DESC]).strip(),
                'grade':       str(grade_raw).strip() if grade_raw is not None else '00',
                'efficiency':  _val(COL_EFF),
                'tech_eff':    _val(COL_TECH_EFF),
                'afa':         _val(COL_AFA),
                'lifetime':    _val(COL_LIFETIME),
                'invcost':     invcost,
                'fixom':       fixom,
                'stock':       stock,
                'demand_by_year': dict(current_demand_by_year),
            })

    def _read_cost_cols(self, df, row_idx: int, start_col: int) -> dict:
        """Read 7 key-year cost values starting at start_col."""
        result = {}
        for j, yr in enumerate(COST_KEY_YEARS):
            col = start_col + j
            if col < len(df.columns):
                v = df.iloc[row_idx, col]
                if not _is_nan(v):
                    result[yr] = float(v)
        return result

    # ── Row builder ───────────────────────────────────────────────────────────

    def _build_rows(self, proc: dict) -> list[EndoRecord]:
        """Build 27 EndoRecord rows (one per even year 2018-2070)."""
        code = proc['code']

        # Commodity (Comm-IN) = first 6 chars of process code = 'WATELE'
        # All WAT processes: WATELEDSA00 → 'WATELE', etc.
        commodity = code[:6] if len(code) >= 6 else EMPTY

        # Static fields — uniform across all WAT processes
        lifetime = safe_int(proc['lifetime'])
        grade    = proc['grade']
        eff      = _parse_eff(proc['efficiency'])
        tech_eff = proc['tech_eff'] if proc['tech_eff'] is not None else EMPTY
        afa_val  = safe_float(proc['afa'])

        # All WAT processes are pure ELE — no combustion emissions
        ef_val = 0

        # All WAT processes have STOCK~2018
        stock = proc['stock']
        capacity      = stock if stock is not None else EMPTY
        capacity_type = 'FX'  if stock is not None else EMPTY

        rows = []
        for year in YEARS:
            capex      = self._pick_cost(proc['invcost'], year)
            fixed_opex = self._pick_cost(proc['fixom'],   year)
            # VAROM is always NaN for WAT → EMPTY

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
                capex_unit='GW',
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
