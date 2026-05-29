"""
vt_sg_trn.py — Converter for VT_SG_TRN_GMIT → WP12345 EcoTEA Endo (Sheet 1).

TRN (Transportation) covers 7 sub-sectors in a single Data_BY sheet:

  TRP — Road Transport (personal cars, buses, motorcycles, LGV/HGV)  ← largest
  TPS — Sea Transport Bunkers (tankers, tug boats, others)
  TAF — Airport Facilities (HVAC, motors, lighting, forklifts, LGV/HGV)
  TPF — General Port Facilities (★ all demand = 0 ★)
  TOF — Other Transport Facilities (street lighting + misc)
  TPA — PSA Port Facilities (cranes, forklifts, LGV/HGV, LNG heavy vehicles)
  TPJ — Jurong Port Facilities (similar structure to TPA)

★ TRN-specific differences from all other modules ★

1. STOCK column is at Col 96 (CR, 0-based index 95) — NOT Col 85 (CG).
   Cols 85-95 are occupied by TAXCOST and SUBCOST (TRN-only fields).

2. INVCOST/FIXOM/VAROM key year order differs:
   Other modules: [2018, 2020, 2030, 2040, 2050, 2060, 2070]
   TRN:           [2018, 2020, 2030, 2040, 2050, 2060, 2025]  ← BR = 2025, not 2070

3. TAXCOST (cols 85-89, years 2018/2020/2024/2030/2050):
   Road tax / licence fees. Non-zero only for TRP road vehicle processes.

4. SUBCOST (cols 90-95, years 2018/2020/2021/2024/2030/2050):
   Government EV purchase subsidies (negative values).
   Only TRPELECAR91, TRPELEPHC91, TRPELETXI91 have SUBCOST.

Data_BY column layout (0-based):
  Col 0   Type  ("Demand" / "Process" / 0)
  Col 1   Code  (TechName)
  Col 2   Description
  Col 3   Fuel Type  (GSL/DSL/ELE/TGA/LNG)
  Col 4   Grade  ('00' / int 90 / int 91)
  Col 5   Efficiency  (text or numeric)
  Col 6   Technology Efficiency
  Col 7   AFA
  Col 8   Lifetime
  Col 9   Unit
  Cols 10-62  Annual demand 2018-2070 (53 years)
  Cols 63-69  INVCOST: 2018,2020,2030,2040,2050,2060,2025
  Cols 70-76  FIXOM:   2018,2020,2030,2040,2050,2060,2025
  Cols 77-83  VAROM:   all NaN
  Cols 84-88  TAXCOST: 2018,2020,2024,2030,2050
  Cols 89-94  SUBCOST: 2018,2020,2021,2024,2030,2050
  Col  95     STOCK~2018 (only col 95 used; NaN → EMPTY)

Emission factors by Fuel Type (kt CO2 / PJ):
  ELE=0, GSL=69.3, DSL=74.1, TGA=55.733, LNG=64.2, HYD=0, BDS=70.8

~71 processes × 27 even years (2018-2070) = ~1917 EndoRecord rows.
(TPF sub-sector: 7 processes with demand=0, still output 27 rows each.)

Mapping reference: VT_TRN_GMIT_to_EcoTEA_Mapping.xlsx
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
    data_source_desc='SG GREF v8.14; VT_SG_TRN_GMIT',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

# Output years — 27 rows per process (even years only)
YEARS = list(range(2018, 2072, 2))

# Emission factors by Fuel Type (kt CO2 / PJ)
EF_BY_FUEL = {
    'ELE': 0,       'GSL': 69.3,    'DSL': 74.1,
    'TGA': 55.733,  'LNG': 64.2,    'HYD': 0,
    'BDS': 70.8,
}

# Data_BY column indices (0-based)
COL_TYPE          = 0     # "Demand" / "Process" / 0
COL_CODE          = 1     # TechName
COL_DESC          = 2     # Description
COL_FUEL_TYPE     = 3     # Fuel Type
COL_GRADE         = 4     # Grade ('00' / 90 / 91)
COL_EFF           = 5     # Efficiency (text or numeric)
COL_TECH_EFF      = 6     # Technology Efficiency
COL_AFA           = 7     # AFA (capacity-to-activity factor)
COL_LIFETIME      = 8     # Lifetime (years)

# Year-keyed column blocks (0-based start indices)
COL_DEMAND_START  = 10    # Annual demand 2018-2070 (cols 10-62)
COL_INVCOST_START = 63    # INVCOST: 7 cols, years 2018,2020,2030,2040,2050,2060,2025
COL_FIXOM_START   = 70    # FIXOM:   7 cols, same year order as INVCOST
COL_VAROM_START   = 77    # VAROM:   7 cols, all NaN — not read
COL_TAXCOST_START = 84    # TAXCOST: 5 cols, years 2018,2020,2024,2030,2050 (TRP only)
COL_SUBCOST_START = 89    # SUBCOST: 6 cols, years 2018,2020,2021,2024,2030,2050 (EV only)
COL_STOCK_2018    = 95    # ★ STOCK~2018 (0-based 95 = 1-based col 96, CR column) ★

# Key-year sequences matching the actual Data_BY column layout
INVCOST_KEY_YEARS  = [2018, 2020, 2030, 2040, 2050, 2060, 2025]  # ★ 2025 not 2070 ★
FIXOM_KEY_YEARS    = [2018, 2020, 2030, 2040, 2050, 2060, 2025]
TAXCOST_KEY_YEARS  = [2018, 2020, 2024, 2030, 2050]
SUBCOST_KEY_YEARS  = [2018, 2020, 2021, 2024, 2030, 2050]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _is_nan(v) -> bool:
    return isinstance(v, float) and np.isnan(v)


def _parse_eff(v) -> object:
    """
    Parse an Efficiency value that may be a number or descriptive text.

    Examples:
        "13.33 km/litre"  → 13.33
        "5.67 km/kWh"     → 5.67
        "COP = 3.91"      → 3.91
        "37% eff"         → 0.37
        "86% eff"         → 0.86
        "9W/m2"           → 9.0
        "90 lm/W"         → 90.0
        "0.53 L/kWh"      → 0.53
        0                 → 0.0
        0.86              → 0.86

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

    Most rows use the string "Process". Five TRN processes use numeric 0:
      TRPELEPRB91, TAFGSLFLT00, TPAGSLFLT00, TPAELELGT00, TPJGSLFLT00.
    """
    if isinstance(row_type, str):
        return row_type.strip().lower() == 'process'
    if not _is_nan(row_type) and row_type == 0:
        return True
    return False


def _normalize_grade(grade_raw) -> object:
    """
    Normalize grade to a consistent string representation.

    TRN grades:
      '00' (string) — standard fuel processes
      91 (int/float stored in Excel) — TRP electric vehicles (EV)
      90 (int/float stored in Excel) — TAF/TPA/TPJ electric forklifts / HGVs

    Returns EMPTY if blank/None.
    """
    if grade_raw is None or _is_nan(grade_raw):
        return EMPTY
    if isinstance(grade_raw, (int, float)):
        return str(int(float(grade_raw)))
    return str(grade_raw).strip() or EMPTY


# ─────────────────────────────────────────────────────────────────────────────
# Converter
# ─────────────────────────────────────────────────────────────────────────────

class VTSGTRNConverter(BaseConverter):
    """
    Converts VT_SG_TRN_GMIT → WP12345 EcoTEA Endo format.

    All 7 TRN sub-sectors live in a single Data_BY sheet.
    The converter scans that sheet sequentially using the same
    Demand-row tracking pattern as other endo models.

    TARGET_WRITER = 'endo' signals engine.py to use wp_endo_writer.

    TRN-specific:
      - STOCK is at col 95 (0-based), not col 84 as in other modules.
      - INVCOST/FIXOM last key year is 2025, not 2070.
      - TAXCOST (tax_cost field) and SUBCOST (sub_cost field) are TRN-only.
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
        Parse the single Data_BY sheet (all 7 sub-sectors).

        Row 0: column headers
        Row 1: year numbers (demand at cols 10-62; cost blocks at cols 63+)
        Row 2+: alternating Demand and Process rows (with NaN gap rows).

        Each Process row inherits demand from the most-recent Demand row above.
        Process rows with type=0 (5 TRN processes) are handled by _is_process_row.
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
        n_cols = df.shape[1]

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

            # Cost blocks
            invcost  = self._read_cost_cols(df, i, COL_INVCOST_START,  INVCOST_KEY_YEARS,  n_cols)
            fixom    = self._read_cost_cols(df, i, COL_FIXOM_START,    FIXOM_KEY_YEARS,    n_cols)
            taxcost  = self._read_cost_cols(df, i, COL_TAXCOST_START,  TAXCOST_KEY_YEARS,  n_cols)
            subcost  = self._read_cost_cols(df, i, COL_SUBCOST_START,  SUBCOST_KEY_YEARS,  n_cols)

            # STOCK~2018: at col 95 (0-based); NaN → None
            stock_raw = _val(COL_STOCK_2018) if n_cols > COL_STOCK_2018 else None
            stock = float(stock_raw) if stock_raw is not None else None

            fuel_raw  = _val(COL_FUEL_TYPE)
            fuel_type = str(fuel_raw).strip().upper() if fuel_raw is not None else ''
            grade_raw = _val(COL_GRADE)

            self._processes.append({
                'code':        code.strip(),
                'description': str(df.iloc[i, COL_DESC]).strip(),
                'fuel_type':   fuel_type,
                'grade':       _normalize_grade(grade_raw),
                'efficiency':  _val(COL_EFF),
                'tech_eff':    _val(COL_TECH_EFF),
                'afa':         _val(COL_AFA),
                'lifetime':    _val(COL_LIFETIME),
                'invcost':     invcost,
                'fixom':       fixom,
                'taxcost':     taxcost,
                'subcost':     subcost,
                'stock':       stock,
                'demand_by_year': dict(current_demand_by_year),
            })

    def _read_cost_cols(self, df, row_idx: int, start_col: int,
                        key_years: list, n_cols: int) -> dict:
        """Read cost values at start_col + j for each year in key_years."""
        result = {}
        for j, yr in enumerate(key_years):
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

        # Commodity (Comm-IN) = first 6 chars of process code.
        # Includes sub-sector prefix (TRPGSL, TAFELE, TPSDSL, TPAELE, etc.)
        # — NOT the generic fuel type string.
        commodity = code[:6] if len(code) >= 6 else EMPTY

        # Static fields
        lifetime  = safe_int(proc['lifetime'])
        grade     = proc['grade']
        eff       = _parse_eff(proc['efficiency'])
        tech_eff  = proc['tech_eff'] if proc['tech_eff'] is not None else EMPTY
        afa_val   = safe_float(proc['afa'])

        # Emission factor by Fuel Type
        ef_val = EF_BY_FUEL.get(fuel_type, 0)

        # Initial capacity: STOCK~2018 (col 95); None → EMPTY
        stock = proc['stock']
        capacity      = stock if stock is not None else EMPTY
        capacity_type = 'FX'  if stock is not None else EMPTY

        rows = []
        for year in YEARS:
            capex      = self._pick_cost(proc['invcost'],  year)
            fixed_opex = self._pick_cost(proc['fixom'],    year)
            tax_cost   = self._pick_cost(proc['taxcost'],  year)
            sub_cost   = self._pick_cost(proc['subcost'],  year)

            # Commodity demand from most-recent Demand row above this process
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
                tax_cost=tax_cost,
                sub_cost=sub_cost,
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
