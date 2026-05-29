"""
vt_sg_hhd.py — Converter for VT_SG_HHD_GMIT → WP12345 EcoTEA Endo (Sheet 1).

HHD (Household Devices) covers residential appliances and lighting, classified
by NEA energy-label star grade.

Source file structure (VT_SG_HHD_GMIT.xlsx):
  - Coef sheet: emission factors by fuel type (kt CO2/PJ)
      ELE=0, LPG=63.1, TGA=55.733
  - Data_BY sheet: 57 rows × 91 cols
      Row 0: column headers (static cols 0-9, year-keyed cols 10+)
      Row 1: year numbers
          cols 10-62 = annual demand 2018-2070
          cols 63-69 = INVCOST key years (2018,2020,2030,2040,2050,2060,2070)
          cols 70-76 = FIXOM key years
          cols 77-83 = VAROM key years (all NaN for HHD)
          cols 84-90 = STOCK key years (only col 84 = STOCK~2018 used)
      Row 2+: alternating Demand and Process rows (with NaN gap rows)
              Two process rows have type=0 (numeric) instead of "Process":
                HHDELECDW02 and HHDELEHWH01.

10 End-uses: RFG, ACC, LGT, TVS, CDW, HWH, CKG, OTD, NCD, STB
29 Process rows × 27 even years (2018-2070) = 783 EndoRecord rows.

Column indices (0-based) in Data_BY:
  0  Type ("Demand" / "Process" / 0)
  1  Code (TechName)
  2  Description
  3  Fuel Type (ELE/LPG/TGA)
  4  Grade  (NEA star level: '00','01','02','03','04','90', or int 10/20/90)
  5  Efficiency (may be text: "COP = 5.95", "0.57 kWh/L", "100 lm/W", etc.)
  6  Technology Efficiency
  7  AFA (= 1 for all HHD processes)
  8  Lifetime (years, 5-16)
  9  Unit (GWh / PJ)
  10-62   Annual demand columns (2018-2070)
  63-69   INVCOST key years
  70-76   FIXOM key years
  77-83   VAROM key years (all NaN)
  84      STOCK~2018 initial capacity (processes without stock have NaN here)

Mapping reference: VT_HHD_GMIT_to_EcoTEA_Mapping.xlsx
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
    data_source_desc='SG GREF v8.14; VT_SG_HHD_GMIT',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

# Output years — 27 rows per process (even years only, matching EcoTEA convention)
YEARS = list(range(2018, 2072, 2))

# Emission factors by fuel type abbreviation (kt CO2/PJ, from Coef sheet)
EF_BY_FUEL = {
    'ELE': 0,
    'LPG': 63.1,
    'TGA': 55.733,
}

# Data_BY column indices (0-based)
COL_TYPE        = 0    # "Demand" / "Process" / 0
COL_CODE        = 1    # TechName
COL_DESC        = 2    # Description
COL_FUEL_TYPE   = 3    # Fuel Type (ELE/LPG/TGA)
COL_GRADE       = 4    # Grade (NEA star level)
COL_EFF         = 5    # Efficiency (possibly text)
COL_TECH_EFF    = 6    # Technology Efficiency
COL_AFA         = 7    # AFA (= 1 for all HHD processes)
COL_LIFETIME    = 8    # Lifetime (years)
COL_UNIT        = 9    # Unit (GWh / PJ)

# Year-keyed column blocks (0-based start indices)
COL_DEMAND_START  = 10   # Annual demand 2018-2070 (cols 10-62)
COL_INVCOST_START = 63   # INVCOST 7 key years
COL_FIXOM_START   = 70   # FIXOM 7 key years
COL_VAROM_START   = 77   # VAROM 7 key years (all NaN for HHD)
COL_STOCK_2018    = 84   # STOCK~2018 initial capacity

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
        "COP = 5.95"        → 5.95
        "0.57 kWh/L"        → 0.57
        "100 lm/W"          → 100.0
        "82.5 lm/W"         → 82.5
        "1.50 kW"           → 1.5
        0.37                → 0.37
        "86% efficiency"    → 0.86  (% → divide by 100)
        0                   → 0.0

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
    Some rows use the string "Process"; others use numeric 0 (Excel quirk in HHD).
    Affected processes: HHDELECDW02, HHDELEHWH01.
    """
    if isinstance(row_type, str):
        return row_type.strip().lower() == 'process'
    if not _is_nan(row_type) and row_type == 0:
        return True
    return False


def _normalize_grade(grade_raw) -> object:
    """
    Normalize grade to a consistent string representation.

    HHD grades are either:
      - Strings already ('00', '01', '02', '03', '04', '90')
      - Integers (10, 20, 90) stored as float/int in Excel

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

class VTSGHHDConverter(BaseConverter):
    """
    Converts VT_SG_HHD_GMIT → WP12345 EcoTEA Endo format.

    TARGET_WRITER = 'endo' signals engine.py to use wp_endo_writer instead of
    the default ecotea_writer.
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
        Row 1: year header row (cols 10-62 = annual demand; cols 63+ = cost blocks)
        Row 2+: Demand rows, Process rows, and NaN gap rows.

        Process rows in HHD have type = "Process" (string) or 0 (numeric):
          - HHDELECDW02 and HHDELEHWH01 use type=0.
        Each Process row inherits demand from the most recent Demand row above it.
        """
        df = self._sheets['Data_BY']
        self._processes = []

        # ── Build year→col_index for annual demand columns (cols 10-62 only)
        year_row = df.iloc[1]
        demand_col_map: dict[int, int] = {}
        for col in range(COL_DEMAND_START, COL_INVCOST_START):
            v = year_row.iloc[col]
            if not _is_nan(v):
                yr = int(float(v))
                demand_col_map[yr] = col

        # ── Scan rows 2+, tracking most-recent Demand row
        current_demand_by_year: dict[int, object] = {}

        for i in range(2, len(df)):
            row_type = df.iloc[i, COL_TYPE]

            # ── Demand row: snapshot year-keyed demand values
            if isinstance(row_type, str) and row_type.strip().lower() == 'demand':
                current_demand_by_year = {}
                for yr, col in demand_col_map.items():
                    v = df.iloc[i, col]
                    if not _is_nan(v):
                        current_demand_by_year[yr] = float(v)
                continue

            # ── Skip non-process rows (NaN gaps)
            if not _is_process_row(row_type):
                continue

            code = df.iloc[i, COL_CODE]
            if not isinstance(code, str) or not code.strip():
                continue

            def _val(col, row=i):
                v = df.iloc[row, col]
                return None if _is_nan(v) else v

            # STOCK~2018 (initial capacity): only present for some processes
            stock_raw = _val(COL_STOCK_2018)
            stock = float(stock_raw) if stock_raw is not None else None

            invcost = self._read_cost_cols(df, i, COL_INVCOST_START)
            fixom   = self._read_cost_cols(df, i, COL_FIXOM_START)
            # VAROM: all NaN for HHD — skip reading

            fuel_type = _val(COL_FUEL_TYPE)
            grade_raw = _val(COL_GRADE)

            self._processes.append({
                'code':        code.strip(),
                'description': str(df.iloc[i, COL_DESC]).strip(),
                'fuel_type':   str(fuel_type).strip() if fuel_type is not None else None,
                'grade':       _normalize_grade(grade_raw),
                'efficiency':  _val(COL_EFF),
                'tech_eff':    _val(COL_TECH_EFF),
                'afa':         _val(COL_AFA),
                'lifetime':    _val(COL_LIFETIME),
                'invcost':     invcost,
                'fixom':       fixom,
                'stock':       stock,          # None → no STOCK for this process
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
        code      = proc['code']
        fuel_type = proc['fuel_type'] or ''

        # Static fields
        lifetime  = safe_int(proc['lifetime'])
        grade     = proc['grade']   # already normalised by _normalize_grade
        eff       = _parse_eff(proc['efficiency'])
        tech_eff  = proc['tech_eff'] if proc['tech_eff'] is not None else EMPTY
        afa_val   = safe_float(proc['afa'])

        # Emission factor from fuel type
        ef_val = EF_BY_FUEL.get(fuel_type.upper(), 0)

        # Initial capacity (STOCK~2018): only for processes with a value
        stock = proc['stock']
        capacity      = stock if stock is not None else EMPTY
        capacity_type = 'FX'  if stock is not None else EMPTY

        rows = []
        for year in YEARS:
            capex      = self._pick_cost(proc['invcost'], year)
            fixed_opex = self._pick_cost(proc['fixom'],   year)
            # VAROM is always NaN for HHD → leave as EMPTY

            # Commodity demand from the associated Demand row
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
                commodity=fuel_type if fuel_type else EMPTY,
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
