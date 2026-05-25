"""
vt_sg_agr.py — Converter for VT_SG_AGR_GREF → WP12345 EcoTEA Endo (Sheet 1).

Source file structure (VT_SG_AGR_GMIT.xlsx):
  - Coef sheet: emission factors by fuel type (kt CO2/PJ)
      ELE=0, DSL=74.1, GSL=69.3, NGA=56.1
  - Data_BY sheet: rows × 84 cols
      Row 0: column headers (static cols 0-9, year-keyed cols 10+)
      Row 1: year numbers (demand cols 10-62 = annual 2018-2070;
              cost cols 63-69 = INVCOST key years; 70-76 = FIXOM; 77-83 = VAROM)
      Row 2+: alternating Demand rows and Process rows (with NaN gap rows)

Output: 27 EndoRecord rows (even years 2018, 2020, …, 2070) per process.
Currently VT_SG_AGR_GMIT contains 21 processes across 4 sub-sectors
(AFV, AFO, AFP, AFF).

Column indices (0-based) in Data_BY:
  0  Type (Demand/Process)
  1  Code (TechName)
  2  Description
  3  Fuel Type (ELE/DSL/GSL/NGA)
  4  Grade
  5  Efficiency (may be text: "COP = 3.91", "86% efficiency", etc.)
  6  Technology Efficiency
  7  AFA
  8  Lifetime
  9  Unit
  10-62  Annual demand columns (2018-2070 inclusive)
  63-69  INVCOST key years (2018, 2020, 2030, 2040, 2050, 2060, 2070)
  70-76  FIXOM key years
  77-83  VAROM key years

Mapping reference: VT_AGR_GREF_to_EcoTEA_Mapping.xlsx
"""

import re

import numpy as np

from core.base_model import BaseConverter, EndoRecord, EMPTY


# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

TRACEABILITY = dict(
    data_owner='WP1',
    data_provider='WP1',
    data_source='GMIT/GREF Model',
    data_source_desc='SG GREF v8.14; VT_SG_AGR_GREF',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

# Output years — 27 rows per process (even years only, matching EcoTEA convention)
YEARS = list(range(2018, 2072, 2))

# Emission factors by fuel type abbreviation (kt CO2/PJ, from Coef sheet)
EF_BY_FUEL = {
    'ELE': 0,
    'DSL': 74.1,
    'GSL': 69.3,
    'NGA': 56.1,
}

# Data_BY column indices (0-based)
COL_TYPE        = 0    # "Demand" / "Process"
COL_CODE        = 1    # TechName
COL_DESC        = 2    # Description
COL_FUEL_TYPE   = 3    # Fuel Type (ELE/DSL/GSL/NGA)
COL_GRADE       = 4    # Grade
COL_EFF         = 5    # Efficiency (possibly text)
COL_TECH_EFF    = 6    # Technology Efficiency
COL_AFA         = 7    # AFA (capacity to activity factor)
COL_LIFETIME    = 8    # Lifetime (years)
COL_UNIT        = 9    # Unit (GWh/ktoe etc.)

# Demand / annual year columns (2018-2070 inclusive, cols 10-62)
COL_DEMAND_START  = 10
COL_INVCOST_START = 63  # INVCOST: 7 key years 2018, 2020, 2030, 2040, 2050, 2060, 2070
COL_FIXOM_START   = 70  # FIXOM: 7 key years
COL_VAROM_START   = 77  # VAROM: 7 key years

COST_KEY_YEARS = [2018, 2020, 2030, 2040, 2050, 2060, 2070]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _parse_eff(v) -> object:
    """
    Parse an Efficiency value that may be a number or descriptive text.

    Examples:
        0.86            → 0.86
        "COP = 3.91"    → 3.91
        "86% efficiency"→ 0.86  (% → divide by 100)
        "62.3 lm/W"     → 62.3
        "0.45 litres/kWh" → 0.45
        "EFF:1"         → 1.0

    Returns EMPTY when no numeric value can be extracted.
    """
    if v is None or (isinstance(v, float) and np.isnan(v)):
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


def _is_nan(v) -> bool:
    return isinstance(v, float) and np.isnan(v)


# ─────────────────────────────────────────────────────────────────────────────
# Converter
# ─────────────────────────────────────────────────────────────────────────────

class VTSGAGRConverter(BaseConverter):
    """
    Converts VT_SG_AGR_GREF → WP12345 EcoTEA Endo format.

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

    # ── Sheet parsers ─────────────────────────────────────────────────────────

    def _parse_data_by(self):
        """
        Parse Data_BY sheet.

        Row 0: column headers
        Row 1: year header row (col 10-62 = annual demand years; col 63+ = cost years)
        Row 2+: alternating Demand and Process rows (NaN gap rows between groups)

        Demand rows carry annual demand values in cols 10-62.
        Each Process row inherits the demand from the most recent Demand row above it.
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
        current_demand_by_year: dict[int, object] = {}  # year → demand value

        for i in range(2, len(df)):
            row_type = df.iloc[i, COL_TYPE]

            # Skip gap / non-string rows
            if not isinstance(row_type, str):
                continue

            row_type_lower = row_type.strip().lower()

            # ── Demand row: update the current demand context
            if row_type_lower == 'demand':
                current_demand_by_year = {}
                for yr, col in demand_col_map.items():
                    v = df.iloc[i, col]
                    if not _is_nan(v):
                        current_demand_by_year[yr] = float(v)
                continue

            # ── Process row
            if row_type_lower != 'process':
                continue

            code = df.iloc[i, COL_CODE]
            if not isinstance(code, str) or not code.strip():
                continue

            def _val(col):
                v = df.iloc[i, col]
                return None if _is_nan(v) else v

            invcost = self._read_cost_cols(df, i, COL_INVCOST_START)
            fixom   = self._read_cost_cols(df, i, COL_FIXOM_START)
            varom   = self._read_cost_cols(df, i, COL_VAROM_START)

            fuel_type = _val(COL_FUEL_TYPE)
            grade_raw = _val(COL_GRADE)

            self._processes.append({
                'code':        code.strip(),
                'description': str(df.iloc[i, COL_DESC]).strip(),
                'fuel_type':   str(fuel_type).strip() if fuel_type is not None else None,
                'grade':       str(grade_raw).strip() if grade_raw is not None else None,
                'efficiency':  _val(COL_EFF),
                'tech_eff':    _val(COL_TECH_EFF),
                'afa':         _val(COL_AFA),
                'lifetime':    _val(COL_LIFETIME),
                'invcost':     invcost,
                'fixom':       fixom,
                'varom':       varom,
                'demand_by_year': dict(current_demand_by_year),  # snapshot
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
        lifetime  = int(proc['lifetime']) if proc['lifetime'] is not None else EMPTY
        grade     = proc['grade'] if proc['grade'] is not None else EMPTY
        eff       = _parse_eff(proc['efficiency'])
        tech_eff  = proc['tech_eff'] if proc['tech_eff'] is not None else EMPTY
        afa_val   = proc['afa']      if proc['afa']       is not None else EMPTY

        # Emission factor from fuel type (Coef table)
        ef_val = EF_BY_FUEL.get(fuel_type.upper(), 0)

        rows = []
        for year in YEARS:
            capex      = self._pick_cost(proc['invcost'], year)
            fixed_opex = self._pick_cost(proc['fixom'],   year)
            varom_raw  = self._pick_cost(proc['varom'],   year)
            variable_opex = varom_raw if varom_raw is not EMPTY else EMPTY

            # Commodity demand: from the associated Demand row (even years only)
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
                variable_opex=variable_opex,
                variable_opex_unit='PJ (2018)',
                efficiency=eff,
                tech_efficiency=tech_eff,
                commodity_share=100,
                commodity=fuel_type if fuel_type else EMPTY,
                commodity_demand=demand,
                afa=afa_val,
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
