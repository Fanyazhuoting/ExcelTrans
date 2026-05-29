"""
vt_sg_hfc.py — Converter for VT_SG_HFC_GMIT → WP12345 EcoTEA Endo (Sheet 1).

HFC (Hydrofluorocarbon) tracks refrigerant / fire-suppression gas leakage in
CO2-equivalent units, not conventional energy consumption.

Source file structure (VT_SG_HFC_GMIT.xlsx):
  - Data_BY sheet: 36 rows × 84 cols
      Row 0: column headers
      Row 1: year numbers
          cols 10-62 = annual demand 2018-2070 (kt-CO2eq)
          cols 63-69 = INVCOST key years (all = 0 for HFC)
          cols 70-76 = FIXOM key years   (all = 0 for HFC)
          cols 77-83 = VAROM key years   (all NaN for HFC)
      Row 2+: alternating Demand / Process rows (with NaN gap rows)

Key HFC-specific differences from other modules:
  - ef = 1 ktCO2e/PJ (not 0) — virtual NRG commodity maps 1:1 to CO2e.
  - Fuel Type in Data_BY col 3 is the short suffix (NCH, NCW, …);
    the full Comm-IN code is "HFC" + suffix (e.g. HFCNCH).
  - F-column header is "Percentage" (= 0), not "Efficiency".
  - INVCOST and FIXOM are explicitly 0 (not NA).
  - capex_unit = 'kt'; fixed_opex_unit = 'kt*yr'.
  - No STOCK column (only 84 cols total).

9 HFC sub-categories × 27 even years (2018-2070) = 243 EndoRecord rows.

Mapping reference: VT_HFC_GMIT_to_EcoTEA_Mapping.xlsx
"""

import numpy as np

from core.base_model import BaseConverter, EndoRecord, EMPTY, safe_int, safe_float


# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

TRACEABILITY = dict(
    data_owner='WP1',
    data_provider='WP1',
    data_source='GMIT/GREF Model',
    data_source_desc='SG GREF v8.14; VT_SG_HFC_GMIT',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

# Output years — 27 rows per process (even years, matching EcoTEA convention)
YEARS = list(range(2018, 2072, 2))

# Data_BY column indices (0-based)
COL_TYPE        = 0    # "Demand" / "Process"
COL_CODE        = 1    # TechName
COL_DESC        = 2    # Description
COL_FUEL_TYPE   = 3    # Fuel Type suffix (NCH / NCW / NCC / …)
COL_GRADE       = 4    # Grade (always '00')
COL_PERCENTAGE  = 5    # Percentage (= 0; analogous to Efficiency in other modules)
COL_TECH_EFF    = 6    # Technology Efficiency (= 1)
COL_AFA         = 7    # AFA (= 1)
COL_LIFETIME    = 8    # Lifetime (= 100)

# Year-keyed column blocks (0-based start indices)
COL_DEMAND_START  = 10   # Annual demand/emission cols 2018-2070 (cols 10-62)
COL_INVCOST_START = 63   # INVCOST 7 key years (all = 0)
COL_FIXOM_START   = 70   # FIXOM 7 key years   (all = 0)
COL_VAROM_START   = 77   # VAROM 7 key years   (all NaN)

COST_KEY_YEARS = [2018, 2020, 2030, 2040, 2050, 2060, 2070]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _is_nan(v) -> bool:
    return isinstance(v, float) and np.isnan(v)


# ─────────────────────────────────────────────────────────────────────────────
# Converter
# ─────────────────────────────────────────────────────────────────────────────

class VTSGHFCConverter(BaseConverter):
    """
    Converts VT_SG_HFC_GMIT → WP12345 EcoTEA Endo format.

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

        Row 0: column headers (F-col is 'Percentage', not 'Efficiency').
        Row 1: year header row.
        Row 2+: alternating Demand and Process rows, with NaN gap rows.

        Each Process row inherits HFC emission demand from the most recent
        Demand row above it (same pattern as AGR/BLD).
        """
        df = self._sheets['Data_BY']
        self._processes = []

        # ── Build year→col_index for annual demand columns (cols 10-62)
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

            if not isinstance(row_type, str):
                continue

            row_type_lower = row_type.strip().lower()

            # ── Demand row: capture year-keyed emission values
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

            # Fuel type suffix → full Comm-IN code = "HFC" + suffix
            fuel_suffix = _val(COL_FUEL_TYPE)
            fuel_type = ('HFC' + str(fuel_suffix).strip()) if fuel_suffix else None

            grade_raw = _val(COL_GRADE)
            grade = str(grade_raw).strip() if grade_raw is not None else None

            invcost = self._read_cost_cols(df, i, COL_INVCOST_START)
            fixom   = self._read_cost_cols(df, i, COL_FIXOM_START)
            # VAROM: all NaN for HFC — skip reading

            self._processes.append({
                'code':        code.strip(),
                'description': str(df.iloc[i, COL_DESC]).strip(),
                'fuel_type':   fuel_type,
                'grade':       grade,
                'percentage':  _val(COL_PERCENTAGE),   # = 0 → efficiency field
                'tech_eff':    _val(COL_TECH_EFF),
                'afa':         _val(COL_AFA),
                'lifetime':    _val(COL_LIFETIME),
                'invcost':     invcost,
                'fixom':       fixom,
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
        lifetime = safe_int(proc['lifetime'])
        grade    = proc['grade']  if proc['grade']    is not None else EMPTY
        pct      = safe_float(proc['percentage'], 0.0)
        tech_eff = proc['tech_eff'] if proc['tech_eff'] is not None else EMPTY
        afa_val  = safe_float(proc['afa'])

        # ef = 1 for all HFC processes (virtual NRG commodity → 1 ktCO2e/PJ)
        ef_val = 1

        rows = []
        for year in YEARS:
            # INVCOST / FIXOM: 0 for all years (fallback to last known = 0 at 2018)
            capex      = self._pick_cost(proc['invcost'], year)
            fixed_opex = self._pick_cost(proc['fixom'],   year)

            # VAROM: all NaN → EMPTY
            variable_opex = EMPTY

            # HFC emission demand (kt-CO2eq) from the associated Demand row
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
                capex_unit='kt',          # HFC unit: kt (not GW)
                fixed_opex=fixed_opex,
                fixed_opex_unit='kt*yr',  # HFC unit: kt*yr (not GW*yr)
                variable_opex=variable_opex,
                variable_opex_unit='PJ (2018)',
                efficiency=pct,           # "Percentage" field = 0
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
        Exact match first, then most recent prior year.
        Returns EMPTY if no valid value found.
        """
        if year in cost_dict:
            return cost_dict[year]
        candidates = [(yr, v) for yr, v in cost_dict.items() if yr <= year]
        if candidates:
            return max(candidates, key=lambda x: x[0])[1]
        return EMPTY
