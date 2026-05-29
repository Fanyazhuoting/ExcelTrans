"""
vt_sg_ifc.py — Converter for VT_SG_IFC_GMIT → WP12345 EcoTEA Endo (Sheet 1).

IFC (ICT Facilities / Data Centres) covers three sub-sectors of electricity-
consuming ICT infrastructure:

  IFE — Enterprise Data Centres   (5 processes)
  IFS — Service Provider DCs      (5 processes)
  IFC — Telecom Networks          (1 process)

All 11 processes are pure ELE consumers (ef = 0, EFF = 0, VAROM = NaN).

Key IFC-specific difference from other modules:
  - Commodity (Comm-IN) is NOT the Fuel Type column value ("ELE").
    It is derived from the first 6 characters of the process code:
      IFEELE* → commodity = "IFEELE"
      IFSELE* → commodity = "IFSELE"
      IFCELE* → commodity = "IFCELE"
  - IFE and IFS processes have STOCK~2018 (initial capacity, col 84).
  - IFCELETEL00 has no STOCK (NaN → EMPTY).
  - IFEELESVR00 is the only process with Tech_EFF ≠ 1 (= 1.83).

Source file structure (VT_SG_IFC_GMIT.xlsx):
  - Data_BY sheet: 44 rows × 91 cols
      Row 0: column headers
      Row 1: year numbers
          cols 10-62 = annual demand 2018-2070
          cols 63-69 = INVCOST key years (2018,2020,2030,2040,2050,2060,2070)
          cols 70-76 = FIXOM key years
          cols 77-83 = VAROM key years (all NaN)
          cols 84-90 = STOCK key years (only col 84 = STOCK~2018 used)
      Row 2+: alternating Demand and Process rows (with NaN gap rows)

11 processes × 27 even years (2018-2070) = 297 EndoRecord rows.

Mapping reference: VT_IFC_GMIT_to_EcoTEA_Mapping.xlsx
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
    data_source_desc='SG GREF v8.14; VT_SG_IFC_GMIT',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

# Output years — 27 rows per process (even years only, matching EcoTEA convention)
YEARS = list(range(2018, 2072, 2))

# Data_BY column indices (0-based)
COL_TYPE        = 0    # "Demand" / "Process"
COL_CODE        = 1    # TechName
COL_DESC        = 2    # Description
COL_FUEL_TYPE   = 3    # Fuel Type (always ELE — NOT used for Comm-IN)
COL_GRADE       = 4    # Grade (always '00')
COL_EFF         = 5    # Efficiency (always 0)
COL_TECH_EFF    = 6    # Technology Efficiency (1 for all except IFEELESVR00=1.83)
COL_AFA         = 7    # AFA (= 1 for all IFC processes)
COL_LIFETIME    = 8    # Lifetime (years; varies by End-use: 5/10/15/20)

# Year-keyed column blocks (0-based start indices)
COL_DEMAND_START  = 10   # Annual demand 2018-2070 (cols 10-62)
COL_INVCOST_START = 63   # INVCOST 7 key years
COL_FIXOM_START   = 70   # FIXOM 7 key years
COL_VAROM_START   = 77   # VAROM 7 key years (all NaN for IFC)
COL_STOCK_2018    = 84   # STOCK~2018 initial capacity (IFE/IFS only; IFC TEL = NaN)

COST_KEY_YEARS = [2018, 2020, 2030, 2040, 2050, 2060, 2070]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _is_nan(v) -> bool:
    return isinstance(v, float) and np.isnan(v)


# ─────────────────────────────────────────────────────────────────────────────
# Converter
# ─────────────────────────────────────────────────────────────────────────────

class VTSGIFCConverter(BaseConverter):
    """
    Converts VT_SG_IFC_GMIT → WP12345 EcoTEA Endo format.

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
        Row 1: year header row (cols 10-62 annual; cols 63+ cost blocks)
        Row 2+: alternating Demand and Process rows, with NaN gap rows.

        Each Process row inherits demand from the most-recent Demand row above it.
        All IFC processes use string "Process" as type (no type=0 quirk here).
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

            # STOCK~2018: IFE/IFS processes have a value; IFCELETEL00 has NaN
            stock_raw = _val(COL_STOCK_2018)
            stock = float(stock_raw) if stock_raw is not None else None

            invcost = self._read_cost_cols(df, i, COL_INVCOST_START)
            fixom   = self._read_cost_cols(df, i, COL_FIXOM_START)
            # VAROM: all NaN for IFC — skip reading

            grade_raw = _val(COL_GRADE)

            self._processes.append({
                'code':        code.strip(),
                'description': str(df.iloc[i, COL_DESC]).strip(),
                'grade':       str(grade_raw).strip() if grade_raw is not None else '00',
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

        # Comm-IN = first 6 chars of process code (IFEELE / IFSELE / IFCELE).
        # This is IFC-specific: the sub-sector prefix is part of the commodity.
        commodity = code[:6] if len(code) >= 6 else EMPTY

        # Static fields
        lifetime  = safe_int(proc['lifetime'])
        grade     = proc['grade']
        tech_eff  = proc['tech_eff'] if proc['tech_eff'] is not None else EMPTY
        afa_val   = safe_float(proc['afa'])

        # All IFC processes: ef = 0 (pure ELE, no combustion emissions),
        # efficiency = 0 (direct electrical use, no conversion concept)
        ef_val  = 0
        eff_val = 0

        # STOCK~2018: IFE/IFS processes have a value; IFCELETEL00 does not
        stock = proc['stock']
        capacity      = stock if stock is not None else EMPTY
        capacity_type = 'FX'  if stock is not None else EMPTY

        rows = []
        for year in YEARS:
            capex      = self._pick_cost(proc['invcost'], year)
            fixed_opex = self._pick_cost(proc['fixom'],   year)
            # VAROM is always NaN for IFC → EMPTY

            # Commodity demand from the most-recent Demand row above this process
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
                efficiency=eff_val,
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
