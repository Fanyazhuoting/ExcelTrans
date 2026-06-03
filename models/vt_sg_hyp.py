"""
vt_sg_hyp.py — Converter for VT_SG_HYP_GMIT → WP12345 EcoTEA Endo (Sheet 1).

HYP (Hydrogen Production) covers three hydrogen production pathways in Singapore:
  HYPHYB  Blue Hydrogen  (Natural Gas + CCS)      ef = 56.1 ktCO2/PJ (placeholder)
  HYPNGA  Grey Hydrogen  (Natural Gas SMR, no CCS) ef = 56.1 ktCO2/PJ
  HYPELE  Green Hydrogen (Renewable Electricity)   ef = 0

★ IMPORTANT: VT_SG_HYP_GMIT Data_BY is currently empty (no process rows).
  This converter gracefully returns 0 records when Data_BY has no data.
  Once process data is populated in Data_BY, it will extract records normally.

★ HYP-specific parameters (differ from all other endo modules):
  CAP2ACT = 31.536  (1 GW × 8760 h/yr × 3600 s/h ÷ 10^15 × 10^3 ≈ 31.536 PJ/GWa)
             → stored in EndoRecord.afa field (col 43, EcoTEA AR column)
  ef by pathway:
    HYPNGA → 56.1 ktCO2/PJ  (NGA combustion, from EMI sheet = Coef C11)
    HYPELE → 0               (electrolysis, zero direct emission)
    HYPHYB → 56.1 ktCO2/PJ  (conservative placeholder; actual < 56.1 after CCS)

Data_BY column layout (0-based) — HYP-specific, more compact than other modules:
  Col 0   Type         ("Demand" / "Process")
  Col 1   Code         (TechName)
  Col 2   Description
  Col 3   Fuel Type
  Col 4   Grade
  Col 6   EFF          (energy conversion efficiency)
  Col 7   Tech_EFF / AFA (typically 1)
  Col 8   Lifetime     (years)
  Cols 9-15   INVCOST: 2018,2020,2030,2040,2050,2060,2070  (J–P)
  Cols 17-23  FIXOM:   2018,2020,2030,2040,2050,2060,2070  (R–X)
  Cols 25-31  VAROM:   2018,2020,2030,2040,2050,2060,2070  (Z–AF)
  (No STOCK column — capacity = EMPTY for all HYP processes)

Mapping reference: VT_HYP_GMIT_to_EcoTEA_Mapping.xlsx
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
    data_source_desc='SG GREF v8.14; VT_SG_HYP_GMIT',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

YEARS = list(range(2018, 2072, 2))  # 27 even years

# Data_BY column indices (0-based)
COL_TYPE          = 0   # A
COL_CODE          = 1   # B
COL_DESC          = 2   # C
COL_FUEL_TYPE     = 3   # D
COL_GRADE         = 4   # E
# Col 5 (F): Percentage/Efficiency label — skip
COL_EFF           = 6   # G
COL_TECH_EFF      = 7   # H — AFA / tech_eff (typically 1)
COL_LIFETIME      = 8   # I

COL_INVCOST_START = 9   # J — 7 key years: 2018,2020,2030,2040,2050,2060,2070
COL_FIXOM_START   = 17  # R — 7 key years
COL_VAROM_START   = 25  # Z — 7 key years

COST_KEY_YEARS = [2018, 2020, 2030, 2040, 2050, 2060, 2070]

# HYP special: CAP2ACT = 31.536 (GW → PJ conversion)
CAP_TO_ACT = 31.536

# Emission factors by hydrogen production pathway (ktCO2/PJ)
# Source: EMI sheet → Coef C11 (NGA) = 56.1
EF_BY_PATHWAY = {
    'HYPNGA': 56.1,   # grey hydrogen (natural gas SMR, no CCS)
    'HYPELE': 0,      # green hydrogen (electrolysis, zero direct emission)
    'HYPHYB': 56.1,   # blue hydrogen (NGA + CCS; placeholder — actual < 56.1)
}


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _is_nan(v) -> bool:
    return isinstance(v, float) and np.isnan(v)


def _ef_for_code(code: str) -> float:
    """Look up emission factor by process code prefix (first 6 chars)."""
    prefix = code[:6].upper() if len(code) >= 6 else code.upper()
    return EF_BY_PATHWAY.get(prefix, 0)


# ─────────────────────────────────────────────────────────────────────────────
# Converter
# ─────────────────────────────────────────────────────────────────────────────

class VTSGHYPConverter(BaseConverter):
    """
    Converts VT_SG_HYP_GMIT → WP12345 EcoTEA Endo format.

    ★ Data_BY is currently empty in the source file. This converter returns
      0 records gracefully when no process rows are found. Once Data_BY is
      populated by an analyst, the full extraction will run automatically.

    Key HYP differences from other endo modules:
      - afa (CAP2ACT) = 31.536 fixed (vs. 1 in all other modules)
      - ef varies by pathway: HYPNGA=56.1, HYPELE=0, HYPHYB=56.1(placeholder)
      - No STOCK column → capacity = EMPTY
      - VAROM columns may contain actual values
      - INVCOST starts at col J (index 9), not col BL as in other modules

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
        df = self._sheets['Data_BY']
        self._processes = []

        # Row 1 (index 1) should contain year numbers in demand columns
        if len(df) < 2:
            return  # Empty sheet — no data

        year_row = df.iloc[1]
        demand_col_map: dict[int, int] = {}
        for col in range(COL_INVCOST_START):
            v = year_row.iloc[col] if col < len(df.columns) else None
            if v is not None and not _is_nan(v):
                try:
                    demand_col_map[int(float(v))] = col
                except (ValueError, TypeError):
                    pass

        # If no year columns found in row 1, try scanning from INVCOST start
        # (some modules have demand cols between col 10-62)
        if not demand_col_map:
            for col in range(10, min(COL_INVCOST_START, len(df.columns))):
                v = year_row.iloc[col]
                if v is not None and not _is_nan(v):
                    try:
                        demand_col_map[int(float(v))] = col
                    except (ValueError, TypeError):
                        pass

        current_demand_by_year: dict[int, object] = {}

        for i in range(2, len(df)):
            row_type = df.iloc[i, COL_TYPE]

            if not isinstance(row_type, str):
                continue

            row_type_lower = row_type.strip().lower()

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
                if col >= len(df.columns):
                    return None
                v = df.iloc[row, col]
                return None if _is_nan(v) else v

            invcost = self._read_cost_cols(df, i, COL_INVCOST_START)
            fixom   = self._read_cost_cols(df, i, COL_FIXOM_START)
            varom   = self._read_cost_cols(df, i, COL_VAROM_START)

            grade_raw = _val(COL_GRADE)

            self._processes.append({
                'code':        code.strip(),
                'description': str(df.iloc[i, COL_DESC]).strip(),
                'grade':       str(grade_raw).strip() if grade_raw is not None else '00',
                'efficiency':  _val(COL_EFF),
                'tech_eff':    _val(COL_TECH_EFF),
                'lifetime':    _val(COL_LIFETIME),
                'invcost':     invcost,
                'fixom':       fixom,
                'varom':       varom,
                'demand_by_year': dict(current_demand_by_year),
            })

    def _read_cost_cols(self, df, row_idx: int, start_col: int) -> dict:
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
        code = proc['code']

        # Comm-IN = first 6 chars (HYPHYB / HYPNGA / HYPELE)
        commodity = code[:6] if len(code) >= 6 else EMPTY

        lifetime = safe_int(proc['lifetime'])
        grade    = proc['grade']

        eff_raw  = proc['efficiency']
        eff      = float(eff_raw) if eff_raw is not None else EMPTY

        tech_eff_raw = proc['tech_eff']
        tech_eff = float(tech_eff_raw) if tech_eff_raw is not None else EMPTY

        ef_val = _ef_for_code(code)

        # HYP special: CAP2ACT = 31.536 (not read from Data_BY)
        afa_val = CAP_TO_ACT

        # No STOCK column in HYP
        capacity      = EMPTY
        capacity_type = EMPTY

        rows = []
        for year in YEARS:
            capex      = self._pick_cost(proc['invcost'], year)
            fixed_opex = self._pick_cost(proc['fixom'],   year)
            var_opex   = self._pick_cost(proc['varom'],   year)

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
                variable_opex=var_opex,
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
        if year in cost_dict:
            return cost_dict[year]
        candidates = [(yr, v) for yr, v in cost_dict.items() if yr <= year]
        if candidates:
            return max(candidates, key=lambda x: x[0])[1]
        return EMPTY
