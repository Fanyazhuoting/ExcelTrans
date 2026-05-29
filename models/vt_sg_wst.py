"""
vt_sg_wst.py — Converter for VT_SG_WST_GMIT → WP12345 EcoTEA Endo (Sheet 1).

WST (Waste) covers electricity consumption across three waste-management
end-uses in Singapore.  Structure is almost identical to WAT, with the
key difference that Lifetime varies by process (not uniform).

The 3 waste-facility processes:
  WSTELESOR00  Waste Sorting Facilities      EFF = 25 kWh/tone  LIFE = 10
  WSTELERCL00  Recycling Facilities          EFF = 50 kWh/tone  LIFE = 20
  WSTELEOTD00  Other/Landfill Facilities     EFF = 0            LIFE = 35

Shared characteristics across all 3 processes:
  Fuel: ELE → ef = 0
  Grade: '00'
  Tech_EFF: 1
  AFA: 1
  VAROM: all NaN
  STOCK~2018: present for all 3 (no NaN case)
  Comm-IN: WSTELE (first 6 chars of process code, includes WST prefix)

Data_BY column layout (0-based) — standard layout identical to WAT:
  Col 0   Type  ("Demand" / "Process")
  Col 1   Code  (TechName)
  Col 2   Description
  Col 3   Fuel Type  (ELE)
  Col 4   Grade  ('00')
  Col 5   Efficiency  (text "X kWh/tone" or 0)
  Col 6   Technology Efficiency (1)
  Col 7   AFA (1)
  Col 8   Lifetime  (10 / 20 / 35 — varies by process)
  Cols 10-62  Annual demand 2018-2070
  Cols 63-69  INVCOST: 2018,2020,2030,2040,2050,2060,2070
  Cols 70-76  FIXOM:   2018,2020,2030,2040,2050,2060,2070
  Cols 77-83  VAROM:   all NaN
  Col  84     STOCK~2018

3 processes × 27 even years (2018-2070) = 81 EndoRecord rows.

Mapping reference: VT_WST_GMIT_to_EcoTEA_Mapping.xlsx
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
    data_source_desc='SG GREF v8.14; VT_SG_WST_GMIT',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

YEARS = list(range(2018, 2072, 2))

# Data_BY column indices (0-based) — standard layout
COL_TYPE          = 0
COL_CODE          = 1
COL_DESC          = 2
COL_FUEL_TYPE     = 3
COL_GRADE         = 4
COL_EFF           = 5
COL_TECH_EFF      = 6
COL_AFA           = 7
COL_LIFETIME      = 8

COL_DEMAND_START  = 10
COL_INVCOST_START = 63
COL_FIXOM_START   = 70
COL_VAROM_START   = 77    # all NaN — not read
COL_STOCK_2018    = 84    # CG column (1-based col 85)

COST_KEY_YEARS = [2018, 2020, 2030, 2040, 2050, 2060, 2070]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _is_nan(v) -> bool:
    return isinstance(v, float) and np.isnan(v)


def _parse_eff(v) -> object:
    """
    Extract numeric value from Efficiency string or number.

    Examples:
        "25 kWh/tone" → 25.0
        "50 kWh/tone" → 50.0
        0             → 0.0
        0.0           → 0.0

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

class VTSGWSTConverter(BaseConverter):
    """
    Converts VT_SG_WST_GMIT → WP12345 EcoTEA Endo format.

    Same Demand-row tracking pattern as WAT.  All 3 processes: ef=0,
    Comm-IN=WSTELE, VAROM=EMPTY, STOCK present for all.

    ★ Unlike WAT, Lifetime differs by process: SOR=10, RCL=20, OTD=35. ★

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

        year_row = df.iloc[1]
        demand_col_map: dict[int, int] = {}
        for col in range(COL_DEMAND_START, COL_INVCOST_START):
            v = year_row.iloc[col]
            if not _is_nan(v):
                demand_col_map[int(float(v))] = col

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
                v = df.iloc[row, col]
                return None if _is_nan(v) else v

            stock_raw = _val(COL_STOCK_2018)
            stock = float(stock_raw) if stock_raw is not None else None

            invcost = self._read_cost_cols(df, i, COL_INVCOST_START)
            fixom   = self._read_cost_cols(df, i, COL_FIXOM_START)

            grade_raw = _val(COL_GRADE)

            self._processes.append({
                'code':        code.strip(),
                'description': str(df.iloc[i, COL_DESC]).strip(),
                'grade':       str(grade_raw).strip() if grade_raw is not None else '00',
                'efficiency':  _val(COL_EFF),
                'tech_eff':    _val(COL_TECH_EFF),
                'afa':         _val(COL_AFA),
                'lifetime':    _val(COL_LIFETIME),   # ★ 10 / 20 / 35 — read per-process ★
                'invcost':     invcost,
                'fixom':       fixom,
                'stock':       stock,
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

        # Comm-IN = first 6 chars: WSTELE for all 3 processes
        commodity = code[:6] if len(code) >= 6 else EMPTY

        lifetime = safe_int(proc['lifetime'])    # 10 / 20 / 35
        grade    = proc['grade']
        eff      = _parse_eff(proc['efficiency'])
        tech_eff = proc['tech_eff'] if proc['tech_eff'] is not None else EMPTY
        afa_val  = safe_float(proc['afa'])

        ef_val = 0  # all ELE, no combustion emissions

        stock = proc['stock']
        capacity      = stock if stock is not None else EMPTY
        capacity_type = 'FX'  if stock is not None else EMPTY

        rows = []
        for year in YEARS:
            capex      = self._pick_cost(proc['invcost'], year)
            fixed_opex = self._pick_cost(proc['fixom'],   year)

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
        if year in cost_dict:
            return cost_dict[year]
        candidates = [(yr, v) for yr, v in cost_dict.items() if yr <= year]
        if candidates:
            return max(candidates, key=lambda x: x[0])[1]
        return EMPTY
