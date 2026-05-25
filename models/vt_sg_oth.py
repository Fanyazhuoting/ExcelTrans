"""
vt_sg_oth.py — Converter for VT_SG_OTH_GREF → WP12345 EcoTEA Endo (Sheet 1).

Source file structure (VT_SG_OTH_GREF.xlsx):
  - Data_BY sheet: 4 rows × 91 cols
      Row 0: column headers
      Row 1: year numbers (demand cols 10-62 = annual 2018-2070;
              cost cols 63-69 = INVCOST key years; 70-76 = FIXOM; 77-83 = VAROM;
              84-90 = STOCK)
      Row 2: OTHOTD  — demand data
      Row 3: OTHELEOTD00 — process data
  - OTD sheet: used to read NCAP_BND (initial capacity) and capacity_type

Output: 27 EndoRecord rows (even years 2018, 2020, …, 2070) per process.
Currently VT_SG_OTH_GREF contains one process: OTHELEOTD00.

Mapping reference: VT_OTH_GMIT_to_EcoTEA_Mapping.xlsx
"""

import pandas as pd
import numpy as np

from core.base_model import BaseConverter, EndoRecord, EMPTY


# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

TRACEABILITY = dict(
    data_owner='WP1',
    data_provider='WP1',
    data_source='GMIT/GREF Model',
    data_source_desc='SG GREF v8.14; VT_SG_OTH_GMIT',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

# Output years — 27 rows per process (even years only, matching EcoTEA convention)
YEARS = list(range(2018, 2072, 2))

# Data_BY column indices (0-based)
# Static / per-process columns
COL_TYPE        = 0    # "Demand" / "Process"
COL_CODE        = 1    # TechName
COL_DESC        = 2    # Description
COL_FUEL_TYPE   = 3    # Fuel Type (→ Commodity)
COL_GRADE       = 4    # Grade
COL_EFF         = 5    # Efficiency
COL_TECH_EFF    = 6    # Technology Efficiency
COL_AFA         = 7    # AFA
COL_LIFETIME    = 8    # Lifetime

# Demand / STOCK year columns (annual, 2018-2070 inclusive)
COL_DEMAND_START = 10  # col 10 = year 2018, col 11 = 2019, …

# Cost columns (7 key years: 2018, 2020, 2030, 2040, 2050, 2060, 2070)
COST_KEY_YEARS = [2018, 2020, 2030, 2040, 2050, 2060, 2070]

COL_INVCOST_START = 63  # BL = col 63 → 2018
COL_FIXOM_START   = 70  # BS = col 70 → 2018
COL_VAROM_START   = 77  # BZ = col 77 → 2018
COL_STOCK_START   = 84  # CG = col 84 → 2018


# ─────────────────────────────────────────────────────────────────────────────
# Converter
# ─────────────────────────────────────────────────────────────────────────────

class VTSGOTHConverter(BaseConverter):
    """
    Converts VT_SG_OTH_GREF → WP12345 EcoTEA Endo format.

    TARGET_WRITER = 'endo' signals engine.py to use wp_endo_writer instead of
    the default ecotea_writer.
    """

    TARGET_WRITER = 'endo'

    def extract_power_records(self) -> list[EndoRecord]:
        self._load_sheets()
        self._parse_data_by()
        self._parse_otd()

        records = []
        for proc in self._processes:
            records.extend(self._build_rows(proc))
        return records

    # ── Sheet parsers ─────────────────────────────────────────────────────────

    def _parse_data_by(self):
        """
        Parse Data_BY sheet.

        Row 1: year header row (col 10+ = demand years; col 63+ = cost years)
        Row 2: OTHOTD demand row
        Row 3+: Process rows
        """
        df = self._sheets['Data_BY']
        self._processes = []
        self._demand_by_year: dict[int, object] = {}

        # Build year→col_index for demand columns (annual, cols 10-62 only).
        # Cols 63+ are INVCOST/FIXOM/VAROM/STOCK blocks which also carry year
        # headers; we must NOT include those to avoid overwriting the annual
        # demand column mapping with cost-block columns.
        year_row = df.iloc[1]
        demand_col_map: dict[int, int] = {}
        for col in range(COL_DEMAND_START, COL_INVCOST_START):
            v = year_row.iloc[col]
            if not (isinstance(v, float) and np.isnan(v)):
                yr = int(float(v))
                demand_col_map[yr] = col

        # Parse OTHOTD (demand) row — row index 2
        demand_row = df.iloc[2]
        for yr, col in demand_col_map.items():
            v = demand_row.iloc[col]
            if not (isinstance(v, float) and np.isnan(v)):
                self._demand_by_year[yr] = float(v)

        # Parse process rows (row 3 onwards)
        for i in range(3, len(df)):
            row_type = df.iloc[i, COL_TYPE]
            if not isinstance(row_type, str):
                continue
            if row_type.strip().lower() != 'process':
                continue

            code = df.iloc[i, COL_CODE]
            if not isinstance(code, str) or not code.strip():
                continue

            def _val(col):
                v = df.iloc[i, col]
                return None if (isinstance(v, float) and np.isnan(v)) else v

            # Build year-keyed cost dicts
            invcost = self._read_cost_cols(df, i, COL_INVCOST_START)
            fixom   = self._read_cost_cols(df, i, COL_FIXOM_START)
            varom   = self._read_cost_cols(df, i, COL_VAROM_START)
            stock   = self._read_cost_cols(df, i, COL_STOCK_START)

            self._processes.append({
                'code':        code.strip(),
                'description': str(df.iloc[i, COL_DESC]).strip(),
                'fuel_type':   _val(COL_FUEL_TYPE),
                'grade':       _val(COL_GRADE),
                'efficiency':  _val(COL_EFF),
                'tech_eff':    _val(COL_TECH_EFF),
                'afa':         _val(COL_AFA),
                'lifetime':    _val(COL_LIFETIME),
                'invcost':     invcost,
                'fixom':       fixom,
                'varom':       varom,
                'stock':       stock,
            })

    def _read_cost_cols(self, df, row_idx: int, start_col: int) -> dict:
        """Read 7 key-year cost values starting at start_col."""
        result = {}
        for j, yr in enumerate(COST_KEY_YEARS):
            col = start_col + j
            if col < len(df.columns):
                v = df.iloc[row_idx, col]
                if not (isinstance(v, float) and np.isnan(v)):
                    result[yr] = float(v)
        return result

    def _parse_otd(self):
        """
        Parse OTD sheet to get NCAP_BND (initial capacity) and capacity_type.

        OTD row 9 contains column headers; data row(s) follow from row 12.
        We locate 'NCAP_BND~0' column header to find capacity; capacity_type
        is 'FX' when that column is present and has a value, otherwise EMPTY.
        """
        df = self._sheets['OTD']
        self._otd_capacity: dict[str, object] = {}    # code → capacity GW
        self._otd_cap_type: dict[str, object] = {}    # code → 'FX'/'UP'/EMPTY

        # Find header row (contains 'TechName')
        hdr_row = None
        for i in range(min(15, len(df))):
            row = df.iloc[i].tolist()
            if 'TechName' in row and 'NCAP_BND~0' in row:
                hdr_row = i
                break
        if hdr_row is None:
            return

        headers = df.iloc[hdr_row].tolist()
        try:
            tech_col    = headers.index('TechName')
            ncap_col    = headers.index('NCAP_BND~0')
        except ValueError:
            return

        # Read process data rows below the header
        for i in range(hdr_row + 1, len(df)):
            code = df.iloc[i, tech_col]
            if not isinstance(code, str) or not code.strip():
                continue
            if code.strip().startswith('*') or code.strip() == 'TechName':
                continue

            code = code.strip()
            ncap_val = df.iloc[i, ncap_col]
            if isinstance(ncap_val, float) and np.isnan(ncap_val):
                self._otd_capacity[code] = EMPTY
                self._otd_cap_type[code] = EMPTY
            else:
                self._otd_capacity[code] = float(ncap_val)
                self._otd_cap_type[code] = 'FX'

    # ── Row builder ───────────────────────────────────────────────────────────

    def _build_rows(self, proc: dict) -> list[EndoRecord]:
        """Build 27 EndoRecord rows (one per even year 2018-2070)."""
        code = proc['code']

        # Static fields
        lifetime  = int(proc['lifetime']) if proc['lifetime'] is not None else EMPTY
        grade     = str(proc['grade']).strip() if proc['grade'] is not None else EMPTY
        eff       = proc['efficiency'] if proc['efficiency'] is not None else EMPTY
        tech_eff  = proc['tech_eff']   if proc['tech_eff']  is not None else EMPTY
        afa_val   = proc['afa']        if proc['afa']       is not None else EMPTY
        commodity = str(proc['fuel_type']).strip() if proc['fuel_type'] is not None else EMPTY

        capacity      = self._otd_capacity.get(code, EMPTY)
        capacity_type = self._otd_cap_type.get(code, EMPTY)

        rows = []
        for year in YEARS:
            capex     = self._pick_cost(proc['invcost'], year)
            fixed_opex = self._pick_cost(proc['fixom'],  year)
            varom_raw  = self._pick_cost(proc['varom'],  year)
            # VAROM: leave as EMPTY (blank) if no value found
            variable_opex = varom_raw if varom_raw is not EMPTY else EMPTY

            # Commodity demand from OTHOTD (even years only; fall back nearest)
            demand = self._demand_by_year.get(year, EMPTY)

            rows.append(EndoRecord(
                **TRACEABILITY,
                process_code=code,
                description=proc['description'],
                year=year,
                start_year=2018,
                lifetime=lifetime,
                grade=grade,
                ef=0,
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
