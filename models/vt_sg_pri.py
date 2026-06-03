"""
vt_sg_pri.py — Converter for VT_SG_PRI_GREF → WP12345 EcoTEA Endo (Sheet 1).

Supports 39 processes:
  • 33 Import processes (rows 7-39 of Import sheet)
  • 6  Mining processes (rows 7-12 of Mining sheet)

Each process expands to 27 rows (years 2018-2070 step 2).

Uses EndoRecord + TARGET_WRITER = 'endo' to match the 62-column WP12345 template.
"""

import numpy as np

from core.base_model import BaseConverter, EndoRecord, EMPTY


# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

TRACEABILITY = dict(
    data_owner='ESI',
    data_provider='WP1',
    data_source='GREF',
    data_source_desc='SG GREF v8.14; VT_SG_PRI_GREF',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

# Hardcoded emission factors (ktCO2/PJ) as fallback if Coef sheet parse fails
EF_FALLBACK = {
    'COA': 94.6,    'PCK': 100.833, 'NGA': 56.1,   'LNG': 64.2,
    'HFO': 77.4,    'BDS': 70.8,    'DSL': 74.1,   'GSL': 69.3,
    'BMS': 0,       'URA': 0,       'HYG': 0,      'BEN': 43.079,
    'LPG': 63.1,    'RGA': 57.567,  'TGA': 55.733, 'WAS': 50,
    'EEE': 0,       'SOL': 0,       'WAT': 0,
}

# Commodities with no direct emission factor (non-energy goods, OIL, etc.)
NO_EF_COMMS = {
    'OIL',
    '2MM', '3MM', 'CFO',
    'NCH', 'NCW', 'NCC', 'NCM', 'NCI', 'NCV', 'NCT', 'NCA', 'NCF',
    'AGG', 'WAG', 'WSG', 'LUF',
}

# Processes that start in 2071 (set START=2071 in EcoTEA L column)
START_2071_PROCS = {'IMPHYG00', 'IMPHYG01', 'IMPHYG02', 'IMPEEE01'}

# Processes that have ACT_BND constraints
ACT_BND_PROCS = {'IMPEEE00', 'MINRGA00', 'MINWAS00'}

# Of those, only IMPEEE00 also fills the AJ "max import possible" (constraint) column
CONSTRAINT_PROCS = {'IMPEEE00'}

# All output years — 27 rows per process
YEARS = list(range(2018, 2072, 2))

# ─────────────────────────────────────────────────────────────────────────────
# Converter
# ─────────────────────────────────────────────────────────────────────────────

class VTSGPRIConverter(BaseConverter):
    """Converter for VT_SG_PRI_GREF → WP12345 EcoTEA Endo (Sheet 1)."""

    TARGET_WRITER = 'endo'

    def extract_power_records(self) -> list[EndoRecord]:
        self._load_sheets()
        self._parse_coef()
        self._processes: list[dict] = []
        self._parse_import()
        self._parse_mining()

        records = []
        for proc in self._processes:
            records.extend(self._build_rows(proc))
        return records

    # ── Sheet parsers ─────────────────────────────────────────────────────────

    def _parse_coef(self):
        """Parse Coef sheet: commodity abbreviation (col B) → ef ktCO2/PJ (col C)."""
        df = self._sheets['Coef']
        self._ef_by_comm: dict[str, float] = {}
        for i in range(len(df)):
            abbr = df.iloc[i, 1]
            ef   = df.iloc[i, 2]
            if isinstance(abbr, str) and abbr.strip() and not abbr.startswith('*'):
                try:
                    self._ef_by_comm[abbr.strip()] = float(ef)
                except (ValueError, TypeError):
                    pass

    def _find_col_maps(self, df):
        """
        Scan the first 15 rows for a header row containing 'Cost~YEAR' entries.

        Returns (header_row_idx, cost_cols, act_bnd_cols) where cost_cols and
        act_bnd_cols are dicts mapping year (int) → column index (0-based).
        """
        for i in range(min(15, len(df))):
            row = list(df.iloc[i])
            if not any(isinstance(v, str) and v.startswith('Cost~') for v in row):
                continue

            cost_cols: dict[int, int] = {}
            act_bnd_cols: dict[int, int] = {}

            for j, v in enumerate(row):
                if not isinstance(v, str):
                    continue
                if v.startswith('Cost~'):
                    try:
                        yr = int(v.split('~')[1])
                        cost_cols[yr] = j
                    except (IndexError, ValueError):
                        pass
                elif 'ACT_BND' in v:
                    # Header like "ACT_BND~FX~2018" or "ACT_BND~2018~FX"
                    for part in v.split('~'):
                        try:
                            yr = int(part)
                            if 2000 <= yr <= 2100:
                                act_bnd_cols[yr] = j
                                break
                        except ValueError:
                            pass
            return i, cost_cols, act_bnd_cols

        return None, {}, {}

    def _parse_import(self):
        """Parse Import sheet: 33 processes (rows starting with 'IMP' at col C)."""
        df = self._sheets['Import']
        hdr_row, cost_cols, act_bnd_cols = self._find_col_maps(df)
        start_row = (hdr_row + 1) if hdr_row is not None else 6

        SKIP = {'*', 'TechName', 'Technology Name'}

        for i in range(start_row, len(df)):
            code = df.iloc[i, 2]                    # Col C: TechName
            if not isinstance(code, str):
                continue
            code = code.strip()
            if code in SKIP or not code.upper().startswith('IMP'):
                continue

            desc     = df.iloc[i, 3]                # Col D: TechDesc
            comm_out = df.iloc[i, 10]               # Col K: Comm-OUT
            start_v  = df.iloc[i, 11]               # Col L: START
            afa_v    = df.iloc[i, 12]               # Col M: AFA

            # START year
            start_yr = 2071 if code in START_2071_PROCS else 2018
            if not isinstance(start_v, float) or not np.isnan(start_v):
                try:
                    candidate = int(float(start_v))
                    if candidate > 2000:
                        start_yr = candidate
                except (ValueError, TypeError):
                    pass

            self._processes.append({
                'sheet':      'Import',
                'code':       code,
                'description': str(desc).strip() if isinstance(desc, str) else str(desc),
                'comm_out':   comm_out.strip() if isinstance(comm_out, str) else EMPTY,
                'start_year': start_yr,
                'afa':        afa_v if not (isinstance(afa_v, float) and np.isnan(afa_v)) else EMPTY,
                'costs':      self._read_year_vals(df, i, cost_cols),
                'act_bnd':    self._read_year_vals(df, i, act_bnd_cols),
            })

    def _parse_mining(self):
        """Parse Mining sheet: 6 processes (rows starting with 'MIN' at col C)."""
        df = self._sheets['Mining']
        hdr_row, cost_cols, act_bnd_cols = self._find_col_maps(df)
        start_row = (hdr_row + 1) if hdr_row is not None else 6

        SKIP = {'*', 'TechName', 'Technology Name'}

        for i in range(start_row, len(df)):
            code = df.iloc[i, 2]                    # Col C: TechName
            if not isinstance(code, str):
                continue
            code = code.strip()
            if code in SKIP or not code.upper().startswith('MIN'):
                continue

            desc     = df.iloc[i, 3]                # Col D: TechDesc
            comm_out = df.iloc[i, 11]               # Col L: Comm-OUT (Mining uses L, not K)

            self._processes.append({
                'sheet':      'Mining',
                'code':       code,
                'description': str(desc).strip() if isinstance(desc, str) else str(desc),
                'comm_out':   comm_out.strip() if isinstance(comm_out, str) else EMPTY,
                'start_year': 2018,
                'afa':        EMPTY,                # Mining has no AFA column
                'costs':      self._read_year_vals(df, i, cost_cols),
                'act_bnd':    self._read_year_vals(df, i, act_bnd_cols),
            })

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _read_year_vals(self, df, row_idx: int, col_map: dict) -> dict:
        """Extract year-keyed numeric values from a row using col_map."""
        result = {}
        for yr, col in col_map.items():
            if col < len(df.columns):
                v = df.iloc[row_idx, col]
                if not (isinstance(v, float) and np.isnan(v)):
                    result[yr] = v
        return result

    def _pick_value(self, val_dict: dict, year: int):
        """Return value for year; fall back to most recent prior year; else EMPTY."""
        if year in val_dict:
            return val_dict[year]
        candidates = [(yr, v) for yr, v in val_dict.items() if yr <= year]
        if candidates:
            return max(candidates, key=lambda x: x[0])[1]
        return EMPTY

    def _get_ef(self, comm_out: str):
        """Emission factor for a commodity (ktCO2/PJ)."""
        if not isinstance(comm_out, str) or comm_out is EMPTY:
            return EMPTY
        if comm_out in NO_EF_COMMS:
            return EMPTY
        if comm_out in self._ef_by_comm:
            return self._ef_by_comm[comm_out]
        return EF_FALLBACK.get(comm_out, EMPTY)

    # ── Row builder ───────────────────────────────────────────────────────────

    def _build_rows(self, proc: dict) -> list[EndoRecord]:
        """Build 27 EndoRecord rows (one per year 2018-2070) for one process."""
        code     = proc['code']
        comm_out = proc['comm_out']
        ef       = self._get_ef(comm_out)

        has_act_bnd = bool(proc['act_bnd']) and code in ACT_BND_PROCS
        cap_type    = 'FX' if has_act_bnd else EMPTY

        rows = []
        for year in YEARS:
            varom = self._pick_value(proc['costs'], year)

            if has_act_bnd:
                act_val = self._pick_value(proc['act_bnd'], year)
                cap_val = act_val
                # IMPEEE00: constraint (max import) → commodity_demand column
                demand  = act_val if code in CONSTRAINT_PROCS else EMPTY
            else:
                cap_val = EMPTY
                demand  = EMPTY

            rows.append(EndoRecord(
                **TRACEABILITY,
                wp6_title='Primary',
                process_code=code,
                description=proc['description'],
                year=year,
                start_year=proc['start_year'],
                lifetime='NA',          # Import/Mining无寿命参数
                grade=EMPTY,
                ef=ef,
                ef_unit='PJ',
                currency='MSGD2016',
                capex='NA',             # Import/Mining无投资成本
                capex_unit='PJ',
                fixed_opex='NA',        # Import/Mining无固定运维成本
                fixed_opex_unit='PJ*yr(2018)',
                variable_opex=varom,
                variable_opex_unit='PJ (2018)',
                efficiency='NA',        # Import/Mining效率1:1，填NA
                tech_efficiency=EMPTY,
                commodity_share=1,
                commodity=comm_out,
                commodity_demand=demand,
                afa=proc['afa'] if proc['afa'] is not EMPTY else 'NA',
                capacity=cap_val if cap_val is not EMPTY else 'NA',
                capacity_type=cap_type if cap_type is not EMPTY else 'NA',
            ))
        return rows
