"""
vt_sg_pwr.py — Converter for VT_SG_PWR_GMIT → WP12345 EcoTEA Endo (Sheet 1).

PWR (Power) is a supply-side module. Key differences from demand-side models:
  - No Demand/Process row pairing; no commodity_demand (AK) values
  - Data_BY column layout is compact (58 cols, not 91):
      E=EFF, F=AFA, G=LIFE, H=Bound, I~O=INVCOST(7yrs), P~V=FIXOM, W~AC=VAROM,
      AD~BF=RESID(initial capacity MW, by year)
  - CAP2ACT = 31.536 for all ELE/SOL/CHP processes (afa field, not AFA factor)
  - EFF = thermal efficiency (e.g. 0.497 for CCGT-F)
  - RESID (MW) = initial installed capacity → EcoTEA capacity column

Mapping reference: VT_PWR_GMIT_to_EcoTEA_Mapping.xlsx
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
    data_source_desc='SG GREF v8.14; VT_SG_PWR_GMIT',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

# Emission factor lookup keyed by VT commodity name
EF_MAP = {
    'PWRNGA':  56.1,
    'PWRCOA':  94.6,
    'PWRBMS':  0,
    'PWRWAS':  50,
    'PWRPCK':  100.833,
    'PWRHFO':  77.4,
    'PWRDSL':  74.1,
    'PWRURA':  0,
    'PWRSOL':  0,
    'PWRHYD':  0,
}

# CAP2ACT values by process type (EcoTEA afa field = capacity to activity factor)
CAP2ACT_ELE_SOL = 31.536   # all ELE/SOL/CHP processes: 1 GW × 8760 h × 3600 s ÷ 10^15 × 10^9
CAP2ACT_PRE     = 1.0      # PRE (retrofit) processes

# Processes whose row-year expansion is driven by INVCOST year columns
# (solar-type: no capacity RESID data; expand per cost year)
SOLAR_PROCESSES = {
    'PWRSOLLPV00', 'BDESOLLPV00', 'BDNSOLLPV00', 'BDFSOLLPV00',
    'WATSOLFPV00', 'TPASOLLPV00', 'TPJSOLLPV00',
}

# Processes that have multiple Comm-IN inputs → commodity shown as '-'
MULTI_INPUT_PROCS = {'PWRWASWTE00'}

# Processes where EcoTEA variable-opex-unit is suppressed (shown as '-')
NO_VAROM_UNIT_PROCS = {'BDESOLLPV00', 'BDNSOLLPV00', 'BDFSOLLPV00',
                        'WATSOLFPV00', 'TPASOLLPV00', 'TPJSOLLPV00'}

# Cogeneration processes: EcoTEA uses electrical efficiency (0.3), not total
COGEN_ELEC_EFF = {
    'IRFNGACGP00': 0.3,
    'IBMNGACGP00': 0.3,
}

# Processes where capacity and capacity_type are forced to EMPTY
# (retrofit technology — capacity comes from a separate mechanism)
NO_CAPACITY_PROCS = {'PWRNGACCH11'}

# For PWRWASWTE00 bound is NaN in Data_BY but EcoTEA shows UP
BOUND_OVERRIDES = {'PWRWASWTE00': 'UP'}

# Fixed commodity overrides (where simple Comm-IN lookup doesn't apply)
COMMODITY_OVERRIDES = {
    'PWRBMCSTP00': 'PWRBMS+PWACOA',
    'PWRWASWTE00': EMPTY,
    'PWRSOLLPV00': EMPTY,
    'BDESOLLPV00': EMPTY,
    'BDNSOLLPV00': EMPTY,
    'BDFSOLLPV00': EMPTY,
    'WATSOLFPV00': EMPTY,
    'TPASOLLPV00': EMPTY,
    'TPJSOLLPV00': EMPTY,
    'IRFNGACGP00': EMPTY,
    'IBMNGACGP00': EMPTY,
}

# Canonical process output order (matches EcoTEA reference sheet)
CANONICAL_PROCESS_ORDER = [
    'PWRNGACCF01', 'PWRNGACCF00', 'PWRBMCSTP00', 'PWRWASWTE00',
    'PWRNGACCF26', 'PWRNGACCH11',
    'PWRSOLLPV00', 'IRFNGACGP00', 'IBMNGACGP00',
    'BDESOLLPV00', 'BDNSOLLPV00', 'BDFSOLLPV00',
    'WATSOLFPV00', 'TPASOLLPV00', 'TPJSOLLPV00',
]

# Solar variants (excludes PWRSOLLPV00) — these use '-' for efficiency
SOLAR_VARIANTS = {
    'BDESOLLPV00', 'BDNSOLLPV00', 'BDFSOLLPV00',
    'WATSOLFPV00', 'TPASOLLPV00', 'TPJSOLLPV00',
}

# Commodity-share overrides
COMMODITY_SHARE_OVERRIDES = {
    'PWRBMCSTP00': '20%+80%',
    # These processes have commodity='-' but still carry share=1
    'PWRWASWTE00': 1,
    'IRFNGACGP00': 1,
    'IBMNGACGP00': 1,
}

# Interpolation rule overrides (most processes use EMPTY)
INTERPOLATION_OVERRIDES = {
    'PWRBMCSTP00': 5,
}


# ─────────────────────────────────────────────────────────────────────────────
# Converter class
# ─────────────────────────────────────────────────────────────────────────────

class VTSGPWRConverter(BaseConverter):

    TARGET_WRITER = 'endo'

    def extract_power_records(self) -> list[EndoRecord]:
        self._load_sheets()
        self._parse_data_by()
        self._parse_pwr()
        self._parse_sol()
        self._parse_emi()

        # Apply canonical order; fall back to parsed order for any unknowns
        known = set(CANONICAL_PROCESS_ORDER)
        ordered = [c for c in CANONICAL_PROCESS_ORDER if c in set(self._process_order)]
        extras  = [c for c in self._process_order if c not in known]
        process_order = ordered + extras

        records = []
        for code in process_order:
            records.extend(self._build_rows(code))
        return records

    # ── Parsing helpers ──────────────────────────────────────────────────────

    def _parse_data_by(self):
        """Parse the Data_BY sheet into per-process dicts."""
        df = self._sheets['Data_BY']

        # Row 1 contains column headers; capacity year headers start at col 29
        self._cap_year_headers = [
            int(float(v)) for v in df.iloc[1, 29:].values
            if not (isinstance(v, float) and np.isnan(v))
        ]

        self._data_by: dict[str, dict] = {}
        self._process_order: list[str] = []

        for i in range(2, len(df)):
            code = df.iloc[i, 1]
            if not isinstance(code, str) or code in ('*', 'Main Grid',
                    'Embeded or Autoproducer', 'Code'):
                continue
            if pd.isna(code):
                continue
            # Skip new-additions rows (col 0 == 0)
            if df.iloc[i, 0] == 0:
                continue

            self._process_order.append(code)
            n_cap = len(self._cap_year_headers)
            cap_vals = df.iloc[i, 29: 29 + n_cap].values

            self._data_by[code] = {
                'description': df.iloc[i, 2],
                'heat_rate':   df.iloc[i, 3],
                'efficiency':  df.iloc[i, 4],
                'afa':         df.iloc[i, 5],
                'lifetime':    df.iloc[i, 6],
                'bound':       df.iloc[i, 7],
                # INVCOST columns: indices 8-14 (years 2018,2020,2030,2040,2050,2060,2070)
                'invcost': {
                    int(float(df.iloc[1, 8 + j])): df.iloc[i, 8 + j]
                    for j in range(7)
                    if not (isinstance(df.iloc[1, 8 + j], float) and
                            np.isnan(df.iloc[1, 8 + j]))
                },
                # FIXOM columns: indices 15-21
                'fixom': {
                    int(float(df.iloc[1, 15 + j])): df.iloc[i, 15 + j]
                    for j in range(7)
                    if not (isinstance(df.iloc[1, 15 + j], float) and
                            np.isnan(df.iloc[1, 15 + j]))
                },
                # VAROM columns: indices 22-28
                'varom': {
                    int(float(df.iloc[1, 22 + j])): df.iloc[i, 22 + j]
                    for j in range(7)
                    if not (isinstance(df.iloc[1, 22 + j], float) and
                            np.isnan(df.iloc[1, 22 + j]))
                },
                # Capacity by year
                'capacity': {
                    self._cap_year_headers[j]: cap_vals[j]
                    for j in range(n_cap)
                },
            }

    def _parse_pwr(self):
        """Parse PWR sheet: Comm-IN, Share info, START year."""
        df = self._sheets['PWR']
        self._pwr_comm: dict[str, str] = {}       # code → Comm-IN
        self._pwr_share0: dict[str, object] = {}  # code → interpolation rule
        self._pwr_share: dict[str, object] = {}   # code → share value
        self._pwr_start: dict[str, int] = {}      # code → START year

        current_code = None
        for i in range(6, len(df)):
            code = df.iloc[i, 2]
            if isinstance(code, str) and code not in ('*', 'Technology Name'):
                current_code = code
            # ~FI_T section: TechName at col 9, Comm-IN at col 10, START at col 15
            fi_t_code = df.iloc[i, 9]
            if isinstance(fi_t_code, str) and not fi_t_code.startswith('*'):
                start_val = df.iloc[i, 15]
                if not (isinstance(start_val, float) and np.isnan(start_val)):
                    try:
                        yr = int(float(start_val))
                        if yr > 2000:
                            self._pwr_start[fi_t_code] = yr
                    except (ValueError, TypeError):
                        pass
            comm = df.iloc[i, 10]
            if current_code and isinstance(comm, str) and comm != '*' and not pd.isna(comm):
                # Only store the FIRST Comm-IN for each process
                if current_code not in self._pwr_comm:
                    self._pwr_comm[current_code] = comm
                    self._pwr_share0[current_code] = df.iloc[i, 12]
                    self._pwr_share[current_code] = df.iloc[i, 13]

        # Parse CGP sheet for CHP process START years and Comm-IN
        if 'CGP' in self._sheets:
            df_cgp = self._sheets['CGP']
            for i in range(len(df_cgp)):
                fi_t_code = df_cgp.iloc[i, 9]
                if isinstance(fi_t_code, str) and not fi_t_code.startswith('*'):
                    start_val = df_cgp.iloc[i, 15]
                    if not (isinstance(start_val, float) and np.isnan(start_val)):
                        try:
                            yr = int(float(start_val))
                            if yr > 2000:
                                self._pwr_start[fi_t_code] = yr
                        except (ValueError, TypeError):
                            pass
                    comm_in = df_cgp.iloc[i, 10]
                    if isinstance(comm_in, str) and fi_t_code not in self._pwr_comm:
                        self._pwr_comm[fi_t_code] = comm_in

    def _parse_sol(self):
        """
        Parse SOL sheet.

        The sheet has two data blocks:
          Block 1 (first occurrence of each code): INVCOST / FIXOM / VAROM cost data
          Block 2 (second occurrence of each code): UC_RHSRT constraint values,
              stored in the same INVCOST~YEAR columns (year-keyed).
        """
        df = self._sheets['SOL']
        self._sol_data: dict[str, dict] = {}
        self._sol_uc_rhsrt: dict[str, dict] = {}   # code → {year: uc_rhsrt_value}

        # Headers in row 3 look like "INVCOST~2018" — extract year from suffix
        def _year_from_header(h) -> int | None:
            if not isinstance(h, str):
                return None
            if '~' in h:
                try:
                    return int(h.split('~')[-1])
                except ValueError:
                    return None
            return None

        headers = list(df.iloc[3])

        # Locate column ranges by header prefix
        invcost_cols = [(i, _year_from_header(headers[i]))
                        for i in range(len(headers))
                        if isinstance(headers[i], str) and
                        headers[i].startswith('INVCOST~')]
        fixom_cols   = [(i, _year_from_header(headers[i]))
                        for i in range(len(headers))
                        if isinstance(headers[i], str) and
                        headers[i].startswith('FIXOM~')]
        varom_cols   = [(i, _year_from_header(headers[i]))
                        for i in range(len(headers))
                        if isinstance(headers[i], str) and
                        headers[i].startswith('VAROM~')]

        SKIP_CODES = {'*', 'TechName', '*Technology name', '*Units',
                      '~FI_T:AF~UP', '~FI_T:AFA~UP'}

        for i in range(len(df)):
            # TechName is in col 9 for the data block (col 2 is process-def TechName)
            code = df.iloc[i, 9]
            if not isinstance(code, str):
                continue
            if code in SKIP_CODES:
                continue

            if code not in self._sol_data:
                # First occurrence → INVCOST/FIXOM/VAROM cost data
                self._sol_data[code] = {
                    'invcost': {yr: df.iloc[i, ci] for ci, yr in invcost_cols},
                    'fixom':   {yr: df.iloc[i, ci] for ci, yr in fixom_cols},
                    'varom':   {yr: df.iloc[i, ci] for ci, yr in varom_cols},
                }
            elif code not in self._sol_uc_rhsrt:
                # Second occurrence → UC_RHSRT constraint values.
                # The T-period columns in this block are shifted 1 to the left
                # of the INVCOST~YEAR columns (T08 sits at INVCOST~2018 - 1,
                # T09 sits at INVCOST~2018, etc.).  Using ci-1 aligns the
                # T-period values with the correct INVCOST year.
                self._sol_uc_rhsrt[code] = {
                    yr: df.iloc[i, ci - 1] for ci, yr in invcost_cols
                    if ci - 1 >= 0
                }

    def _parse_emi(self):
        """Parse EMI sheet: commodity → emission factor."""
        df = self._sheets['EMI']
        # Row 4 = commodity names, row 6 = PWRCO2 values
        commodities = df.iloc[4, 2:].values
        ef_values   = df.iloc[6, 2:].values
        self._emi: dict[str, float] = {}
        for comm, val in zip(commodities, ef_values):
            if isinstance(comm, str) and not pd.isna(val):
                self._emi[comm] = float(val)

    # ── Emission factor lookup ───────────────────────────────────────────────

    def _get_ef(self, code: str) -> object:
        """Look up emission factor (ktCO2/PJ) for a process."""
        EF_OVERRIDES = {
            'PWRWASWTE00': EF_MAP['PWRWAS'],   # 50 — waste is the primary fuel
            'PWRNGACCH11': EF_MAP['PWRNGA'],   # 56.1 — retrofit NGA plant
            'IRFNGACGP00': EF_MAP['PWRNGA'],
            'IBMNGACGP00': EF_MAP['PWRNGA'],
        }
        if code in EF_OVERRIDES:
            return EF_OVERRIDES[code]
        if code in SOLAR_PROCESSES:
            return 0
        comm = self._pwr_comm.get(code)
        if comm:
            if comm in self._emi:
                return self._emi[comm]
            if comm in EF_MAP:
                return EF_MAP[comm]
        return EMPTY

    # ── Commodity lookup ─────────────────────────────────────────────────────

    def _get_commodity(self, code: str) -> object:
        if code in COMMODITY_OVERRIDES:
            v = COMMODITY_OVERRIDES[code]
            return EMPTY if v is EMPTY else v
        return self._pwr_comm.get(code, EMPTY)

    # ── Row-year expansion ───────────────────────────────────────────────────

    def _get_capacity_rows(self, code: str) -> list[tuple[int, object]]:
        """
        Return list of (year, capacity_MW) tuples for this process.

        Rules:
        - Solar processes → expand by INVCOST change years, no capacity value
        - UP processes with RESID → deduplicate non-zero values, keep first year
        - FX processes with RESID → single row at year=2018, first non-zero cap
        - No RESID / forced no-cap → single row at year=2018, cap=EMPTY
        """
        if code in NO_CAPACITY_PROCS:
            return [(2018, EMPTY)]

        if code in SOLAR_PROCESSES:
            sol = self._sol_data.get(code, {})
            invcost = sol.get('invcost', {})
            years = sorted(yr for yr, v in invcost.items()
                           if not (isinstance(v, float) and np.isnan(v)))
            if not years:
                years = [2018]
            return [(yr, EMPTY) for yr in years]

        db = self._data_by[code]
        bound = BOUND_OVERRIDES.get(code, db['bound'])
        cap_by_year = db['capacity']

        valid = []
        for yr in self._cap_year_headers:
            v = cap_by_year.get(yr)
            if v is not None and not (isinstance(v, float) and np.isnan(v)) and v != 0:
                valid.append((int(yr), v))

        if not valid:
            return [(2018, EMPTY)]

        if isinstance(bound, str) and bound == 'FX':
            return [(2018, valid[0][1])]

        rows = []
        prev_val = None
        for yr, v in valid:
            if v != prev_val:
                rows.append((yr, v))
                prev_val = v
        return rows

    # ── Build all EndoRecord rows for one process ────────────────────────────

    def _build_rows(self, code: str) -> list[EndoRecord]:
        db = self._data_by[code]
        cap_rows = self._get_capacity_rows(code)

        ef        = self._get_ef(code)
        commodity = self._get_commodity(code)
        comm_share = COMMODITY_SHARE_OVERRIDES.get(
            code, 100 if commodity is not EMPTY else EMPTY)
        bound = BOUND_OVERRIDES.get(code, db['bound'])

        # START year: per-process from PWR/CGP sheet, default 2018
        start_year = self._pwr_start.get(code, 2018)

        # Grade = Bound type (UP/FX/NA) from Data_BY col H
        grade = str(bound).strip() if isinstance(bound, str) and bound.strip() else EMPTY

        # Efficiency: solar variants → EMPTY; cogen → electrical eff override
        raw_eff = db['efficiency']
        if code in SOLAR_VARIANTS:
            efficiency = EMPTY
        else:
            efficiency = COGEN_ELEC_EFF.get(
                code,
                float(raw_eff) if not (isinstance(raw_eff, float) and np.isnan(raw_eff))
                else EMPTY)

        # Heat rate → tech_efficiency field (EcoTEA AG col per mapping)
        # Solar processes have no heat rate
        raw_hr = db['heat_rate']
        tech_eff = (EMPTY if code in SOLAR_PROCESSES
                    else (float(raw_hr) if not (isinstance(raw_hr, float) and np.isnan(raw_hr))
                          else EMPTY))

        # Lifetime
        raw_lt = db['lifetime']
        lifetime = (int(raw_lt) if not (isinstance(raw_lt, float) and np.isnan(raw_lt))
                    else EMPTY)

        # CAP2ACT (afa field) = 31.536 for all ELE/SOL/CHP; 1 for PRE retrofit
        cap2act = CAP2ACT_PRE if code in COMMODITY_OVERRIDES and code == 'PWRNGACCH11' \
                  else CAP2ACT_ELE_SOL

        # Capacity type
        if code in NO_CAPACITY_PROCS:
            cap_type = EMPTY
        elif isinstance(bound, str) and bound in ('FX', 'UP'):
            cap_type = bound
        else:
            cap_type = EMPTY

        rows = []
        for (year, cap_mw) in cap_rows:
            # Cost values
            if code in SOLAR_PROCESSES:
                sol   = self._sol_data.get(code, {})
                capex = self._pick_cost(sol.get('invcost', {}), year)
                fixom = self._pick_cost(sol.get('fixom', {}), year)
                varom = EMPTY
            else:
                capex = self._pick_cost(db['invcost'], year)
                fixom = self._pick_cost(db['fixom'], year)
                varom = self._pick_cost(db['varom'], year)

            var_unit = EMPTY if code in NO_VAROM_UNIT_PROCS else 'PJ'

            # Capacity type for solar: use bound from Data_BY
            ct = (str(bound) if isinstance(bound, str) and bound in ('FX', 'UP')
                  else EMPTY) if code in SOLAR_PROCESSES else cap_type

            rows.append(EndoRecord(
                **TRACEABILITY,
                wp6_title='Power',
                process_code=code,
                description=str(db['description']).strip(),
                year=year,
                start_year=start_year,
                lifetime=lifetime,
                grade=grade,
                ef=ef,
                ef_unit='PJ',
                currency='MSGD2016',
                capex=capex,
                capex_unit='GW',
                fixed_opex=fixom,
                fixed_opex_unit='GW*yr',
                variable_opex=varom,
                variable_opex_unit=var_unit,
                efficiency=efficiency,
                tech_efficiency=tech_eff,   # Heat Rate → EcoTEA AG col
                commodity_share=comm_share,
                commodity=commodity,
                commodity_demand=EMPTY,     # Supply-side: no demand
                afa=cap2act,                # CAP2ACT = 31.536
                capacity=cap_mw,
                capacity_type=ct,
            ))
        return rows

    def _pick_cost(self, cost_dict: dict, year: int) -> object:
        """
        Return cost value for the given year.
        Uses exact match first, then falls back to the most recent prior year.
        Returns EMPTY if no valid value found.
        """
        if year in cost_dict:
            v = cost_dict[year]
            if not (isinstance(v, float) and np.isnan(v)):
                return v

        # Fallback: latest year ≤ requested year with a valid value
        candidates = [(yr, v) for yr, v in cost_dict.items()
                      if yr <= year and not (isinstance(v, float) and np.isnan(v))]
        if candidates:
            return max(candidates, key=lambda x: x[0])[1]
        return EMPTY
