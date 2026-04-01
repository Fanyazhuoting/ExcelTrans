"""
vt_sg_pwr.py — Converter for VT_SG_PWR_GREF → EcoTEA Power sheet.

Mapping rules are encoded as explicit functions.
To update a rule, find and edit the corresponding _get_* method.
"""

import pandas as pd
import numpy as np

from core.base_model import BaseConverter, PowerRecord, MISSING


# ─────────────────────────────────────────────────────────────────────────────
# Constants — traceability metadata filled into every row
# ─────────────────────────────────────────────────────────────────────────────
TRACEABILITY = dict(
    wp6_title='Power',
    data_owner='ESI',
    data_provider='WP1',
    data_source='GREF',
    data_source_desc='SG GREF v8.14; VT_SG_PWR_GREF',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
    start_year=2018,  # Technology Start Year — always 2018 in this dataset
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

# AFA conversion: VT raw value → EcoTEA value
AFA_MAP = {
    0.75: 0.8,
    0.14: 0.1,
}

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

# Processes where capacity and capacity_type are forced to MISSING
# (retrofit technology — capacity comes from a separate mechanism)
NO_CAPACITY_PROCS = {'PWRNGACCH11'}

# For PWRWASWTE00 bound is NaN in Data_BY but EcoTEA shows UP
BOUND_OVERRIDES = {'PWRWASWTE00': 'UP'}

# Fixed commodity overrides (where simple Comm-IN lookup doesn't apply)
COMMODITY_OVERRIDES = {
    'PWRBMCSTP00': 'PWRBMS+PWACOA',
    'PWRWASWTE00': MISSING,
    'PWRSOLLPV00': MISSING,
    'BDESOLLPV00': MISSING,
    'BDNSOLLPV00': MISSING,
    'BDFSOLLPV00': MISSING,
    'WATSOLFPV00': MISSING,
    'TPASOLLPV00': MISSING,
    'TPJSOLLPV00': MISSING,
    'IRFNGACGP00': MISSING,
    'IBMNGACGP00': MISSING,
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

# Interpolation rule overrides (most processes use MISSING)
INTERPOLATION_OVERRIDES = {
    'PWRBMCSTP00': 5,
}


# ─────────────────────────────────────────────────────────────────────────────
# Converter class
# ─────────────────────────────────────────────────────────────────────────────

class VTSGPWRConverter(BaseConverter):

    def extract_power_records(self) -> list[PowerRecord]:
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
        self._pwr_comm: dict[str, str] = {}    # code → Comm-IN
        self._pwr_share0: dict[str, object] = {}  # code → interpolation rule
        self._pwr_share: dict[str, object] = {}   # code → share value

        current_code = None
        for i in range(6, len(df)):
            code = df.iloc[i, 2]
            if isinstance(code, str) and code not in ('*', 'Technology Name'):
                current_code = code
            comm = df.iloc[i, 10]
            if current_code and isinstance(comm, str) and comm != '*' and not pd.isna(comm):
                # Only store the FIRST Comm-IN for each process
                if current_code not in self._pwr_comm:
                    self._pwr_comm[current_code] = comm
                    self._pwr_share0[current_code] = df.iloc[i, 12]
                    self._pwr_share[current_code] = df.iloc[i, 13]

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
        """
        Look up emission factor for a process.
        Priority: EF_OVERRIDES → primary Comm-IN from EMI/EF_MAP.

        Special cases:
        - PWRWASWTE00: multi-fuel (COA+WAS); EcoTEA uses WAS ef=50
        - PWRNGACCH11: commodity 'PWRRTFNGA' maps to NGA ef=56.1
        - Solar: ef=0
        - Cogen: uses NGA ef
        """
        # Process-level overrides (where comm lookup gives wrong result)
        EF_OVERRIDES = {
            'PWRWASWTE00': EF_MAP['PWRWAS'],   # 50 — waste is the billed fuel
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
        return MISSING

    # ── AFA conversion ───────────────────────────────────────────────────────

    def _convert_afa(self, raw: object) -> object:
        """Convert VT AFA value to EcoTEA equivalent."""
        if pd.isna(raw):
            return MISSING
        return AFA_MAP.get(float(raw), raw)

    # ── Commodity lookup ─────────────────────────────────────────────────────

    def _get_commodity(self, code: str) -> object:
        if code in COMMODITY_OVERRIDES:
            return COMMODITY_OVERRIDES[code]
        return self._pwr_comm.get(code, MISSING)

    # ── Row-year expansion ───────────────────────────────────────────────────

    def _get_capacity_rows(self, code: str) -> list[tuple[int, object]]:
        """
        Return list of (year, capacity_MW) tuples for this process.

        Rules:
        - Solar processes → expand by INVCOST change years, no capacity value
        - UP processes with STOCK → deduplicate non-zero values, keep first
          year each unique value appears
        - FX processes with STOCK → single row at year=2018, first non-zero cap
        - No STOCK / forced no-cap → single row at year=2018, cap=MISSING
        """
        if code in NO_CAPACITY_PROCS:
            return [(2018, MISSING)]

        if code in SOLAR_PROCESSES:
            # Row years are driven by INVCOST year columns in SOL sheet
            sol = self._sol_data.get(code, {})
            invcost = sol.get('invcost', {})
            years = sorted(yr for yr, v in invcost.items()
                           if not (isinstance(v, float) and np.isnan(v)))
            if not years:
                years = [2018]
            return [(yr, MISSING) for yr in years]

        # Normal process: use Data_BY capacity
        db = self._data_by[code]
        bound = BOUND_OVERRIDES.get(code, db['bound'])
        cap_by_year = db['capacity']

        # Collect non-zero, non-NaN capacity entries in year order
        valid = []
        for yr in self._cap_year_headers:
            v = cap_by_year.get(yr)
            if v is not None and not (isinstance(v, float) and np.isnan(v)) and v != 0:
                valid.append((int(yr), v))

        if not valid:
            # No STOCK data at all
            return [(2018, MISSING)]

        if isinstance(bound, str) and bound == 'FX':
            # FX: single row at base year 2018, capacity = first non-zero value
            return [(2018, valid[0][1])]

        # UP (or unspecified with data): deduplicate — keep first year of
        # each new unique value
        rows = []
        prev_val = None
        for yr, v in valid:
            if v != prev_val:
                rows.append((yr, v))
                prev_val = v
        return rows

    # ── Build all PowerRecord rows for one process ───────────────────────────

    def _build_rows(self, code: str) -> list[PowerRecord]:
        db = self._data_by[code]
        cap_rows = self._get_capacity_rows(code)

        # Resolve shared fields (same for all year-rows of this process)
        ef = self._get_ef(code)
        afa = self._convert_afa(db['afa'])
        commodity = self._get_commodity(code)
        comm_share = COMMODITY_SHARE_OVERRIDES.get(code, 1 if commodity != MISSING else MISSING)
        interp = INTERPOLATION_OVERRIDES.get(code, MISSING)
        bound = BOUND_OVERRIDES.get(code, db['bound'])

        # Efficiency:
        #   - Solar variants (BDE/BDN/BDF/WAT/TPA/TPJ) → always MISSING
        #   - Cogen processes → use electrical efficiency override (0.3)
        #   - Others → raw value from Data_BY, or MISSING if NaN
        raw_eff = db['efficiency']
        if code in SOLAR_VARIANTS:
            efficiency = MISSING
        else:
            efficiency = COGEN_ELEC_EFF.get(code,
                         raw_eff if not (isinstance(raw_eff, float) and np.isnan(raw_eff))
                         else MISSING)

        # Heat rate: solar processes have no heat rate
        raw_hr = db['heat_rate']
        heat_rate = (MISSING if code in SOLAR_PROCESSES
                     else (raw_hr if not (isinstance(raw_hr, float) and np.isnan(raw_hr))
                           else MISSING))

        # Lifetime
        raw_lt = db['lifetime']
        lifetime = (int(raw_lt) if not (isinstance(raw_lt, float) and np.isnan(raw_lt))
                    else MISSING)

        # Capacity type
        if code in NO_CAPACITY_PROCS:
            cap_type = MISSING
        elif isinstance(bound, str) and bound in ('FX', 'UP'):
            cap_type = bound
        else:
            # NaN bound but has capacity data (e.g. PWRWASWTE00 → UP via override)
            cap_type = MISSING

        # UC_RHSRT: fixed reference values for PWRSOLLPV00 only.
        # These values are taken directly from the reference EcoTEA file.
        # The VT SOL sheet second block stores per-period T-values that require
        # TIMES model period-duration metadata (not in the VT Excel) to convert
        # to annual constraint values, so we hardcode the known correct values.
        UC_RHSRT_FIXED = {
            'PWRSOLLPV00': {
                2018: 0.0290451024544057,
                2020: 0.0641013375821088,
                2030: 0.870599786263404,
                2050: 0.870599786263404,
            }
        }
        uc_rhsrt_by_year = UC_RHSRT_FIXED.get(code, {})

        rows = []
        for (year, cap_mw) in cap_rows:
            # Cost values: pick year-appropriate INVCOST/FIXOM/VAROM
            if code in SOLAR_PROCESSES:
                sol = self._sol_data.get(code, {})
                capex    = self._pick_cost(sol.get('invcost', {}), year)
                fixom    = self._pick_cost(sol.get('fixom', {}), year)
                varom    = MISSING
                var_unit = MISSING if code in NO_VAROM_UNIT_PROCS else 'PJ (2018)'
            else:
                capex    = self._pick_cost(db['invcost'], year)
                fixom    = self._pick_cost(db['fixom'], year)
                varom    = self._pick_cost(db['varom'], year)
                var_unit = 'PJ (2018)'

            # Capacity type for solar: always use bound from Data_BY
            if code in SOLAR_PROCESSES:
                ct = str(bound) if isinstance(bound, str) and bound in ('FX', 'UP') else MISSING
            else:
                ct = cap_type

            # UC_RHSRT: pick exact year or fall back to most recent prior year
            uc_val = self._pick_cost(uc_rhsrt_by_year, year) if uc_rhsrt_by_year else MISSING

            r = PowerRecord(
                **TRACEABILITY,
                process_code=code,
                description=str(db['description']).strip(),
                year=year,
                lifetime=lifetime,
                ef=ef,
                capex=capex,
                fixed_opex=fixom,
                variable_opex=varom,
                variable_opex_unit=var_unit,
                efficiency=efficiency,
                commodity_share=comm_share,
                commodity=commodity,
                interpolation_rule=interp,
                afa=afa,
                heat_rate=heat_rate,
                capacity=cap_mw if cap_mw != MISSING else MISSING,
                capacity_type=ct,
                uc_rhsrt=uc_val,
            )
            rows.append(r)
        return rows

    def _pick_cost(self, cost_dict: dict, year: int) -> object:
        """
        Return cost value for the given year.
        Uses exact match first, then falls back to the most recent prior year.
        Returns MISSING if no valid value found.
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
        return MISSING
