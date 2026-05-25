"""
base_model.py — Abstract base class for all VT model converters.

To add a new model (e.g. VT_SG_IND):
  1. Create models/vt_sg_ind.py
  2. Subclass BaseConverter
  3. Implement extract_power_records() (and future sector methods)
  4. Register in engine.py
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional


MISSING = '-'   # Standard placeholder for missing/N-A values in EcoTEA
EMPTY = None    # True empty cell — used in WP Endo format where blank ≠ '-'


@dataclass
class PowerRecord:
    """One row in the EcoTEA Power sheet (one process × one year)."""

    # Traceability columns (A-G) — constant per model
    wp6_title: str = 'Power'
    data_owner: str = MISSING
    data_provider: str = MISSING
    data_source: str = MISSING
    data_source_desc: str = MISSING
    data_user: str = MISSING
    usage_purpose: str = MISSING

    # Dataset descriptor columns (H-M)
    process_code: str = MISSING          # H — technology/process
    description: str = MISSING          # I — technology/process description
    geography: str = MISSING            # J
    year: int = 2018                    # K — year of data
    start_year: int = 2018             # L — Technology Start Year
    lifetime: object = MISSING          # M — technology lifetime (years)

    # EcoTEA columns (N-AK)
    grade: str = MISSING                # N
    ef: object = MISSING               # O — emission factor
    ef_unit: str = 'PJ'               # P
    currency: str = 'MSGD2016'        # Q
    capex: object = MISSING            # R
    capex_unit: str = 'GW'            # S
    fixed_opex: object = MISSING       # T
    fixed_opex_unit: str = 'GW*yr(2018)'  # U
    variable_opex: object = MISSING    # V
    variable_opex_unit: str = 'PJ (2018)'  # W
    tax_cost: str = MISSING            # X
    sub_cost: str = MISSING            # Y
    efficiency: object = MISSING       # Z
    tech_efficiency: str = MISSING     # AA
    commodity_share: object = MISSING  # AB
    commodity: object = MISSING        # AC
    commodity_demand: str = MISSING    # AD
    interpolation_rule: object = MISSING  # AE
    afa: object = MISSING              # AF — capacity to activity factor
    heat_rate: object = MISSING        # AG
    capacity: object = MISSING         # AH
    capacity_type: object = MISSING    # AI
    constraint: str = MISSING          # AJ
    uc_rhsrt: object = MISSING         # AK — UC_RHSRT (GW), solar only
    uc_rhsrt_note: str = MISSING       # AL


@dataclass
class EndoRecord:
    """
    One row in the WP12345 EcoTEA Endo 'Sheet 1' (one process × one year).

    Column order matches the 61-column WP12345 template exactly.
    Use EMPTY (None) for cells that should be left blank; MISSING ('-') is
    reserved for EcoTEA-style dash placeholders.
    """

    # ── Traceability (cols 0-6) ───────────────────────────────────────────────
    wp6_title: object = EMPTY               # col 0  — WP6 Title
    data_owner: str = 'WP1'                # col 1
    data_provider: str = 'WP1'             # col 2
    data_source: str = EMPTY               # col 3
    data_source_desc: str = EMPTY          # col 4
    data_user: str = 'WP1'                 # col 5
    usage_purpose: str = 'Scenario analysis'  # col 6

    # ── Dataset descriptor (cols 7-12) ───────────────────────────────────────
    process_code: str = EMPTY              # col 7  — technology/process
    description: str = EMPTY              # col 8  — technology/process description
    geography: str = 'SG'                  # col 9
    year: int = 2018                       # col 10 — year of data
    start_year: object = 2018             # col 11 — Technology Start Year
    lifetime: object = EMPTY              # col 12 — technology lifetime (years)

    # ── EcoTEA core parameters (cols 13-22) ──────────────────────────────────
    grade: object = EMPTY                  # col 13 — Grade
    ef: object = 0                         # col 14 — emission factor
    ef_unit: str = 'PJ'                    # col 15 — ef ref unit
    currency: str = 'MSGD2016'             # col 16 — base currency
    capex: object = EMPTY                  # col 17 — capex
    capex_unit: str = 'GW'                 # col 18 — capex ref unit
    fixed_opex: object = EMPTY             # col 19 — fixed opex
    fixed_opex_unit: str = 'GW*yr'         # col 20 — fixed opex ref unit
    variable_opex: object = EMPTY          # col 21 — variable opex
    variable_opex_unit: str = 'PJ (2018)'  # col 22 — variable opex ref unit

    # ── WP-specific cost columns (cols 23-29) ────────────────────────────────
    currency_opex: object = EMPTY          # col 23 — Currency for OPEX
    opex: object = EMPTY                   # col 24 — OPEX
    tax_cost: object = EMPTY               # col 25 — Tax cost
    sub_cost: object = EMPTY               # col 26 — Sub cost
    fixed_annuity: object = EMPTY          # col 27 — Fixed annuity
    capex_annuity: object = EMPTY          # col 28 — CAPEX annuity
    start_cost: object = EMPTY             # col 29 — Start cost

    # ── Technology efficiency (cols 30-33) ───────────────────────────────────
    efficiency: object = EMPTY             # col 30 — efficiency
    efficiency_unit: object = EMPTY        # col 31 — efficiency ref unit
    tech_efficiency: object = EMPTY        # col 32 — Technology efficiency
    tech_efficiency_unit: object = EMPTY   # col 33 — Technology efficiency ref unit

    # ── Commodity (cols 34-36) ───────────────────────────────────────────────
    commodity_share: object = EMPTY        # col 34 — commodity share
    commodity: object = EMPTY             # col 35 — Commodity
    commodity_demand: object = EMPTY       # col 36 — Commodity demand

    # ── Additional WP descriptors (cols 37-55) ───────────────────────────────
    commodity_cost_base: object = EMPTY    # col 37
    commodity_cost_low: object = EMPTY     # col 38
    maintenance_rate: object = EMPTY       # col 39
    maintenance_schedule: object = EMPTY   # col 40
    mean_time_to_repair: object = EMPTY    # col 41
    interpolation_rule: object = EMPTY     # col 42
    afa: object = EMPTY                    # col 43 — capacity to activity factor
    heat_rate: object = EMPTY             # col 44 — heat rate
    forced_outage_rate: object = EMPTY     # col 45
    max_rating_factor: object = EMPTY      # col 46
    ramp_rate: object = EMPTY             # col 47
    min_stable_load: object = EMPTY        # col 48
    battery_capacity: object = EMPTY       # col 49
    max_range: object = EMPTY             # col 50
    max_ac_rate: object = EMPTY            # col 51
    max_dc_rate: object = EMPTY            # col 52
    fuel_type: object = EMPTY             # col 53 — Fuel Type
    ccs_rate: object = EMPTY              # col 54
    temperature_class: object = EMPTY      # col 55

    # ── Capacity constraints (cols 56-57) ────────────────────────────────────
    capacity: object = EMPTY              # col 56 — capacity
    capacity_type: object = EMPTY         # col 57 — capacity type (fixed or upper)

    # ── Spare columns (cols 58-60) — kept blank ──────────────────────────────
    extra1: object = EMPTY                 # col 58
    extra2: object = EMPTY                 # col 59
    extra3: object = EMPTY                 # col 60


class BaseConverter(ABC):
    """
    Abstract base class every model adapter must inherit from.

    Subclasses implement extract_*_records() for each sector.
    The engine calls the appropriate method based on target sheet.
    """

    def __init__(self, file_path: str):
        self.file_path = file_path
        self._sheets: dict = {}

    def _load_sheets(self):
        """Lazy-load all sheets from the source Excel file."""
        import pandas as pd
        if not self._sheets:
            self._sheets = pd.read_excel(self.file_path, sheet_name=None, header=None)

    @abstractmethod
    def extract_power_records(self):
        """
        Return all data records for this model.

        Depending on TARGET_WRITER, returns list[PowerRecord] (EcoTEA models)
        or list[EndoRecord] (WP Endo models).
        """
        pass

    # Future sector hooks (not yet implemented):
    # def extract_industry_records(self): pass
    # def extract_building_records(self): pass
