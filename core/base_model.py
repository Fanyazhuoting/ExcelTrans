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
    def extract_power_records(self) -> list[PowerRecord]:
        """Return all PowerRecord rows for the EcoTEA Power sheet."""
        pass

    # Future sector hooks (not yet implemented):
    # def extract_industry_records(self): pass
    # def extract_building_records(self): pass
