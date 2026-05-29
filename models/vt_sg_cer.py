"""
vt_sg_cer.py — Converter for VT_SG_CER_GMIT → WP12345 EcoTEA Endo (Sheet 1).

CER (Carbon trading / Emission Mixing) is the simplest SG-TIMES module:

  Source file structure (VT_SG_CER_GMIT.xlsx):
    - CER_MIXCO2 sheet: the only data sheet.
        ~FI_Process section (cols 1-8):  1 process definition.
        ~FI_T section (cols 10-18):      per-process parameters + Comm-IN list.
        ~FI_Comm section (interleaved):  1 output commodity (MIXCO2).

  Process: MIXCO2CAPTURE
    Sets = PRE (upstream/infrastructure, not demand-side)
    Description: 'MIX CO2 capture tech'
    EFF = 1   (1:1 CO2 aggregation, no loss)
    CAP2ACT = 1  (AFA)
    LIFE = 100 years
    No INVCOST / FIXOM / VAROM / demand / STOCK

  Output: 27 EndoRecord rows (even years 2018-2070), all with the same
  static parameter values.  Year is included so the output conforms to the
  same EcoTEA year-keyed format used by all other modules.

Source layout in CER_MIXCO2 (0-based row/col indices):
  Row 4: ~FI_T column headers at col 10-18:
          TechName | Comm-IN | Comm-OUT | EFF | CAP2ACT | AFA | LIFE | ACT_BND | STOCK
  Row 7: MIXCO2CAPTURE process row
          col  2 = TechName     ('MIXCO2CAPTURE')
          col  3 = TechDesc     ('MIX CO2 capture tech')
          col 13 = EFF          (1)
          col 14 = CAP2ACT/AFA  (1)
          col 15 = AFA          (1)
          col 16 = LIFE         (100)

Mapping reference: VT_CER_GMIT_to_EcoTEA_Mapping.xlsx
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
    data_source_desc='SG GREF v8.14; VT_SG_CER_GMIT',
    data_user='WP1',
    usage_purpose='Scenario analysis',
    geography='SG',
)

# Output years — 27 rows (even years, matching EcoTEA convention for all modules)
YEARS = list(range(2018, 2072, 2))

# CER_MIXCO2 sheet column indices (0-based) for the ~FI_T section
# Headers at row 4: TechName|Comm-IN|Comm-OUT|EFF|CAP2ACT|AFA|LIFE|ACT_BND|STOCK
COL_FIT_TECHNAME = 10
COL_FIT_COMM_IN  = 11
COL_FIT_COMM_OUT = 12
COL_FIT_EFF      = 13
COL_FIT_CAP2ACT  = 14
COL_FIT_AFA      = 15
COL_FIT_LIFE     = 16

# ~FI_Process section column indices (0-based)
COL_FIP_TECHNAME = 2
COL_FIP_TECHDESC = 3

# Data row index (0-based): the MIXCO2CAPTURE process row
PROCESS_ROW_IDX  = 7


# ─────────────────────────────────────────────────────────────────────────────
# Converter
# ─────────────────────────────────────────────────────────────────────────────

class VTSGCERConverter(BaseConverter):
    """
    Converts VT_SG_CER_GMIT → WP12345 EcoTEA Endo format.

    TARGET_WRITER = 'endo' signals engine.py to use wp_endo_writer instead of
    the default ecotea_writer.

    CER has no time-series parameters (no demand, INVCOST, FIXOM, VAROM or
    STOCK).  All field values are static constants read once from the
    CER_MIXCO2 sheet's ~FI_T row.  The converter still emits 27 EndoRecord
    rows (one per even year 2018-2070) so the output conforms to the same
    year-keyed structure used by every other module.
    """

    TARGET_WRITER = 'endo'

    def extract_power_records(self) -> list[EndoRecord]:
        self._load_sheets()
        return self._build_rows()

    # ── Parser + builder ──────────────────────────────────────────────────────

    def _build_rows(self) -> list[EndoRecord]:
        """
        Read the single MIXCO2CAPTURE process from CER_MIXCO2 and emit
        27 identical (except for year) EndoRecord rows.
        """
        df = self._sheets['CER_MIXCO2']

        def _val(col):
            v = df.iloc[PROCESS_ROW_IDX, col]
            return None if (isinstance(v, float) and np.isnan(v)) else v

        # Process identity — from ~FI_Process columns
        code = str(_val(COL_FIP_TECHNAME)).strip()
        desc = str(_val(COL_FIP_TECHDESC)).strip()

        # Parameters — from ~FI_T columns
        eff_raw  = _val(COL_FIT_EFF)
        afa_raw  = _val(COL_FIT_AFA)
        life_raw = _val(COL_FIT_LIFE)

        eff  = safe_float(eff_raw,  1.0)
        afa  = safe_float(afa_raw,  1.0)
        life = safe_int(life_raw, 100)

        rows = []
        for year in YEARS:
            rows.append(EndoRecord(
                **TRACEABILITY,
                process_code=code,
                description=desc,
                year=year,
                start_year=2018,
                lifetime=life,
                grade=EMPTY,
                ef=0,
                ef_unit='PJ',
                currency='MSGD2016',
                # No capex / opex / demand / capacity for CER
                capex=EMPTY,
                capex_unit='GW',
                fixed_opex=EMPTY,
                fixed_opex_unit='GW*yr',
                variable_opex=EMPTY,
                variable_opex_unit='PJ (2018)',
                efficiency=eff,
                tech_efficiency=EMPTY,
                commodity_share=100,
                commodity=EMPTY,
                commodity_demand=EMPTY,
                afa=afa,
            ))
        return rows
