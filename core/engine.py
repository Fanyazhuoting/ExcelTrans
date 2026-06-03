"""
engine.py — Orchestrates the conversion pipeline.

To register a new model:
  1. Import its converter class
  2. Add it to MODEL_REGISTRY with a display name as key

Writer dispatch:
  - Converters with TARGET_WRITER = 'endo' → wp_endo_writer.write_endo_output()
  - All others → ecotea_writer.write_output()  (default)
"""

from pathlib import Path

from core.ecotea_writer import write_output
from core.wp_endo_writer import write_endo_output


# ── Model registry: display name → converter class ────────────────────────────
def _build_registry():
    from models.vt_sg_pwr import VTSGPWRConverter
    from models.vt_sg_pri import VTSGPRIConverter
    from models.vt_sg_oth import VTSGOTHConverter
    from models.vt_sg_agr import VTSGAGRConverter
    from models.vt_sg_bld import VTSGBLDConverter
    from models.vt_sg_cer import VTSGCERConverter
    from models.vt_sg_hfc import VTSGHFCConverter
    from models.vt_sg_hhd import VTSGHHDConverter
    from models.vt_sg_ifc import VTSGIFCConverter
    from models.vt_sg_ind import VTSGINDConverter
    from models.vt_sg_trn import VTSGTRNConverter
    from models.vt_sg_wat import VTSGWATConverter
    from models.vt_sg_wst import VTSGWSTConverter
    from models.vt_sg_hyp import VTSGHYPConverter
    return {
        'VT_SG_PWR': VTSGPWRConverter,
        'VT_SG_PRI': VTSGPRIConverter,
        'VT_SG_OTH': VTSGOTHConverter,
        'VT_SG_AGR': VTSGAGRConverter,
        'VT_SG_BLD': VTSGBLDConverter,
        'VT_SG_CER': VTSGCERConverter,
        'VT_SG_HFC': VTSGHFCConverter,
        'VT_SG_HHD': VTSGHHDConverter,
        'VT_SG_IFC': VTSGIFCConverter,
        'VT_SG_IND': VTSGINDConverter,
        'VT_SG_TRN': VTSGTRNConverter,
        'VT_SG_WAT': VTSGWATConverter,
        'VT_SG_WST': VTSGWSTConverter,
        'VT_SG_HYP': VTSGHYPConverter,
    }


def get_available_models() -> list[str]:
    return list(_build_registry().keys())


def convert(model_name: str,
            vt_source,
            template_source) -> dict:
    """
    Run the full conversion pipeline entirely in memory.

    Args:
        model_name:      Registry key for the converter model.
        vt_source:       File path (str/Path) or BytesIO of the VT source file.
        template_source: File path (str/Path) or BytesIO of the EcoTEA template.

    Returns a result dict with keys:
        success (bool), output (BytesIO), row_count (int), errors (list[str])
    """
    registry = _build_registry()
    if model_name not in registry:
        return {
            'success': False,
            'errors': [f"Unknown model '{model_name}'. "
                       f"Available: {list(registry.keys())}"],
        }

    errors = []
    try:
        ConverterClass = registry[model_name]
        converter = ConverterClass(vt_source)
        records = converter.extract_power_records()

        if not records:
            errors.append("No records were extracted from the source file. "
                          "Please verify it is a valid VT file.")
            return {'success': False, 'errors': errors}

        target_writer = getattr(ConverterClass, 'TARGET_WRITER', 'ecotea')

        if target_writer == 'endo':
            output = write_endo_output(
                records=records,
                template_source=template_source,
            )
        else:
            target_sheet = getattr(ConverterClass, 'TARGET_SHEET', 'Power')
            output = write_output(
                records=records,
                template_source=template_source,
                sheet_name=target_sheet,
            )

        return {
            'success': True,
            'output': output,
            'row_count': len(records),
            'errors': [],
        }

    except KeyError as e:
        errors.append(f"Missing expected sheet or column in VT file: {e}. "
                      f"Is this really a {model_name} file?")
    except Exception as e:
        errors.append(f"Conversion error: {type(e).__name__}: {e}")

    return {'success': False, 'errors': errors}
