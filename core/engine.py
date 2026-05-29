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
    }


def get_available_models() -> list[str]:
    return list(_build_registry().keys())


def convert(model_name: str,
            vt_file_path: str,
            template_path: str,
            output_path: str) -> dict:
    """
    Run the full conversion pipeline.

    The writer used depends on the converter's TARGET_WRITER attribute:
      - 'endo'  → write_endo_output() into WP12345 'Sheet 1'
      - (other) → write_output()      into EcoTEA template sheet

    Returns a result dict with keys:
        success (bool), output_path (str), row_count (int), errors (list[str])
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
        converter = ConverterClass(vt_file_path)
        records = converter.extract_power_records()

        if not records:
            errors.append("No records were extracted from the source file. "
                          "Please verify it is a valid VT file.")
            return {'success': False, 'errors': errors}

        target_writer = getattr(ConverterClass, 'TARGET_WRITER', 'ecotea')

        if target_writer == 'endo':
            write_endo_output(
                records=records,
                template_path=template_path,
                output_path=output_path,
            )
        else:
            target_sheet = getattr(ConverterClass, 'TARGET_SHEET', 'Power')
            write_output(
                records=records,
                template_path=template_path,
                output_path=output_path,
                sheet_name=target_sheet,
            )

        return {
            'success': True,
            'output_path': str(output_path),
            'row_count': len(records),
            'errors': [],
        }

    except KeyError as e:
        errors.append(f"Missing expected sheet or column in VT file: {e}. "
                      f"Is this really a {model_name} file?")
    except Exception as e:
        errors.append(f"Conversion error: {type(e).__name__}: {e}")

    return {'success': False, 'errors': errors}
