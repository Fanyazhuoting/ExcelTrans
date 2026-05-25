"""
engine.py — Orchestrates the conversion pipeline.

To register a new model:
  1. Import its converter class
  2. Add it to MODEL_REGISTRY with a display name as key
"""

from pathlib import Path

from core.ecotea_writer import write_output


# ── Model registry: display name → converter class ────────────────────────────
def _build_registry():
    from models.vt_sg_pwr import VTSGPWRConverter
    from models.vt_sg_pri import VTSGPRIConverter
    return {
        'VT_SG_PWR': VTSGPWRConverter,
        'VT_SG_PRI': VTSGPRIConverter,
    }


def get_available_models() -> list[str]:
    return list(_build_registry().keys())


def convert(model_name: str,
            vt_file_path: str,
            template_path: str,
            output_path: str) -> dict:
    """
    Run the full conversion pipeline.

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
