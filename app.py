"""
app.py — Flask web server for the VT → EcoTEA conversion tool.
"""

import os
import sys
import tempfile
import threading
import webbrowser
from pathlib import Path

from flask import (Flask, jsonify, render_template, request,
                   send_file, send_from_directory)

# Ensure local modules are importable
sys.path.insert(0, str(Path(__file__).parent))

from core.engine import convert, get_available_models

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB upload limit

TEMPLATE_PATH = Path(__file__).parent / 'assets' / 'ecotea_template.xlsx'
UPLOAD_DIR = Path(__file__).parent / 'uploads'
OUTPUT_DIR = Path(__file__).parent / 'outputs'

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html', models=get_available_models())


@app.route('/convert', methods=['POST'])
def run_conversion():
    # Validate VT file upload
    if 'vt_file' not in request.files:
        return jsonify({'success': False, 'errors': ['未上传 VT 源文件。']}), 400
    f = request.files['vt_file']
    if not f.filename:
        return jsonify({'success': False, 'errors': ['VT 文件名为空。']}), 400
    if not f.filename.lower().endswith(('.xlsx', '.xlsm', '.xls')):
        return jsonify({'success': False,
                        'errors': ['VT 文件请上传 Excel 格式 (.xlsx)。']}), 400

    # Validate EcoTEA template upload
    if 'ecotea_file' not in request.files:
        return jsonify({'success': False, 'errors': ['未上传 EcoTEA 目标文件。']}), 400
    eco = request.files['ecotea_file']
    if not eco.filename:
        return jsonify({'success': False, 'errors': ['EcoTEA 文件名为空。']}), 400
    if not eco.filename.lower().endswith(('.xlsx', '.xlsm', '.xls')):
        return jsonify({'success': False,
                        'errors': ['EcoTEA 文件请上传 Excel 格式 (.xlsx)。']}), 400

    model_name = request.form.get('model_type', 'VT_SG_PWR')

    # Save uploaded files
    vt_path = UPLOAD_DIR / f.filename
    f.save(str(vt_path))

    eco_path = UPLOAD_DIR / eco.filename
    eco.save(str(eco_path))

    # Output file path
    stem = Path(f.filename).stem
    output_path = OUTPUT_DIR / f'EcoTEA_{stem}_converted.xlsx'

    result = convert(
        model_name=model_name,
        vt_file_path=str(vt_path),
        template_path=str(eco_path),
        output_path=str(output_path),
    )

    if result['success']:
        return jsonify({
            'success': True,
            'row_count': result['row_count'],
            'download_name': output_path.name,
        })
    else:
        return jsonify({'success': False, 'errors': result['errors']}), 500


@app.route('/download/<filename>')
def download_file(filename):
    # Sanitise filename to prevent path traversal
    filename = Path(filename).name
    return send_from_directory(str(OUTPUT_DIR), filename, as_attachment=True)


def open_browser():
    webbrowser.open('http://127.0.0.1:5000')


if __name__ == '__main__':
    print("=" * 55)
    print("  VT → EcoTEA Conversion Tool")
    print("  Opening browser at http://127.0.0.1:5000")
    print("  Press Ctrl+C to stop the server")
    print("=" * 55)
    threading.Timer(1.2, open_browser).start()
    app.run(debug=False, port=5000)
