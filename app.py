"""
app.py  –  QuoToCon  Flask application
========================================
Converts a Word-format quotation (报价单) into a filled contract (合同)
using a user-supplied Word template.

Routes
------
GET  /               – Upload form
POST /process        – Accept files, process, redirect to result
GET  /download/<fn>  – Serve generated contract file
"""

import os
import uuid
import logging

from flask import (
    Flask,
    render_template,
    request,
    send_file,
    redirect,
    url_for,
    flash,
)
from werkzeug.utils import secure_filename

from utils.docx_processor import (
    extract_data_from_quotation,
    fill_template,
    generate_output_filename,
)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s  %(levelname)s  %(name)s – %(message)s',
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Application setup
# ---------------------------------------------------------------------------
app = Flask(__name__)

# Secret key – override with SECRET_KEY env-var in production
app.secret_key = os.environ.get('SECRET_KEY', 'change-me-in-production')

# Runtime directories (relative to this file so the app works from any CWD)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB

ALLOWED_EXTENSIONS = {'docx'}

# Ensure runtime dirs exist on startup
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _allowed(filename: str) -> bool:
    return (
        '.' in filename
        and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
    )


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    # --- validate presence ---
    if 'quotation' not in request.files or 'template' not in request.files:
        flash('请同时上传报价单和合同模板文件。', 'error')
        return redirect(url_for('index'))

    quotation_file = request.files['quotation']
    template_file = request.files['template']

    if not quotation_file.filename or not template_file.filename:
        flash('请选择要上传的文件。', 'error')
        return redirect(url_for('index'))

    # --- validate types ---
    if not _allowed(quotation_file.filename):
        flash('报价单文件必须是 .docx 格式。', 'error')
        return redirect(url_for('index'))

    if not _allowed(template_file.filename):
        flash('合同模板文件必须是 .docx 格式。', 'error')
        return redirect(url_for('index'))

    # --- save with collision-safe names ---
    session_id = uuid.uuid4().hex
    quotation_path = os.path.join(UPLOAD_FOLDER, f'{session_id}_quotation.docx')
    template_path = os.path.join(UPLOAD_FOLDER, f'{session_id}_template.docx')

    quotation_file.save(quotation_path)
    template_file.save(template_path)

    try:
        data = extract_data_from_quotation(quotation_path)
        logger.info('Extracted %d field(s) from quotation.', len(data))

        output_filename = generate_output_filename(data)
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        fill_template(template_path, data, output_path)
        logger.info('Contract written to %s', output_path)

        # Expose only non-private fields to the template
        public_data = {k: v for k, v in data.items() if not k.startswith('_')}
        return render_template(
            'result.html',
            filename=output_filename,
            data=public_data,
        )

    except Exception as exc:
        logger.exception('Error while processing files.')
        flash(f'处理文件时出错：{exc}', 'error')
        return redirect(url_for('index'))

    finally:
        # Always clean up uploaded temp files
        for path in (quotation_path, template_path):
            try:
                os.remove(path)
            except OSError:
                pass


@app.route('/download/<path:filename>')
def download(filename: str):
    # Guard against path-traversal attacks
    safe = secure_filename(filename)
    if safe != filename:
        return '无效的文件名', 400

    output_path = os.path.join(app.config['OUTPUT_FOLDER'], safe)
    if not os.path.isfile(output_path):
        return '文件不存在或已过期，请重新生成。', 404

    return send_file(output_path, as_attachment=True, download_name=safe)


# ---------------------------------------------------------------------------
# Entry point (development only – use gunicorn in production)
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
