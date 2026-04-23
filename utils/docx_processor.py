"""
docx_processor.py
Utilities for extracting data from a Word quotation file and filling
a Word contract template with that data.

Quotation fields are detected by lines of the form:
    字段名：字段值
or by two-column tables where column 0 is the field name and column 1
is the value.

Contract templates use {{字段名}} placeholders. Every placeholder whose
name matches an extracted field is replaced with the corresponding value.
Unmatched placeholders are left unchanged so the user can see what data
is missing.
"""

import re
import os
from datetime import datetime

from docx import Document
import openpyxl


# ---------------------------------------------------------------------------
# Data extraction
# ---------------------------------------------------------------------------

def extract_data_from_quotation(docx_path: str) -> dict:
    """Return a dict of field→value pairs extracted from *docx_path*.

    Two-pass extraction:
    1. Paragraphs whose text matches ``key：value`` or ``key: value``.
    2. Tables:
       - Two-column tables are treated as key-value maps.
       - Multi-column tables are stored as ``_table_N`` metadata so that
         callers can inspect them (not used for placeholder replacement).
    """
    doc = Document(docx_path)
    data: dict = {}

    # --- paragraphs ---
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        # Match "Field Name：Value" (full-width or ASCII colon)
        match = re.match(r'^([^：:\n]{1,40})[：:]\s*(.+)$', text)
        if match:
            key = match.group(1).strip()
            value = match.group(2).strip()
            if key:
                data[key] = value

    # --- tables ---
    for table_idx, table in enumerate(doc.tables):
        rows = []
        for row in table.rows:
            # Deduplicate merged cells (python-docx repeats merged cells)
            cells = []
            seen = set()
            for cell in row.cells:
                cell_text = cell.text.strip()
                cell_id = id(cell._tc)
                if cell_id not in seen:
                    seen.add(cell_id)
                    cells.append(cell_text)
            rows.append(cells)

        if not rows:
            continue

        # Determine if this looks like a key-value table (≤2 unique columns)
        unique_widths = {len(r) for r in rows}
        if unique_widths <= {2}:
            for row in rows:
                if len(row) == 2 and row[0]:
                    key = row[0].rstrip('：: ').strip()
                    value = row[1].strip()
                    if key:
                        data.setdefault(key, value)
        else:
            # Store raw table data under a private key for future use
            data[f'_table_{table_idx}_headers'] = rows[0]
            data[f'_table_{table_idx}_rows'] = rows[1:]

    return data


def extract_data_from_quotation_xlsx(xlsx_path: str) -> dict:
    """Return a dict of field→value pairs extracted from *xlsx_path*.

    The quotation is expected to be a two-column table with full-width field
    names.  Extraction strategy (applied to every worksheet):

    1. Rows with two or more cells where the **first** cell is non-empty are
       treated as key-value pairs: the first cell is the key (trailing full-width
       colons and spaces are stripped), the second cell is the value.  Empty
       cells beyond the second column are ignored.
    2. Rows where only a single cell is non-empty and its text matches the
       full-width ``key：value`` pattern are also extracted.
    3. Rows with three or more non-empty cells (i.e. genuine multi-column data
       rows) are stored under the private keys ``_sheet_N_table_headers`` /
       ``_sheet_N_table_rows`` for future use.

    When the same key appears in multiple sheets, the first occurrence wins
    (``setdefault`` semantics).
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    data: dict = {}

    for sheet_idx, sheet in enumerate(wb.worksheets):
        table_rows: list = []

        for row in sheet.iter_rows(values_only=True):
            # Stringify and strip every cell, preserving positional order
            str_cells = [str(c).strip() if c is not None else '' for c in row]

            if not any(str_cells):
                continue

            first = str_cells[0] if str_cells else ''
            second = str_cells[1] if len(str_cells) > 1 else ''
            non_empty = [c for c in str_cells if c]

            if first:
                if second or len(str_cells) > 1:
                    # First cell is key; second cell is value (may be empty string)
                    key = first.rstrip('： ').strip()
                    value = second
                    if key and len(non_empty) <= 2:
                        # Normal key-value row
                        data.setdefault(key, value)
                    elif key and len(non_empty) > 2:
                        # Multiple non-empty cells → table row
                        table_rows.append(non_empty)
                else:
                    # Only one cell in the row; check for full-width "key：value"
                    match = re.match(r'^([^：\n]{1,40})：\s*(.+)$', first)
                    if match:
                        key = match.group(1).strip()
                        value = match.group(2).strip()
                        if key:
                            data.setdefault(key, value)

        if table_rows:
            data[f'_sheet_{sheet_idx}_table_headers'] = table_rows[0]
            data[f'_sheet_{sheet_idx}_table_rows'] = table_rows[1:]

    return data


# ---------------------------------------------------------------------------
# Template filling
# ---------------------------------------------------------------------------

def _replace_in_paragraph(para, replacements: dict) -> None:
    """Replace ``{{key}}`` placeholders in *para*, preserving run formatting.

    Word sometimes splits text across multiple runs inside a single paragraph,
    which means a placeholder like ``{{客户名称}}`` may be spread over several
    runs.  We work around this by:
    1. Concatenating all run texts.
    2. Doing all replacements on the combined string.
    3. Putting the result in the first run and clearing all subsequent runs.

    The first run's character formatting is kept; paragraph-level formatting
    (alignment, spacing, etc.) is untouched.
    """
    full_text = ''.join(run.text for run in para.runs)

    # Quick check – skip paragraph if no placeholder present
    if '{{' not in full_text:
        return

    new_text = full_text
    for key, value in replacements.items():
        placeholder = f'{{{{{key}}}}}'
        new_text = new_text.replace(placeholder, str(value) if value is not None else '')

    if new_text == full_text:
        return

    # Apply new text: put everything in run[0], blank all other runs
    if para.runs:
        para.runs[0].text = new_text
        for run in para.runs[1:]:
            run.text = ''
    else:
        # Edge case: paragraph has no runs – add one
        para.add_run(new_text)


def fill_template(template_path: str, data: dict, output_path: str) -> None:
    """Fill *template_path* with *data* and save the result to *output_path*.

    Placeholders in the template must use the ``{{字段名}}`` syntax.
    Only non-private keys (those not starting with ``_``) are used.
    """
    doc = Document(template_path)
    replacements = {k: v for k, v in data.items() if not k.startswith('_')}

    # Body paragraphs
    for para in doc.paragraphs:
        _replace_in_paragraph(para, replacements)

    # Table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _replace_in_paragraph(para, replacements)

    # Headers and footers
    for section in doc.sections:
        for para in section.header.paragraphs:
            _replace_in_paragraph(para, replacements)
        for para in section.footer.paragraphs:
            _replace_in_paragraph(para, replacements)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)


# ---------------------------------------------------------------------------
# Output filename generation
# ---------------------------------------------------------------------------

_INVALID_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]')

# Preferred fields to use when building the output filename, in priority order
_NAME_FIELDS = [
    '报价编号', '合同编号', '编号', '项目编号',
    '项目名称', '工程名称', '合同名称',
    '客户名称', '甲方名称', '买方名称',
]


def generate_output_filename(data: dict) -> str:
    """Return a safe .docx filename derived from extracted quotation data.

    Tries fields in *_NAME_FIELDS* priority order; falls back to "合同".
    A timestamp suffix is appended to guarantee uniqueness.
    """
    name_part = ''
    for field in _NAME_FIELDS:
        if data.get(field):
            name_part = str(data[field]).strip()
            break

    if not name_part:
        name_part = '合同'

    # Sanitise
    name_part = _INVALID_CHARS.sub('_', name_part)
    name_part = name_part.strip('_').strip()[:50]  # limit length

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return f'{name_part}_{timestamp}.docx'
