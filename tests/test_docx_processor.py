"""
tests/test_docx_processor.py
Unit tests for utils/docx_processor.py
"""

import io
import os
import sys
import pytest
from datetime import datetime

# Make the project root importable regardless of where pytest is invoked from.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from docx import Document
from utils.docx_processor import (
    extract_data_from_quotation_xlsx,
    fill_template,
    generate_output_filename,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_xlsx(rows: list[list], sheet_name: str = 'Sheet1') -> str:
    """Write a temporary .xlsx with given rows and return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for row in rows:
        ws.append(row)
    tmp = os.path.join(os.path.dirname(__file__), '_tmp_quotation.xlsx')
    wb.save(tmp)
    return tmp


def _make_template_docx(paragraphs: list[str]) -> str:
    """Write a temporary template .docx and return its path."""
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    tmp = os.path.join(os.path.dirname(__file__), '_tmp_template.docx')
    doc.save(tmp)
    return tmp


def _cleanup(*paths):
    for p in paths:
        try:
            os.remove(p)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# extract_data_from_quotation_xlsx
# ---------------------------------------------------------------------------

class TestExtractDataXlsx:

    def test_two_column_table_fullwidth_key(self, tmp_path):
        """Standard two-column table with full-width field names."""
        q = str(tmp_path / 'q.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['客户名称：', '某某有限公司'])
        ws.append(['报价编号：', 'QT-2024-001'])
        wb.save(q)

        data = extract_data_from_quotation_xlsx(q)
        assert data['客户名称'] == '某某有限公司'
        assert data['报价编号'] == 'QT-2024-001'

    def test_two_column_table_key_without_colon(self, tmp_path):
        """Full-width field name without trailing colon."""
        q = str(tmp_path / 'q.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['项目名称', '智慧园区项目'])
        ws.append(['签约日期', '2024-05-01'])
        wb.save(q)

        data = extract_data_from_quotation_xlsx(q)
        assert data['项目名称'] == '智慧园区项目'
        assert data['签约日期'] == '2024-05-01'

    def test_single_cell_fullwidth_colon_pattern(self, tmp_path):
        """Single cell with full-width 'key：value' format."""
        q = str(tmp_path / 'q.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['合计金额：46,000 元'])
        wb.save(q)

        data = extract_data_from_quotation_xlsx(q)
        assert data.get('合计金额') == '46,000 元'

    def test_empty_rows_ignored(self, tmp_path):
        """Completely empty rows are skipped."""
        q = str(tmp_path / 'q.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append([None, None])
        ws.append(['客户名称：', '测试公司'])
        ws.append([None])
        wb.save(q)

        data = extract_data_from_quotation_xlsx(q)
        assert data['客户名称'] == '测试公司'
        assert len([k for k in data if not k.startswith('_')]) == 1

    def test_multi_column_rows_stored_as_private(self, tmp_path):
        """Rows with 3+ non-empty cells become private table metadata."""
        q = str(tmp_path / 'q.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['型号', '数量', '单价'])
        ws.append(['TypeA', '10', '500'])
        wb.save(q)

        data = extract_data_from_quotation_xlsx(q)
        private_keys = [k for k in data if k.startswith('_')]
        assert len(private_keys) > 0

    def test_first_sheet_value_wins(self, tmp_path):
        """When the same key appears on multiple sheets, first wins."""
        q = str(tmp_path / 'q.xlsx')
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = 'Sheet1'
        ws1.append(['客户名称：', '第一页公司'])
        ws2 = wb.create_sheet('Sheet2')
        ws2.append(['客户名称：', '第二页公司'])
        wb.save(q)

        data = extract_data_from_quotation_xlsx(q)
        assert data['客户名称'] == '第一页公司'

    def test_numeric_value_converted_to_string(self, tmp_path):
        """Numeric cell values are stringified."""
        q = str(tmp_path / 'q.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['合计金额：', 46000])
        wb.save(q)

        data = extract_data_from_quotation_xlsx(q)
        assert data['合计金额'] == '46000'

    def test_single_cell_halfwidth_colon_not_matched(self, tmp_path):
        """Half-width colon in a single cell is NOT extracted (full-width only)."""
        q = str(tmp_path / 'q.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['客户名称: 某公司'])   # half-width colon
        wb.save(q)

        data = extract_data_from_quotation_xlsx(q)
        assert '客户名称' not in data


# ---------------------------------------------------------------------------
# fill_template  (unchanged – kept for regression)
# ---------------------------------------------------------------------------

class TestFillTemplate:

    def test_basic_replacement(self, tmp_path):
        tmpl = str(tmp_path / 't.docx')
        out = str(tmp_path / 'out.docx')
        doc = Document()
        doc.add_paragraph('甲方：{{客户名称}}')
        doc.add_paragraph('编号：{{报价编号}}')
        doc.save(tmpl)

        data = {'客户名称': '某某公司', '报价编号': 'QT-001'}
        fill_template(tmpl, data, out)

        result = Document(out)
        texts = [p.text for p in result.paragraphs]
        assert '甲方：某某公司' in texts
        assert '编号：QT-001' in texts

    def test_unmatched_placeholder_retained(self, tmp_path):
        tmpl = str(tmp_path / 't.docx')
        out = str(tmp_path / 'out.docx')
        doc = Document()
        doc.add_paragraph('合同金额：{{合计金额}}')
        doc.save(tmpl)

        fill_template(tmpl, {}, out)

        result = Document(out)
        texts = [p.text for p in result.paragraphs]
        assert any('{{合计金额}}' in t for t in texts)

    def test_private_keys_not_replaced(self, tmp_path):
        tmpl = str(tmp_path / 't.docx')
        out = str(tmp_path / 'out.docx')
        doc = Document()
        doc.add_paragraph('{{_internal}}')
        doc.save(tmpl)

        fill_template(tmpl, {'_internal': '敏感数据'}, out)

        result = Document(out)
        texts = [p.text for p in result.paragraphs]
        assert any('{{_internal}}' in t for t in texts)

    def test_replacement_in_table_cell(self, tmp_path):
        tmpl = str(tmp_path / 't.docx')
        out = str(tmp_path / 'out.docx')
        doc = Document()
        tbl = doc.add_table(rows=1, cols=2)
        tbl.rows[0].cells[0].text = '金额'
        tbl.rows[0].cells[1].text = '{{合计金额}}'
        doc.save(tmpl)

        fill_template(tmpl, {'合计金额': '99,000 元'}, out)

        result = Document(out)
        cell_text = result.tables[0].rows[0].cells[1].text
        assert cell_text == '99,000 元'

    def test_output_directory_created_if_missing(self, tmp_path):
        tmpl = str(tmp_path / 't.docx')
        doc = Document()
        doc.add_paragraph('test')
        doc.save(tmpl)

        nested_out = str(tmp_path / 'subdir' / 'deep' / 'out.docx')
        fill_template(tmpl, {}, nested_out)
        assert os.path.isfile(nested_out)


# ---------------------------------------------------------------------------
# generate_output_filename
# ---------------------------------------------------------------------------

class TestGenerateOutputFilename:

    def test_uses_quotation_number_first(self):
        data = {'报价编号': 'QT-2024-001', '客户名称': '某公司'}
        name = generate_output_filename(data)
        assert name.startswith('QT-2024-001_')
        assert name.endswith('.docx')

    def test_falls_back_to_project_name(self):
        data = {'项目名称': '智慧园区'}
        name = generate_output_filename(data)
        assert name.startswith('智慧园区_')

    def test_falls_back_to_default(self):
        name = generate_output_filename({})
        assert name.startswith('合同_')

    def test_invalid_chars_sanitised(self):
        data = {'报价编号': 'QT/2024\\001:test'}
        name = generate_output_filename(data)
        assert '/' not in name
        assert '\\' not in name
        assert ':' not in name

    def test_timestamp_suffix_format(self):
        data = {'报价编号': 'QT-001'}
        name = generate_output_filename(data)
        stem = name[:-5]  # strip .docx
        parts = stem.rsplit('_', 2)
        assert len(parts) == 3
        date_part, time_part = parts[1], parts[2]
        datetime.strptime(date_part + time_part, '%Y%m%d%H%M%S')

    def test_name_truncated_to_50_chars(self):
        long_name = '甲' * 100
        data = {'报价编号': long_name}
        name = generate_output_filename(data)
        stem = name[:-5]
        name_part = stem.rsplit('_', 2)[0]
        assert len(name_part) <= 50

    def test_private_keys_not_used_for_name(self):
        data = {'_internal': 'secret', '客户名称': '某公司'}
        name = generate_output_filename(data)
        assert name.startswith('某公司_')
