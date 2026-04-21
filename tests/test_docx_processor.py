"""
tests/test_docx_processor.py
Unit tests for utils/docx_processor.py
"""

import os
import sys
import pytest
from io import BytesIO
from datetime import datetime

# Make the project root importable regardless of where pytest is invoked from.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from utils.docx_processor import (
    extract_data_from_quotation,
    fill_template,
    generate_output_filename,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_quotation_docx(paragraphs: list[str], tables: list[list[list[str]]] = None) -> str:
    """Write a temporary .docx and return its path."""
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    if tables:
        for table_data in tables:
            rows = len(table_data)
            cols = max(len(r) for r in table_data)
            tbl = doc.add_table(rows=rows, cols=cols)
            for r_idx, row_data in enumerate(table_data):
                for c_idx, cell_text in enumerate(row_data):
                    tbl.rows[r_idx].cells[c_idx].text = cell_text
    tmp = os.path.join(os.path.dirname(__file__), '_tmp_quotation.docx')
    doc.save(tmp)
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
# extract_data_from_quotation
# ---------------------------------------------------------------------------

class TestExtractData:

    def test_paragraph_full_width_colon(self, tmp_path):
        q = str(tmp_path / 'q.docx')
        doc = Document()
        doc.add_paragraph('客户名称：某某有限公司')
        doc.add_paragraph('报价编号：QT-2024-001')
        doc.save(q)

        data = extract_data_from_quotation(q)
        assert data['客户名称'] == '某某有限公司'
        assert data['报价编号'] == 'QT-2024-001'

    def test_paragraph_ascii_colon(self, tmp_path):
        q = str(tmp_path / 'q.docx')
        doc = Document()
        doc.add_paragraph('Project Name: Office Equipment')
        doc.save(q)

        data = extract_data_from_quotation(q)
        assert data['Project Name'] == 'Office Equipment'

    def test_empty_paragraph_ignored(self, tmp_path):
        q = str(tmp_path / 'q.docx')
        doc = Document()
        doc.add_paragraph('')
        doc.add_paragraph('   ')
        doc.add_paragraph('合计金额：10,000 元')
        doc.save(q)

        data = extract_data_from_quotation(q)
        assert '合计金额' in data
        assert len([k for k in data if not k.startswith('_')]) == 1

    def test_two_column_table(self, tmp_path):
        q = str(tmp_path / 'q.docx')
        doc = Document()
        tbl = doc.add_table(rows=2, cols=2)
        tbl.rows[0].cells[0].text = '项目名称'
        tbl.rows[0].cells[1].text = '智慧园区项目'
        tbl.rows[1].cells[0].text = '签约日期'
        tbl.rows[1].cells[1].text = '2024-05-01'
        doc.save(q)

        data = extract_data_from_quotation(q)
        assert data.get('项目名称') == '智慧园区项目'
        assert data.get('签约日期') == '2024-05-01'

    def test_multi_column_table_stored_as_private(self, tmp_path):
        q = str(tmp_path / 'q.docx')
        doc = Document()
        tbl = doc.add_table(rows=2, cols=3)
        headers = ['型号', '数量', '单价']
        for i, h in enumerate(headers):
            tbl.rows[0].cells[i].text = h
        for i, v in enumerate(['TypeA', '10', '500']):
            tbl.rows[1].cells[i].text = v
        doc.save(q)

        data = extract_data_from_quotation(q)
        private_keys = [k for k in data if k.startswith('_')]
        assert len(private_keys) > 0

    def test_paragraph_does_not_overwrite_table_value(self, tmp_path):
        """Paragraph extraction runs first; table uses setdefault, so paragraph wins."""
        q = str(tmp_path / 'q.docx')
        doc = Document()
        doc.add_paragraph('客户名称：段落中的公司')
        tbl = doc.add_table(rows=1, cols=2)
        tbl.rows[0].cells[0].text = '客户名称'
        tbl.rows[0].cells[1].text = '表格中的公司'
        doc.save(q)

        data = extract_data_from_quotation(q)
        assert data['客户名称'] == '段落中的公司'


# ---------------------------------------------------------------------------
# fill_template
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
        # name is like  QT-001_20240501_153045.docx
        stem = name[:-5]  # strip .docx
        parts = stem.rsplit('_', 2)
        assert len(parts) == 3
        date_part, time_part = parts[1], parts[2]
        datetime.strptime(date_part + time_part, '%Y%m%d%H%M%S')  # raises if invalid

    def test_name_truncated_to_50_chars(self):
        long_name = '甲' * 100
        data = {'报价编号': long_name}
        name = generate_output_filename(data)
        stem = name[:-5]   # strip .docx
        # stem is  <name_part>_YYYYMMDD_HHMMSS  – name_part ≤ 50 chars
        name_part = stem.rsplit('_', 2)[0]
        assert len(name_part) <= 50

    def test_private_keys_not_used_for_name(self):
        data = {'_internal': 'secret', '客户名称': '某公司'}
        name = generate_output_filename(data)
        assert name.startswith('某公司_')
