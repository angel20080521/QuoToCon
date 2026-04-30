"""
docx_processor.py
Utilities for extracting data from a Word or Excel quotation file and filling
a Word contract template with that data.

Excel quotation format (报价单模板.xlsx):
  - Rows 1-5: header area with multiple field-value pairs per row.
    Labels end with a full-width colon '：' and the following cell is the value.
  - Row 6: product table header (编号 / 商品种类 / 商品名称 / 数量 / ...)
  - Rows 7+: product data rows (software/hardware at 13 % tax, services at 6 %)
  - Summary row: 未税合计 / 税额 / 含税合计

Contract templates (合同模板.docx) use {{字段名}} placeholders. Every
placeholder whose name matches an extracted field is replaced with the
corresponding value.  Unmatched placeholders are left unchanged so the user
can see what data is missing.
"""

import re
import os
from copy import deepcopy
from datetime import datetime

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import openpyxl


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _extract_kv_from_row(str_cells: list) -> dict:
    """Scan a row for cells ending with full-width '：' and return a dict of
    key→value pairs.

    The cell immediately following a label cell is taken as the value (even if
    empty).  If two labels are adjacent, the first gets an empty value.
    """
    result: dict = {}
    i = 0
    while i < len(str_cells):
        cell = str_cells[i]
        if cell and cell.endswith('：'):
            key = cell.rstrip('：').strip()
            if i + 1 < len(str_cells):
                next_cell = str_cells[i + 1]
                # If the next cell is also a label, current key has no value
                if next_cell and next_cell.endswith('：'):
                    value = ''
                    i += 1  # only skip the current label
                else:
                    value = next_cell
                    i += 2  # skip label AND value
            else:
                value = ''
                i += 1
            if key:
                result[key] = value
        else:
            i += 1
    return result


def _four_digit_chinese(n: int) -> str:
    """Convert a 1-to-4 digit integer to Chinese financial characters.

    Leading zeros within the group are represented with 零 (inserted by the
    caller when crossing group boundaries, not here).
    """
    CN_NUM = '零壹贰叁肆伍陆柒捌玖'
    units = [(1000, '仟'), (100, '佰'), (10, '拾'), (1, '')]
    s = ''
    prev_zero = False
    for val, unit in units:
        d = n // val
        n %= val
        if d:
            if prev_zero:
                s += '零'
                prev_zero = False
            s += CN_NUM[d] + unit
        else:
            if s:  # suppress leading zeros
                prev_zero = True
    return s


def _amount_to_chinese(amount) -> str:
    """Convert a numeric amount to Chinese financial (大写) notation.

    Examples:
        12345.67  →  '壹万贰仟叁佰肆拾伍元陆角柒分'
        10000.00  →  '壹万元整'
        0.05      →  '零元零伍分'
    """
    if amount is None:
        return ''
    try:
        amount = float(amount)
    except (ValueError, TypeError):
        return str(amount)

    CN_NUM = '零壹贰叁肆伍陆柒捌玖'
    negative = amount < 0
    amount = abs(round(amount, 2))

    int_part = int(amount)
    frac = round((amount - int_part) * 100)
    jiao = frac // 10
    fen = frac % 10

    # --- integer part ---
    if int_part == 0:
        cn_int = '零'
    else:
        yi = int_part // 100_000_000
        wan = (int_part % 100_000_000) // 10_000
        ge = int_part % 10_000

        parts: list = []
        if yi:
            parts.append(_four_digit_chinese(yi) + '亿')
        if wan:
            need_zero = bool(yi) and wan < 1000
            if need_zero:
                parts.append('零')
            parts.append(_four_digit_chinese(wan) + '万')
        if ge:
            need_zero = bool(yi or wan) and ge < 1000
            if need_zero:
                parts.append('零')
            parts.append(_four_digit_chinese(ge))

        cn_int = ''.join(parts)

    result = cn_int + '元'

    if jiao == 0 and fen == 0:
        result += '整'
    elif jiao == 0:
        result += '零' + CN_NUM[fen] + '分'
    elif fen == 0:
        result += CN_NUM[jiao] + '角整'
    else:
        result += CN_NUM[jiao] + '角' + CN_NUM[fen] + '分'

    return ('负' if negative else '') + result


def _format_amount_str(amount) -> str:
    """Format *amount* as a two-decimal-place string, or '' if None."""
    if amount is None:
        return ''
    try:
        return f'{float(amount):.2f}'
    except (ValueError, TypeError):
        return str(amount)


# ---------------------------------------------------------------------------
# Data extraction – Excel quotation
# ---------------------------------------------------------------------------

def _extract_product_data(data: dict, all_rows: list, header_row_idx: int,
                          sheet_idx: int) -> None:
    """Parse the product table that starts at *header_row_idx* and enrich
    *data* with contract placeholders.

    Populates:
      b1bh / b1spxh / b1spmx / b1sl / b1dw / b1dj / b1wsje / b1hsje / b1zj
      b2bh / b2spxh / b2spmx / b2sl / b2dw / b2dj / b2wsje / b2hsje / b2zj
      合同总额大写/小写, 商品总额大写/小写, 商品未税大写/小写, 商品税额大写/小写
      服务总额大写/小写, 服务未税大写/小写, 服务税额大写/小写
    """
    header_row = [str(c).strip() if c is not None else ''
                  for c in all_rows[header_row_idx]]

    def find_col(*names: str) -> int:
        for name in names:
            try:
                return header_row.index(name)
            except ValueError:
                continue
        return -1

    col_bh = find_col('编号')
    col_spzl = find_col('商品种类')
    col_spmc = find_col('商品名称')
    col_mssm = find_col('产品描述', '商品描述', '规格描述')
    col_sl = find_col('数量')
    col_dw = find_col('单位')
    col_dj_hs = find_col('单价（含税）', '单价(含税)', '含税单价', '单价（含税 ）', '单价(含税 )')
    col_je_ws = find_col('金额(未税)', '金额（未税）', '未税金额')
    col_se = find_col('税额')
    col_je_hs = find_col('金额(含税）', '金额（含税）', '含税金额', '金额(含税)')
    col_sl_rate = find_col('税率')

    hardware_products: list = []
    service_products: list = []

    for row in all_rows[header_row_idx + 1:]:
        str_row = [str(c).strip() if c is not None else '' for c in row]

        if not any(str_row):
            continue

        # Stop at the summary row
        first_cell = str_row[0] if str_row else ''
        if first_cell in ('未税合计', '合计', '总计'):
            break

        # A product row has a numeric 编号 in the first column
        bh = str_row[col_bh] if col_bh >= 0 and col_bh < len(str_row) else ''
        try:
            int(bh)
        except (ValueError, TypeError):
            continue

        def gs(idx: int) -> str:
            return str_row[idx] if 0 <= idx < len(str_row) else ''

        def gn(idx: int):
            if idx < 0 or idx >= len(row):
                return None
            v = row[idx]
            if v is None:
                return None
            try:
                return float(v)
            except (ValueError, TypeError):
                return None

        spzl = gs(col_spzl)
        je_hs = gn(col_je_hs)
        je_ws = gn(col_je_ws)
        se = gn(col_se)
        dj_hs = gn(col_dj_hs)
        sl_str = gs(col_sl)
        sl_rate_str = gs(col_sl_rate).replace('%', '')

        # Parse tax rate
        rate: float | None = None
        if sl_rate_str:
            try:
                rate = float(sl_rate_str)
                if rate > 1:
                    rate /= 100
            except ValueError:
                pass

        # Parse quantity
        try:
            sl_num: float | None = float(sl_str) if sl_str else None
        except ValueError:
            sl_num = None

        # Recalculate amounts when formula results (data_only) are absent
        if je_hs is None and dj_hs is not None and sl_num is not None:
            je_hs = dj_hs * sl_num
            if rate is not None and rate > 0:
                je_ws = je_hs / (1 + rate)
                se = je_hs - je_ws
            else:
                je_ws = je_hs
                se = 0.0
        elif je_ws is None and je_hs is not None and rate is not None and rate > 0:
            je_ws = je_hs / (1 + rate)
            se = je_hs - je_ws

        product = {
            'bh': bh,
            'spzl': spzl,
            'spmc': gs(col_spmc),
            'mssm': gs(col_mssm),
            'sl': sl_str,
            'dw': gs(col_dw),
            'dj_hs': dj_hs,
            'je_ws': je_ws,
            'se': se,
            'je_hs': je_hs,
            'rate': rate,
        }

        # 严格按税率分类：13% → 表格1(b1*)，6% → 表格2(b2*)
        is_13pct = (rate is not None and abs(rate - 0.13) < 0.001)
        is_service = (rate is not None and abs(rate - 0.06) < 0.001)
        if is_13pct:
            hardware_products.append(product)
        elif is_service:
            service_products.append(product)
        # 其他税率的商品暂不归入表格

    # --- Compute totals ---
    def safe_sum(rows: list, key: str) -> float:
        total = 0.0
        for r in rows:
            v = r.get(key)
            if v is not None:
                try:
                    total += float(v)
                except (ValueError, TypeError):
                    pass
        return total

    hw_hs = safe_sum(hardware_products, 'je_hs')
    hw_ws = safe_sum(hardware_products, 'je_ws')
    hw_se = safe_sum(hardware_products, 'se')
    svc_hs = safe_sum(service_products, 'je_hs')
    svc_ws = safe_sum(service_products, 'je_ws')
    svc_se = safe_sum(service_products, 'se')
    total_hs = hw_hs + svc_hs

    # --- Map amounts to contract placeholders ---
    for key, val in [
        ('合同总额大写', _amount_to_chinese(total_hs)),
        ('合同总额小写', _format_amount_str(total_hs)),
        ('商品总额大写', _amount_to_chinese(hw_hs)),
        ('商品总额小写', _format_amount_str(hw_hs)),
        ('商品未税大写', _amount_to_chinese(hw_ws)),
        ('商品未税小写', _format_amount_str(hw_ws)),
        ('商品税额大写', _amount_to_chinese(hw_se)),
        ('商品税额小写', _format_amount_str(hw_se)),
        ('服务总额大写', _amount_to_chinese(svc_hs)),
        ('服务总额小写', _format_amount_str(svc_hs)),
        ('服务未税大写', _amount_to_chinese(svc_ws)),
        ('服务未税小写', _format_amount_str(svc_ws)),
        ('服务税额大写', _amount_to_chinese(svc_se)),
        ('服务税额小写', _format_amount_str(svc_se)),
        ('b1zj', _format_amount_str(hw_hs)),
        ('b2zj', _format_amount_str(svc_hs)),
    ]:
        data.setdefault(key, val)

    # First hardware product → b1* fields (bh 从 1 自动编号)
    if hardware_products:
        h = hardware_products[0]
        for key, val in [
            ('b1bh', '1'),
            ('b1spxh', h['spmc']),
            ('b1spmx', h['mssm']),
            ('b1spms', h['mssm']),  # 商品描述
            ('b1sl', h['sl']),
            ('b1dw', h['dw']),
            ('b1dj', _format_amount_str(h['dj_hs'])),
            ('b1wsje', _format_amount_str(h['je_ws'])),
            ('b1hsje', _format_amount_str(h['je_hs'])),
        ]:
            data.setdefault(key, val)

    # First service product → b2* fields
    if service_products:
        s = service_products[0]
        for key, val in [
            ('b2bh', s['bh']),
            ('b2spxh', s['spmc']),
            ('b2spmx', s['mssm']),
            ('b2sl', s['sl']),
            ('b2dw', s['dw']),
            ('b2dj', _format_amount_str(s['dj_hs'])),
            ('b2wsje', _format_amount_str(s['je_ws'])),
            ('b2hsje', _format_amount_str(s['je_hs'])),
        ]:
            data.setdefault(key, val)

    # Store raw product lists for multi-row table expansion in fill_template
    data[f'_sheet_{sheet_idx}_hardware_products'] = hardware_products
    data[f'_sheet_{sheet_idx}_service_products'] = service_products


def _apply_field_mappings(data: dict) -> None:
    """Apply standard field-name aliases from the quotation to contract slots."""
    # 发票类型 (e.g. "增值税-专用发票") → fblx (文字内容)
    fblx_raw = data.get('发票类型', '')
    if fblx_raw and 'fblx' not in data:
        data['fblx'] = '增值税专用发票' if '专用' in fblx_raw else '普通发票'

    # 报价单号码 → 项目名称 (fallback when project name is absent)
    if not data.get('项目名称') and data.get('报价单号码'):
        data.setdefault('项目名称', data['报价单号码'])


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


def _extract_billing_info(data: dict, sheet) -> None:
    """从'开票信息'Sheet按客户名称匹配，填充账号/税务登记号/开票地址电话。"""
    customer = data.get('客户名称', '').strip()
    if not customer:
        return
    rows = list(sheet.iter_rows(values_only=True))
    for row in rows[1:]:  # 跳过标题行
        if not row or row[0] is None:
            continue
        if str(row[0]).strip() == customer:
            if len(row) > 1 and row[1] is not None:
                data.setdefault('客户账号', str(row[1]).strip())
            if len(row) > 2 and row[2] is not None:
                data.setdefault('税务登记号', str(row[2]).strip())
            if len(row) > 3 and row[3] is not None:
                data.setdefault('开票地址电话', str(row[3]).strip())
            break


def _extract_payment_info(data: dict, sheet) -> None:
    """从'支付方式'Sheet生成付款信息文本，填充{{付款信息}}占位符。

    每行格式：付款方式文本 + 支付比例 + 按比例计算的大写/小写金额。
    """
    total_str = data.get('合同总额小写', '0')
    try:
        total = float(total_str)
    except (ValueError, TypeError):
        total = 0.0

    rows = list(sheet.iter_rows(values_only=True))
    items: list = []
    for row in rows[1:]:  # 跳过标题行
        if not row or row[0] is None:
            continue
        method = str(row[0]).strip()
        if not method:
            continue
        ratio_raw = row[1] if len(row) > 1 else None
        ratio: float | None = None
        if ratio_raw is not None:
            try:
                ratio = float(str(ratio_raw).replace('%', ''))
                if ratio > 1:
                    ratio /= 100
            except (ValueError, TypeError):
                pass
        if ratio is not None and total > 0:
            amount = round(total * ratio, 2)
            cn = _amount_to_chinese(amount)
            amt_str = _format_amount_str(amount)
            pct_str = f'{ratio * 100:.0f}%'
            item = (f'{method}{pct_str}，'
                    f'人民币大写[{cn}]，小写[{amt_str}]元')
        else:
            item = f'{method}，人民币大写[          ]，小写[          ]元'
        items.append(item)

    if items:
        # 合同模板已有（1）（2）条款，付款信息从（3）开始编号
        numbered = [f'（{i + 3}）{item}' for i, item in enumerate(items)]
        info = '\n'.join(numbered)
        data.setdefault('付款信息', info)


def extract_data_from_quotation_xlsx(xlsx_path: str) -> dict:
    """Return a dict of field→value pairs extracted from *xlsx_path*.

    Extraction strategy (applied per worksheet, first occurrence wins):

    1. **Multi-field rows** – cells that end with full-width '：' are labels;
       the cell immediately following is the value.  This handles the actual
       quotation format where several field-value pairs share one row.
    2. **Single-cell full-width pattern** – a lone cell whose text matches
       ``key：value`` (full-width colon only).
    3. **Two-column key-value** – when no label cells are present and there are
       exactly two non-empty cells, the first is the key and the second the
       value (backward-compatible with simple two-column quotations).
    4. **Product table** – rows after a header row containing '编号' and
       '商品名称' are parsed as product items; amounts are calculated and
       mapped to the contract's ``b1*`` / ``b2*`` / amount placeholders.
    5. **Multi-column table rows** (3+ non-empty cells, no labels) are stored
       as ``_sheet_N_table_headers`` / ``_sheet_N_table_rows`` private keys.
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    data: dict = {}

    for sheet_idx, sheet in enumerate(wb.worksheets):
        sheet_name = sheet.title

        # 开票信息 sheet → 按客户名称查询账号/税务登记号/开票地址电话
        if sheet_name == '开票信息':
            _extract_billing_info(data, sheet)
            continue

        # 支付方式 sheet → 生成付款信息文本
        if sheet_name == '支付方式':
            _extract_payment_info(data, sheet)
            continue

        all_rows = list(sheet.iter_rows(values_only=True))
        table_rows: list = []
        product_header_row_idx = -1

        for row_idx, row in enumerate(all_rows):
            # Once product table header is found, stop header-area scanning
            if product_header_row_idx >= 0 and row_idx > product_header_row_idx:
                break

            str_cells = [str(c).strip() if c is not None else '' for c in row]

            if not any(str_cells):
                continue

            # Detect product table header row (contains '编号' and a product column)
            if '编号' in str_cells and any(h in str_cells for h in ('商品名称', '商品种类')):
                product_header_row_idx = row_idx
                continue  # header itself is not a data row

            # Strategy 1: scan row for full-width label cells ending with '：'
            kv = _extract_kv_from_row(str_cells)
            if kv:
                for k, v in kv.items():
                    data.setdefault(k, v)
                continue

            # Strategy 2 & 3: no label cells found
            first = str_cells[0]
            non_empty = [c for c in str_cells if c]

            if len(non_empty) == 0:
                continue
            elif len(non_empty) == 1 and first:
                # Single cell: check for full-width "key：value" inline
                m = re.match(r'^([^：\n]{1,40})：\s*(.+)$', first)
                if m:
                    data.setdefault(m.group(1).strip(), m.group(2).strip())
            elif len(non_empty) == 2 and first:
                # Two-column key-value (key may or may not carry a colon)
                second = str_cells[1] if len(str_cells) > 1 else ''
                key = first.rstrip('： ').strip()
                if key:
                    data.setdefault(key, second)
            elif len(non_empty) >= 3:
                # Multi-column row without labels → store as private table data
                table_rows.append(non_empty)

        # Direct cell reference extraction (first sheet only)
        # 项目名称←A1；客户名称←C5；联系人←F5；送货地址←C6；电话号码←F6
        if sheet_idx == 0:
            a1 = all_rows[0][0] if all_rows and len(all_rows[0]) >= 1 else None
            if a1 is not None and str(a1).strip():
                data['项目名称'] = str(a1).strip()

            def _cell(r: int, c: int):
                """0-based row/col safe reader."""
                return (
                    all_rows[r][c]
                    if len(all_rows) > r and len(all_rows[r]) > c
                    else None
                )

            for _field, _r, _c in [
                ('客户名称', 4, 2),   # C5
                ('联系人',   4, 5),   # F5
                ('送货地址', 5, 2),   # C6
                ('电话号码', 5, 5),   # F6
            ]:
                _v = _cell(_r, _c)
                if _v is not None and str(_v).strip():
                    data[_field] = str(_v).strip()

        # Parse product table rows and compute contract amount placeholders
        if product_header_row_idx >= 0:
            _extract_product_data(data, all_rows, product_header_row_idx, sheet_idx)

        if table_rows:
            data.setdefault(f'_sheet_{sheet_idx}_table_headers', table_rows[0])
            data.setdefault(f'_sheet_{sheet_idx}_table_rows', table_rows[1:])

    # Apply cross-field mappings (e.g. 发票类型 → fblx)
    _apply_field_mappings(data)

    return data


# ---------------------------------------------------------------------------
# Template filling
# ---------------------------------------------------------------------------

def _expand_product_rows(table, hardware_products: list,
                         service_products: list) -> None:
    """Insert additional table rows when there are multiple products of the
    same type, cloning the template data row (the one with b1*/b2* placeholders)
    and filling each copy with the corresponding product data.

    This function modifies *table* in-place before normal placeholder
    replacement runs, so that the outer replacement loop will find no
    remaining b1*/b2* placeholders in the expanded rows.
    """
    # Maps prefix → product list
    groups = [('b1', hardware_products), ('b2', service_products)]

    for prefix, products in groups:
        if not products:
            continue

        sentinel = f'{{{{{prefix}bh}}}}'  # e.g. '{{b1bh}}'

        # Find the template data row
        template_row_idx = -1
        for i, row in enumerate(table.rows):
            if any(sentinel in cell.text for cell in row.cells):
                template_row_idx = i
                break

        if template_row_idx < 0:
            continue  # this table doesn't carry these placeholders

        template_row = table.rows[template_row_idx]

        # Insert (N-1) clones immediately after the template row
        if len(products) > 1:
            for _ in range(len(products) - 1):
                new_tr = deepcopy(template_row._tr)
                template_row._tr.addnext(new_tr)

        # Build per-product replacement dicts and fill each row
        for prod_idx, prod in enumerate(products):
            actual_row = table.rows[template_row_idx + prod_idx]
            prod_rep = {
                f'{prefix}bh': str(prod_idx + 1),  # 从 1 开始自动编号
                f'{prefix}spxh': str(prod.get('spmc', '')),
                f'{prefix}spmx': str(prod.get('mssm', '')),
                f'{prefix}spms': str(prod.get('mssm', '')),  # 商品描述
                f'{prefix}sl': str(prod.get('sl', '')),
                f'{prefix}dw': str(prod.get('dw', '')),
                f'{prefix}dj': _format_amount_str(prod.get('dj_hs')),
                f'{prefix}wsje': _format_amount_str(prod.get('je_ws')),
                f'{prefix}hsje': _format_amount_str(prod.get('je_hs')),
            }
            for cell in actual_row.cells:
                for para in cell.paragraphs:
                    _replace_in_paragraph(para, prod_rep)


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
    if '\n' not in new_text:
        # Simple case: single-line replacement
        if para.runs:
            para.runs[0].text = new_text
            for run in para.runs[1:]:
                run.text = ''
        else:
            para.add_run(new_text)
    else:
        # Multi-line: replace \n with Word <w:br/> line breaks
        lines = new_text.split('\n')
        if para.runs:
            first_run = para.runs[0]
            first_run.text = lines[0]
            for run in para.runs[1:]:
                run.text = ''
            r_elem = first_run._r
        else:
            first_run = para.add_run(lines[0])
            r_elem = first_run._r
        for line in lines[1:]:
            br = OxmlElement('w:br')
            r_elem.append(br)
            if line:
                t = OxmlElement('w:t')
                t.text = line
                if line.startswith(' ') or line.endswith(' '):
                    t.set(qn('xml:space'), 'preserve')
                r_elem.append(t)

    # 付款信息不显示下划线
    if '{{付款信息}}' in full_text and para.runs:
        para.runs[0].font.underline = False


def fill_template(template_path: str, data: dict, output_path: str) -> None:
    """Fill *template_path* with *data* and save the result to *output_path*.

    Placeholders in the template must use the ``{{字段名}}`` syntax.
    Only non-private keys (those not starting with ``_``) are used.

    If the extracted data contains multiple hardware or service products
    (stored under ``_sheet_N_hardware_products`` / ``_sheet_N_service_products``),
    the corresponding table rows in the template are duplicated so that every
    product appears on its own row.
    """
    doc = Document(template_path)
    replacements = {k: v for k, v in data.items() if not k.startswith('_')}

    # Collect product lists from private data (any sheet)
    hardware_products: list = []
    service_products: list = []
    for key, val in data.items():
        if key.endswith('_hardware_products') and not hardware_products:
            hardware_products = val
        if key.endswith('_service_products') and not service_products:
            service_products = val

    # Body paragraphs
    for para in doc.paragraphs:
        _replace_in_paragraph(para, replacements)

    # Table cells – expand product rows first, then apply normal replacements
    for table in doc.tables:
        _expand_product_rows(table, hardware_products, service_products)
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
