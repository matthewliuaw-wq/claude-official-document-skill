"""Table rendering: three-line, full-border, and colored-header styles."""
from docx.shared import Pt, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def _set_cell_border(cell, **kwargs):
    """Set borders on a single cell. kwargs: top/right/bottom/left with (size, color)."""
    tc = cell._tc
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)

    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    for edge, val in kwargs.items():
        el = tcBorders.find(qn(f'w:{edge}'))
        if el is None:
            el = OxmlElement(f'w:{edge}')
            tcBorders.append(el)
        el.set(qn('w:val'), 'single' if val else 'nil')
        el.set(qn('w:sz'), str(val[0]) if val else '0')
        el.set(qn('w:color'), val[1] if val else '000000')
        el.set(qn('w:space'), '0')


def _set_cell_shading(cell, color: str):
    """Set background shading on a cell."""
    tc = cell._tc
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)

    shading = tcPr.find(qn('w:shd'))
    if shading is None:
        shading = OxmlElement('w:shd')
        tcPr.append(shading)
    shading.set(qn('w:fill'), color)
    shading.set(qn('w:val'), 'clear')


def _set_cell_text_color(cell, color: str):
    """Set text color in all paragraphs of a cell."""
    for para in cell.paragraphs:
        for run in para.runs:
            run.font.color.rgb = RGBColor.from_string(color)


def _process_text(text: str, cfg: dict) -> str:
    """Apply text processing (quote conversion, etc.) to table cell text."""
    from .document_builder import _maybe_convert_quotes
    return _maybe_convert_quotes(text, cfg)


def build_table(doc, table_data, cfg: dict):
    """Build a formatted table from TableContent data."""
    headers = table_data.headers
    rows = table_data.rows
    num_cols = len(headers)
    table_cfg = cfg.get('table', {})
    style = table_cfg.get('style', 'full_border')
    border_color = table_cfg.get('border_color', '000000')
    border_size = table_cfg.get('border_size', 4)
    header_shading = table_cfg.get('header_shading', 'E8EFF5')
    header_text_color = table_cfg.get('header_text_color', '000000')
    body_size = cfg.get('fonts', {}).get('body', {}).get('size', 10)

    # Add caption if provided
    if table_data.caption:
        p = doc.add_paragraph(table_data.caption)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.size = Pt(10)
            run.font.bold = True

    # Create table
    table = doc.add_table(rows=1 + len(rows), cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set column widths if provided (values in DXA/twips, convert to EMU)
    if table_data.col_widths:
        for i, w in enumerate(table_data.col_widths):
            for row in table.rows:
                row.cells[i].width = Emu(w * 635)

    # Header row
    header_row = table.rows[0]
    for i, text in enumerate(headers):
        text = _process_text(text, cfg)
        cell = header_row.cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        run = p.add_run(text)
        run.font.bold = True
        run.font.size = Pt(body_size)
    if style == 'three_line':
        _apply_three_line(table, border_color, header_shading)
    elif style == 'colored_header':
        _apply_colored_header(table, border_color, border_size, header_shading, header_text_color)
    else:  # full_border
        _apply_full_border(table, border_color, border_size, header_shading)

    # Data rows
    row_heights = table_data.row_heights
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        # Set minimum row height if specified
        if row_heights and r_idx < len(row_heights):
            tr = row._tr
            trPr = tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = OxmlElement('w:trPr')
                tr.insert(0, trPr)
            trHeight = OxmlElement('w:trHeight')
            trHeight.set(qn('w:val'), str(row_heights[r_idx]))
            trHeight.set(qn('w:hRule'), 'atLeast')  # minimum height
            trPr.append(trHeight)
        for c_idx, text in enumerate(row_data):
            if c_idx < num_cols:
                text = _process_text(text, cfg)
                row.cells[c_idx].text = ''
                p = row.cells[c_idx].paragraphs[0]
                # First column: center align (label column)
                # Other columns: left align (content columns)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.first_line_indent = Pt(0)
                run = p.add_run(text)
                run.font.size = Pt(body_size)


def _apply_full_border(table, color, size, header_shading):
    """All cells get uniform borders. Header row gets shading."""
    header_color = (size, color)
    for r_idx, row in enumerate(table.rows):
        for cell in row.cells:
            _set_cell_border(cell,
                             top=header_color, bottom=header_color,
                             left=header_color, right=header_color)
            if r_idx == 0:
                _set_cell_shading(cell, header_shading)


def _apply_three_line(table, color, header_shading):
    """Academic three-line table: thick top/bottom, thin header-bottom, no sides."""
    thick = (8, color)   # ~1.5pt
    thin = (4, color)    # ~0.75pt
    no_border = None

    for r_idx, row in enumerate(table.rows):
        for cell in row.cells:
            if r_idx == 0:
                # Header: thick top, thin bottom, no sides
                _set_cell_border(cell, top=thick, bottom=thin, left=no_border, right=no_border)
                _set_cell_shading(cell, header_shading)
            elif r_idx == len(table.rows) - 1:
                # Last row: thick bottom, no top/sides
                _set_cell_border(cell, bottom=thick, top=no_border, left=no_border, right=no_border)
            else:
                # Middle rows: no borders
                _set_cell_border(cell, top=no_border, bottom=no_border, left=no_border, right=no_border)


def _apply_colored_header(table, color, size, header_shading, header_text_color):
    """Full border + colored header with white text."""
    border = (size, color)
    for r_idx, row in enumerate(table.rows):
        for cell in row.cells:
            _set_cell_border(cell, top=border, bottom=border, left=border, right=border)
            if r_idx == 0:
                _set_cell_shading(cell, header_shading)
                _set_cell_text_color(cell, header_text_color)
