"""Font and style utilities for python-docx with East Asian font support."""
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING


def set_east_asian_font(element, latin_font: str, east_asia_font: str, size_pt: int = None,
                        unify_fonts: bool = False):
    """Set both Latin and East Asian fonts on a style or run element.

    python-docx has no high-level API for East Asian fonts, so we must
    manipulate the XML directly.

    Args:
        unify_fonts: If True, set ascii/hAnsi/eastAsia all to east_asia_font.
                     Needed for some CJK fonts on macOS Word (e.g. 方正小标宋简体).
    """
    # Get or create rPr (run properties)
    rPr = element.find(qn('w:rPr'))
    if rPr is None:
        rPr = element.makeelement(qn('w:rPr'), {})
        element.insert(0, rPr)  # rPr must come before w:t in XML

    # Get or create rFonts
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.insert(0, rFonts)

    # CRITICAL: Remove theme font attributes — they override direct font names in Word.
    # python-docx default heading styles include e.g. asciiTheme="majorHAnsi" which maps
    # to Arial/Calibri and takes precedence over our explicit font names.
    for theme_attr in ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme']:
        attr_qn = qn(f'w:{theme_attr}')
        if rFonts.get(attr_qn) is not None:
            del rFonts.attrib[attr_qn]

    if unify_fonts:
        rFonts.set(qn('w:ascii'), east_asia_font)
        rFonts.set(qn('w:hAnsi'), east_asia_font)
        rFonts.set(qn('w:eastAsia'), east_asia_font)
    else:
        rFonts.set(qn('w:ascii'), latin_font)
        rFonts.set(qn('w:hAnsi'), latin_font)
        rFonts.set(qn('w:eastAsia'), east_asia_font)

    if size_pt is not None:
        # Remove existing sz elements
        for sz in rPr.findall(qn('w:sz')):
            rPr.remove(sz)
        for sz in rPr.findall(qn('w:szCs')):
            rPr.remove(sz)
        # OOXML w:sz val must be integer (half-points), e.g. 10.5pt → val="21" NOT "21.0"
        sz_val = int(round(size_pt * 2))
        sz_el = rPr.makeelement(qn('w:sz'), {qn('w:val'): str(sz_val)})
        rPr.append(sz_el)


def apply_body_style(doc, cfg: dict):
    """Apply body text font and paragraph settings."""
    style = doc.styles['Normal']
    font_cfg = cfg.get('fonts', {}).get('body', {})
    para_cfg = cfg.get('paragraph', {})

    latin = font_cfg.get('latin', 'Times New Roman')
    east = font_cfg.get('east_asia', 'SimSun')
    size = font_cfg.get('size', 12)

    set_east_asian_font(style.element, latin, east, size)

    # Explicit black color for body text (avoid theme color inheritance issues)
    style.font.color.rgb = RGBColor(0, 0, 0)

    pf = style.paragraph_format
    ls = para_cfg.get('line_spacing', 1.5)
    if isinstance(ls, str) and ls.endswith('pt'):
        # Fixed line spacing, e.g. "28.95pt" for Chinese government documents
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(float(ls.replace('pt', '')))
    elif isinstance(ls, float):
        pf.line_spacing = ls
    elif isinstance(ls, (int, float)) and ls > 10:
        pf.line_spacing = Pt(ls)

    sa = para_cfg.get('space_after', 0)
    pf.space_after = Pt(sa)

    fi = para_cfg.get('first_line_indent', 0)
    pf.first_line_indent = Pt(fi) if fi else None

    align = para_cfg.get('alignment', 'justified')
    align_map = {
        'justified': WD_ALIGN_PARAGRAPH.JUSTIFY,
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'center': WD_ALIGN_PARAGRAPH.CENTER,
    }
    pf.alignment = align_map.get(align, WD_ALIGN_PARAGRAPH.JUSTIFY)


def _fix_linked_char_style(doc, char_style_id: str, latin: str, east_asia: str,
                           size: int, color_str: str):
    """Fix theme fonts in a linked character style (e.g. Heading2Char).

    These styles are marked customStyle="1" and may not appear in
    doc.styles dict. We locate them directly in the XML and apply the same
    font cleanup as the paragraph style.
    """
    from lxml import etree
    # Navigate to the styles XML element
    try:
        styles_element = doc.styles.element
    except Exception:
        return

    # Find the character style by styleId
    for style_el in styles_element.findall(qn('w:style')):
        if style_el.get(qn('w:styleId')) == char_style_id:
            set_east_asian_font(style_el, latin, east_asia, size, unify_fonts=True)
            # Also clear theme color
            if color_str:
                rPr = style_el.find(qn('w:rPr'))
                if rPr is not None:
                    for color_el in rPr.findall(qn('w:color')):
                        rPr.remove(color_el)
                    color_new = rPr.makeelement(qn('w:color'), {qn('w:val'): color_str})
                    rPr.append(color_new)
            # Remove bold from char style so it inherits from paragraph style
            rPr = style_el.find(qn('w:rPr'))
            if rPr is not None:
                for b in rPr.findall(qn('w:b')):
                    rPr.remove(b)
                for bcs in rPr.findall(qn('w:bCs')):
                    rPr.remove(bcs)
            return


def apply_heading_styles(doc, cfg: dict):
    """Apply heading styles with fonts, sizes, optional color and per-level alignment/spacing."""
    heading_cfg = cfg.get('fonts', {}).get('heading', {})
    latin = heading_cfg.get('latin', 'Times New Roman')
    east = heading_cfg.get('east_asia', 'SimHei')
    color_str = heading_cfg.get('color')  # optional, e.g. "FF0000"

    align_map = {
        'justified': WD_ALIGN_PARAGRAPH.JUSTIFY,
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right': WD_ALIGN_PARAGRAPH.RIGHT,
    }

    for level, size_key in [(1, 'h1_size'), (2, 'h2_size'), (3, 'h3_size')]:
        style_name = f'Heading {level}'
        if style_name in doc.styles:
            style = doc.styles[style_name]
            size = heading_cfg.get(size_key, 14)
            # Per-level east_asia font override (e.g. h1 uses 方正小标宋, h2/h3 use 黑体)
            level_east = heading_cfg.get(f'h{level}_east_asia', east)
            set_east_asian_font(style.element, latin, level_east, size, unify_fonts=True)

            # CRITICAL: Also fix the linked character style (e.g. Heading2Char).
            # python-docx links each heading paragraph style to a character style via
            # <w:link w:val="Heading2Char"/>. The linked Char style has its own rPr
            # with theme fonts (asciiTheme="majorHAnsi" → Arial/Calibri) that OVERRIDE
            # the paragraph style's direct font settings in Word.
            #
            # These Char styles are marked customStyle="1" and may not be accessible
            # via doc.styles[style_id]. So we find them directly in the XML styles part.
            _fix_linked_char_style(doc, char_style_id := f'Heading{level}Char',
                                   latin, level_east, size, color_str)
            # Per-level bold control (default True for H1, configurable for H2/H3)
            bold_val = heading_cfg.get(f'h{level}_bold')
            if bold_val is not None:
                style.font.bold = bold_val
            else:
                style.font.bold = True  # default bold for headings

            # Fix: python-docx adds empty <w:bCs/> when setting bold,
            # which defaults to true for complex script text.
            # Must explicitly set bCs to match b.
            rPr = style.element.find(qn('w:rPr'))
            if rPr is not None:
                actual_bold = bold_val if bold_val is not None else True
                # Remove existing bCs
                for bcs in rPr.findall(qn('w:bCs')):
                    rPr.remove(bcs)
                # Add explicit bCs matching b
                bcs_el = rPr.makeelement(qn('w:bCs'), {qn('w:val'): str(actual_bold).lower()})
                rPr.append(bcs_el)

            # Apply heading color if configured
            if color_str:
                style.font.color.rgb = RGBColor.from_string(color_str)

            # Apply per-level alignment if configured
            level_align = heading_cfg.get(f'h{level}_alignment')
            if level_align:
                style.paragraph_format.alignment = align_map.get(level_align)
                # Centered headings must have zero indent for true centering
                if level_align == 'center':
                    style.paragraph_format.left_indent = Pt(0)
                    style.paragraph_format.first_line_indent = Pt(0)

            # Per-level line spacing (optional, e.g. h1_line_spacing: 2.4)
            pf = style.paragraph_format
            ls_val = heading_cfg.get(f'h{level}_line_spacing')
            if ls_val is not None:
                if isinstance(ls_val, str) and ls_val.endswith('pt'):
                    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                    pf.line_spacing = Pt(float(ls_val.replace('pt', '')))
                elif isinstance(ls_val, (int, float)):
                    pf.line_spacing = ls_val

            # Per-level space before (optional, in pt)
            sb_val = heading_cfg.get(f'h{level}_space_before')
            if sb_val is not None:
                pf.space_before = Pt(sb_val)

            # Per-level space after (optional, in pt)
            sa_val = heading_cfg.get(f'h{level}_space_after')
            if sa_val is not None:
                pf.space_after = Pt(sa_val)


def configure_page(doc, cfg: dict):
    """Set page size and margins."""
    page_cfg = cfg.get('page', {})
    margin_cfg = page_cfg.get('margins', {})
    section = doc.sections[0]

    # Set page size (default to python-docx default if not specified)
    page_size = page_cfg.get('size', '').upper()
    if page_size == 'A4':
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)

    for key, attr in [('top', 'top_margin'), ('bottom', 'bottom_margin'),
                      ('left', 'left_margin'), ('right', 'right_margin')]:
        val = margin_cfg.get(key)
        if val:
            cm_val = float(val.replace('cm', ''))
            setattr(section, attr, Cm(cm_val))
