"""Cover page generation.

Reference layout (from 初稿撰写规范):
  [top]    副标题/系列名（14pt 黑体）
           空行×2
           书名（26pt 黑体 bold）
           空行×2
           作者信息（12pt 黑体）— 可多行
           空行×N（把出版社推到底部）
  [bottom] 出版社（12pt 黑体，字间距加宽）
           地点（12pt 黑体，字间距加宽）
"""
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from .styles import set_east_asian_font


def _apply_run_color(run, color_str):
    """Apply RGB color to a run if color_str is provided."""
    if color_str:
        run.font.color.rgb = RGBColor.from_string(color_str)


def _add_spacer(doc, count=1):
    """Add empty paragraphs as vertical spacers."""
    for _ in range(count):
        doc.add_paragraph()


def add_cover_page(doc, cover_data, cfg: dict):
    """Generate a cover page from CoverPageContent and template config.

    Layout follows publisher standard: subtitle on top, book title centered,
    author info in middle, publisher/location at bottom of page.
    """
    if not cover_data or not cover_data.show_cover:
        return

    cover_cfg = cfg.get('cover_page', {})
    title_cfg = cover_cfg.get('title_font', {})
    sub_cfg = cover_cfg.get('subtitle_font', {})
    info_cfg = cover_cfg.get('info_font', {})

    # --- Top section: subtitle / series name ---
    if cover_data.subtitle:
        _add_spacer(doc, 2)  # top margin
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Pt(0)
        run = p.add_run(cover_data.subtitle)
        run.font.bold = True
        run.font.size = Pt(sub_cfg.get('size', 14))
        set_east_asian_font(run._element, sub_cfg.get('east_asia', 'SimHei'),
                            sub_cfg.get('east_asia', 'SimHei'), unify_fonts=True)
        _apply_run_color(run, sub_cfg.get('color'))

    # --- Title ---
    _add_spacer(doc, 2)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    run = p.add_run(cover_data.title)
    run.font.bold = True
    run.font.size = Pt(title_cfg.get('size', 26))
    set_east_asian_font(run._element, title_cfg.get('east_asia', 'SimHei'),
                        title_cfg.get('east_asia', 'SimHei'), unify_fonts=True)
    _apply_run_color(run, title_cfg.get('color'))

    # --- Author info (middle section) ---
    _add_spacer(doc, 2)
    if cover_data.info_lines:
        # Detect document type by structure:
        # - Book manuscript: info_lines has >= 3 items (authors + '' + publisher + location)
        #   Convention: last 2 lines are publisher + location, with a '' separator before them
        # - Official document: info_lines has 1-2 items (发文字号 only), no publisher section
        all_lines = cover_data.info_lines

        # Check if there's a publisher section (book manuscript pattern)
        # Publisher section exists when: there's an empty line ('') acting as separator,
        # or when there are >= 4 lines (authors + separator + publisher + location)
        has_publisher = False
        if len(all_lines) >= 3:
            # Look for the empty-line separator before publisher+location
            # Pattern: [author1, author2, '', publisher, location]
            for idx, line in enumerate(all_lines):
                if line.strip() == '' and idx <= len(all_lines) - 2:
                    has_publisher = True
                    author_lines = all_lines[:idx]
                    publisher_lines = all_lines[idx + 1:]
                    break
            if not has_publisher:
                author_lines = all_lines
                publisher_lines = []
        else:
            author_lines = all_lines
            publisher_lines = []

        for line in author_lines:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            run = p.add_run(line)
            run.font.size = Pt(info_cfg.get('size', 12))
            set_east_asian_font(run._element, info_cfg.get('east_asia', 'SimHei'),
                                info_cfg.get('east_asia', 'SimHei'))
            _apply_run_color(run, info_cfg.get('color'))

        # --- Bottom section: publisher + location (book manuscript only) ---
        if has_publisher and publisher_lines:
            # Push to bottom of page using a single paragraph with space_before.
            # A4 usable height ≈ 24.6cm (29.7 - 2.54×2). Top content ≈ 7-8cm.
            # Publisher section ≈ 1cm. Need ~15.5cm gap ≈ 440pt.
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.space_before = Pt(380)

            for line in publisher_lines:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Pt(0)
                # Add letter spacing for publisher name (e.g. "某 某 出 版 社")
                spaced_line = _add_letter_spacing(line)
                run = p.add_run(spaced_line)
                run.font.size = Pt(info_cfg.get('size', 12))
                set_east_asian_font(run._element, info_cfg.get('east_asia', 'SimHei'),
                                    info_cfg.get('east_asia', 'SimHei'))
                _apply_run_color(run, info_cfg.get('color'))
    else:
        _add_spacer(doc, 28)


def _add_letter_spacing(text):
    """Add spaces between characters for publisher/location display.

    Reference: "某 某 出 版 社" / "北 京"
    If the text already contains spaces, return as-is.
    """
    if ' ' in text or '　' in text:
        return text
    # Add space between each character
    return ' '.join(text)


def add_separator_line(doc, cfg: dict):
    """Add a horizontal separator line (e.g. red line for government documents)."""
    sep_cfg = cfg.get('cover_page', {}).get('separator')
    if not sep_cfg:
        return

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Create bottom border on the paragraph
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')

    bottom = OxmlElement('w:bottom')
    color = sep_cfg.get('color', 'FF0000')
    width = sep_cfg.get('width', 2)
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(width * 4))  # sz is in 1/8 pt
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
