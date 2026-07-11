"""Custom heading numbering with Chinese counting, decimal, and chapter formats."""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def create_heading_numbering(doc, cfg: dict) -> int:
    """Create a custom multilevel numbering definition linked to heading styles.

    Returns the numId to reference in heading paragraphs.
    """
    num_cfg = cfg.get('heading_numbering', {})
    if not num_cfg:
        return None

    numbering_part = doc.part.numbering_part
    numbering_el = numbering_part._element

    # Find max abstractNumId
    existing = numbering_el.findall(qn('w:abstractNum'))
    max_id = max(int(el.get(qn('w:abstractNumId'), '0')) for el in existing) if existing else 0
    new_abs_id = str(max_id + 1)

    # Find max numId
    nums = numbering_el.findall(qn('w:num'))
    max_num_id = max(int(el.get(qn('w:numId'), '0')) for el in nums) if nums else 0
    new_num_id = str(max_num_id + 1)

    # Create abstractNum
    abs_num = OxmlElement('w:abstractNum')
    abs_num.set(qn('w:abstractNumId'), new_abs_id)

    multi = OxmlElement('w:multiLevelType')
    multi.set(qn('w:val'), 'multilevel')
    abs_num.append(multi)

    # Create levels for H1, H2, H3
    for level, style_name in [(0, 'Heading1'), (1, 'Heading2'), (2, 'Heading3')]:
        level_key = f'h{level + 1}'
        level_cfg = num_cfg.get(level_key, {})
        fmt = level_cfg.get('format', 'decimal')
        text = level_cfg.get('text', '%1.')

        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(level))

        # start
        start = OxmlElement('w:start')
        start.set(qn('w:val'), '1')
        lvl.append(start)

        # numFmt
        numFmt = OxmlElement('w:numFmt')
        numFmt.set(qn('w:val'), fmt)
        lvl.append(numFmt)

        # pStyle - link to heading style
        pStyle = OxmlElement('w:pStyle')
        pStyle.set(qn('w:val'), style_name)
        lvl.append(pStyle)

        # lvlText
        lvlText = OxmlElement('w:lvlText')
        lvlText.set(qn('w:val'), text)
        lvl.append(lvlText)

        # lvlJc
        lvlJc = OxmlElement('w:lvlJc')
        lvlJc.set(qn('w:val'), 'left')
        lvl.append(lvlJc)

        # paragraph properties
        pPr = OxmlElement('w:pPr')
        ind = OxmlElement('w:ind')
        left = str(level * 420)
        hang = '420' if level > 0 else '0'
        ind.set(qn('w:left'), left)
        ind.set(qn('w:hanging'), hang)
        pPr.append(ind)
        lvl.append(pPr)

        # run properties
        rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:hint'), 'eastAsia')
        rPr.append(rFonts)
        lvl.append(rPr)

        abs_num.append(lvl)

    # Insert abstractNum before first w:num
    first_num = numbering_el.find(qn('w:num'))
    if first_num is not None:
        numbering_el.insert(list(numbering_el).index(first_num), abs_num)
    else:
        numbering_el.append(abs_num)

    # Create num
    num = OxmlElement('w:num')
    num.set(qn('w:numId'), new_num_id)
    abs_ref = OxmlElement('w:abstractNumId')
    abs_ref.set(qn('w:val'), new_abs_id)
    num.append(abs_ref)
    numbering_el.append(num)

    return int(new_num_id)
