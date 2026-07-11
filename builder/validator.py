"""Post-generation validator for official document quality checks."""

import re
from typing import List, Dict


def validate_document(doc, content, template_cfg: dict) -> Dict:
    """Run all validation checks on a generated document and return results.

    Automatically selects checks based on template validation config or
    defaults to official-document checks.

    Args:
        doc: The python-docx Document object
        content: The original DocumentContent (Pydantic model)
        template_cfg: The template configuration dict

    Returns:
        dict with 'passed', 'warnings', 'errors' keys.
    """
    errors: List[str] = []
    warnings: List[str] = []

    blocks = content.content
    validation_cfg = template_cfg.get('validation', {})
    mode = validation_cfg.get('mode', 'official_document')

    # --- Common checks ---

    # Straight quotes (when convert_quotes is enabled)
    # Skip for book_manuscript mode: code blocks legitimately use straight quotes
    if template_cfg.get('text_processing', {}).get('convert_quotes'):
        if mode != 'book_manuscript':
            _check_straight_quotes(doc, errors)

    # Heading hierarchy
    _check_heading_hierarchy(blocks, warnings)

    # --- Mode-specific checks ---

    if mode == 'official_document':
        _check_required_elements_official(blocks, errors)
        _check_attachment_numbering(blocks, warnings)

    elif mode == 'book_manuscript':
        _check_chapter_structure(blocks, errors)
        _check_figure_citations(blocks, warnings)
        _check_reference_format(blocks, warnings)

    return {
        'passed': len(errors) == 0,
        'errors': errors,
        'warnings': warnings,
    }


def _check_straight_quotes(doc, errors: List[str]):
    """Check for residual straight quotes in generated document text."""
    straight_double = 0
    straight_single = 0
    for para in doc.paragraphs:
        text = para.text
        straight_double += text.count('"')
        straight_single += text.count("'")

    # Also check tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text
                straight_double += text.count('"')
                straight_single += text.count("'")

    if straight_double > 0:
        errors.append(f'发现 {straight_double} 个英文直双引号（"）未转换为中文引号')
    if straight_single > 0:
        errors.append(f'发现 {straight_single} 个英文直单引号（\'）未转换为中文引号')


def _check_required_elements_official(blocks, errors: List[str]):
    """Check that required official document elements are present."""
    texts = []
    headings = []
    has_attachment_section = False
    has_date = False
    has_signature = False

    for block in blocks:
        if block.type == 'paragraph':
            t = block.text or ''
            texts.append(t)
            # Date pattern: XXXX年X月X日
            if re.search(r'\d{4}年\d{1,2}月\d{1,2}日', t):
                has_date = True
            # Signature pattern: （XXX）
            if re.match(r'^（.+）$', t.strip()):
                has_signature = True
        elif block.type == 'heading':
            headings.append(block.text or '')
            if block.text == '附件':
                has_attachment_section = True

    # Check title (heading level 1)
    has_title = any(block.type == 'heading' and block.level == 1 for block in blocks)
    if not has_title:
        errors.append('缺少文件标题（一级标题）')

    # Check recipient (first paragraph ending with ：)
    has_recipient = any(t.rstrip().endswith('：') for t in texts[:3])  # check first 3 paragraphs
    if not has_recipient:
        errors.append('缺少主送机关（以全角冒号结尾的段落）')

    # Check date
    if not has_date:
        errors.append('缺少落款日期（格式：XXXX年X月X日）')

    # Check signature
    if not has_signature:
        errors.append('缺少发文机关（格式：（机关名称））')


def _check_heading_hierarchy(blocks, warnings: List[str]):
    """Check for skipped heading levels (e.g. level 1 → level 3 without level 2)."""
    prev_level = 0
    for block in blocks:
        if block.type == 'heading':
            level = block.level or 1
            if prev_level > 0 and level > prev_level + 1:
                warnings.append(
                    f'标题层级跳跃：从 {prev_level} 级（{block.text[:20] if prev_text else "..."}）'
                    f'直接到 {level} 级（{block.text[:20]}）'
                )
            prev_level = level
            prev_text = block.text


def _check_attachment_numbering(blocks, warnings: List[str]):
    """Check that attachment labels are sequential (附件1, 附件2, ...)."""
    attachment_nums = []
    for block in blocks:
        if block.type == 'attachment_label':
            text = block.text or ''
            match = re.search(r'附件(\d+)', text)
            if match:
                attachment_nums.append(int(match.group(1)))

    if attachment_nums:
        for i, num in enumerate(attachment_nums):
            expected = i + 1
            if num != expected:
                warnings.append(
                    f'附件编号不连续：期望 附件{expected}，实际为 附件{num}'
                )


def _check_chapter_structure(blocks, errors: List[str]):
    """Check that each chapter has required elements: intro, objectives, summary, exercises."""
    chapters = []  # list of (chapter_text, block_index_range)
    chapter_start = None
    chapter_title = None

    for i, block in enumerate(blocks):
        if block.type == 'heading' and block.level == 1:
            if chapter_start is not None:
                chapters.append((chapter_title, chapter_start, i))
            chapter_start = i
            chapter_title = block.text or ''

    # Don't forget the last chapter
    if chapter_start is not None:
        chapters.append((chapter_title, chapter_start, len(blocks)))

    for title, start, end in chapters:
        # Skip front matter (内容提要, 前　言, 目　录, 参考文献)
        # Normalize: strip full-width/half-width spaces used for visual spacing
        stripped_title = title.replace('　', '').replace(' ', '')
        if any(kw in stripped_title for kw in ['内容提要', '前言', '目录', '参考文献']):
            continue

        chapter_blocks = blocks[start:end]
        chapter_text = ' '.join(
            (b.text or '') for b in chapter_blocks if b.type == 'paragraph'
        )
        headings_text = ' '.join(
            (b.text or '') for b in chapter_blocks if b.type == 'heading'
        )

        short_title = title[:15]
        if '本章学习目标' not in chapter_text:
            errors.append(f'"{short_title}" 缺少学习目标（本章学习目标：）')
        if '本章小结' not in headings_text:
            errors.append(f'"{short_title}" 缺少本章小结')
        if '习题' not in headings_text:
            errors.append(f'"{short_title}" 缺少习题')


def _check_figure_citations(blocks, warnings: List[str]):
    """Check that figures have text citations using 如图X-Y or 见图X-Y format."""
    figures = []
    for block in blocks:
        if block.type == 'figure' and block.figure:
            figures.append(block.figure.number)

    if not figures:
        return

    # Collect all paragraph text
    all_text = ' '.join(
        (b.text or '') for b in blocks if b.type == 'paragraph'
    )

    for fig_num in figures:
        # Check citation patterns: 如图X-Y or 见图X-Y
        patterns = [f'图{fig_num}']
        if not any(p in all_text for p in patterns):
            warnings.append(f'图{fig_num} 未在正文中引用（应使用"如图{fig_num}所示"或"见图{fig_num}"）')


def _check_reference_format(blocks, warnings: List[str]):
    """Check that references follow GB/T 7714 format with [N] numbering."""
    in_references = False
    ref_numbers = []

    for block in blocks:
        if block.type == 'heading' and '参考文献' in (block.text or ''):
            in_references = True
            continue
        if block.type == 'heading' and in_references:
            in_references = False
            continue

        if in_references and block.type == 'paragraph':
            text = (block.text or '').strip()
            match = re.match(r'^\[(\d+)\]', text)
            if match:
                ref_numbers.append(int(match.group(1)))
            elif text and not text.startswith('注') and text not in ('', ' '):
                warnings.append(f'参考文献条目格式不规范（应以 [N] 开头）：{text[:30]}...')

    # Check sequential numbering
    if ref_numbers:
        for i, num in enumerate(ref_numbers):
            expected = i + 1
            if num != expected:
                warnings.append(f'参考文献编号不连续：期望 [{expected}]，实际为 [{num}]')


def format_validation_report(result: Dict) -> str:
    """Format validation results into a human-readable string."""
    lines = []

    if result['passed'] and not result['warnings']:
        lines.append('✅ 文档检查通过，未发现问题。')
    elif result['passed']:
        lines.append('⚠️ 文档检查通过，但有警告：')
    else:
        lines.append('❌ 文档检查未通过：')

    for err in result['errors']:
        lines.append(f'  ❌ {err}')

    for warn in result['warnings']:
        lines.append(f'  ⚠️ {warn}')

    return '\n'.join(lines)
