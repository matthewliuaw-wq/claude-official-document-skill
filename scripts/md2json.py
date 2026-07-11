#!/usr/bin/env python3
"""公文 Markdown → JSON 内容文件转换器。

将符合公文 Markdown 写作规范的 .md 文件，转换为 generate.py 所需的 JSON 内容文件。

映射规则见 skill.md 中的「步骤 2 详细：Markdown → ContentBlock 映射规则」。

用法:
    python3 scripts/md2json.py <input.md> <output.json>
    python3 scripts/md2json.py <input.md>  # 输出到同名 .json

生成后调用:
    python3 scripts/generate.py <output.json> <result.docx>
"""

import re
import json
import sys
from pathlib import Path


def parse_front_matter(text: str) -> tuple[dict, str]:
    """Split YAML front matter from body. Returns (yaml_dict, body_text)."""
    parts = text.split('---', 2)
    if len(parts) < 3:
        return {}, text
    yaml_text = parts[1].strip()
    body = parts[2]
    # Simple YAML parsing
    meta = {}
    for line in yaml_text.split('\n'):
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        m = re.match(r'^(\w+):\s*(.+)$', line)
        if m:
            meta[m.group(1)] = m.group(2).strip().strip('"').strip("'")
    return meta, body


def build_cover(yaml_meta: dict) -> dict:
    """Build cover block from YAML front matter for official documents."""
    org = yaml_meta.get('org', '××××××')
    doc_no = yaml_meta.get('doc_no', '×〔2026〕×号')
    return {
        'title': f'{org}文件',
        'info_lines': [doc_no],
        'show_cover': True,
    }


def detect_doc_type(title: str, yaml_meta: dict) -> str:
    """Detect document type from title or YAML."""
    if yaml_meta.get('type'):
        return yaml_meta['type']
    if '通知' in title:
        return '通知'
    if '请示' in title:
        return '请示'
    if '报告' in title:
        return '报告'
    if '函' in title:
        return '函'
    return '通知'


def convert(md_path: str, json_path: str | None = None) -> str:
    """Convert an official document Markdown file to JSON content file."""
    md = Path(md_path)
    if json_path is None:
        json_path = str(md.with_suffix('.json'))

    text = md.read_text(encoding='utf-8')
    yaml_meta, body = parse_front_matter(text)

    blocks: list[dict] = []
    lines = body.split('\n')
    i = 0
    in_code = False
    code_lang = ''
    code_buf: list[str] = []
    found_title = False
    doc_title = yaml_meta.get('title', '')
    doc_type = ''

    while i < len(lines):
        line = lines[i]

        # --- Code fences ---
        if line.strip().startswith('```'):
            if in_code:
                blocks.append({
                    'type': 'code_block',
                    'code_block': {'language': code_lang or None, 'code': '\n'.join(code_buf)},
                })
                in_code = False
                code_buf = []
                code_lang = ''
            else:
                in_code = True
                code_lang = line.strip()[3:].strip()
            i += 1
            continue

        if in_code:
            code_buf.append(line)
            i += 1
            continue

        # --- Tables ---
        if '|' in line and line.strip().startswith('|'):
            table_lines: list[str] = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i])
                i += 1
            rows_raw = [l for l in table_lines
                        if not re.match(r'^\|[\s\-:|]+\|$', l.strip())]
            if len(rows_raw) >= 1:
                headers = [c.strip() for c in rows_raw[0].strip('|').split('|')]
                data_rows = [[c.strip() for c in r.strip('|').split('|')]
                             for r in rows_raw[1:]]
                tbl: dict = {'headers': headers, 'rows': data_rows}
                # Auto column widths for 2-column tables (form-like)
                if len(headers) == 2:
                    tbl['col_widths'] = [2800, 6100]
                    # Row heights based on content length
                    row_heights = []
                    for row in data_rows:
                        content_len = max(len(c) for c in row)
                        if content_len > 20:
                            row_heights.append(1600)
                        elif content_len > 5:
                            row_heights.append(800)
                        else:
                            row_heights.append(500)
                    tbl['row_heights'] = row_heights
                blocks.append({'type': 'table', 'table': tbl})
            continue

        # --- Empty lines ---
        if not line.strip():
            i += 1
            continue

        # --- Horizontal rule → page break ---
        if line.strip() == '---':
            blocks.append({'type': 'page_break'})
            i += 1
            continue

        # --- Headings ---
        if line.startswith('### '):
            ht = line[4:].strip()
            blocks.append({'type': 'heading', 'level': 3, 'text': ht})
            i += 1
            continue

        if line.startswith('## '):
            ht = line[2:].strip()

            # Attachment label: ## 附件1：标题
            attach_m = re.match(r'^附件(\d+)[:：](.+)$', ht)
            if attach_m:
                attach_num = attach_m.group(1)
                attach_title = attach_m.group(2).strip()
                blocks.append({
                    'type': 'attachment_label',
                    'text': f'附件{attach_num}',
                })
                # Empty paragraph as separator
                blocks.append({'type': 'paragraph', 'text': ''})
                blocks.append({
                    'type': 'heading', 'level': 1, 'text': attach_title,
                })
                i += 1
                continue

            # Attachment section: ## 附件
            if ht == '附件' or ht.startswith('附件：') or ht.startswith('附件:'):
                # If "附件：XXX" (single attachment), use it as-is
                if ht.startswith('附件：') or ht.startswith('附件:'):
                    attach_name = re.split(r'附件[：:]', ht, 1)[1].strip()
                    blocks.append({
                        'type': 'paragraph',
                        'text': f'附件：{attach_name}',
                        'space_before': 60,
                    })
                else:
                    # "附件" section header — following lines are attachment list
                    # The section title itself becomes a paragraph
                    blocks.append({
                        'type': 'paragraph',
                        'text': '附件：',
                        'space_before': 60,
                    })
                    # Collect attachment list items
                    i += 1
                    while i < len(lines):
                        next_line = lines[i].strip()
                        if not next_line:
                            i += 1
                            continue
                        # Check for numbered items: 1.XXX 2.XXX
                        item_m = re.match(r'^(\d+)\.(.+)$', next_line)
                        if item_m:
                            blocks.append({
                                'type': 'paragraph',
                                'text': f'{item_m.group(1)}.{item_m.group(2).strip()}',
                            })
                            i += 1
                            continue
                        # Non-item line — stop collecting
                        break
                    continue
                i += 1
                continue

            # Regular heading level 2 (一、二、三 etc.)
            blocks.append({'type': 'heading', 'level': 2, 'text': ht})
            i += 1
            continue

        if line.startswith('# '):
            ht = line[2:].strip()
            blocks.append({'type': 'heading', 'level': 1, 'text': ht})
            if not found_title:
                found_title = True
                if not doc_title:
                    doc_title = ht
                doc_type = detect_doc_type(ht, yaml_meta)
            i += 1
            continue

        stripped = line.strip()

        # --- Signature/Date (right-aligned, end of main text) ---
        # Date pattern: XXXX年X月X日
        if re.match(r'^\d{4}年\d{1,2}月\d{1,2}日$', stripped):
            blocks.append({'type': 'paragraph', 'text': stripped, 'alignment': 'right'})
            i += 1
            continue

        # Signature: （机关名称）
        if re.match(r'^（.+）$', stripped):
            blocks.append({
                'type': 'paragraph', 'text': stripped,
                'alignment': 'right', 'space_before': 120,
            })
            i += 1
            continue

        # Main recipient: first paragraph ending with ：
        if (not found_title is False and
                stripped.endswith('：') and
                not any(b.get('type') == 'paragraph' and not b.get('text', '').startswith('（')
                        for b in blocks if b.get('alignment') != 'right')):
            # Check if this is one of the first paragraphs after title
            para_count = sum(1 for b in blocks if b['type'] == 'paragraph')
            if para_count == 0:
                blocks.append({
                    'type': 'paragraph', 'text': stripped, 'alignment': 'left',
                })
                i += 1
                continue

        # Regular paragraph
        blocks.append({'type': 'paragraph', 'text': stripped})
        i += 1

    # Assemble document
    doc = {
        'meta': {'title': doc_title},
        'cover': build_cover(yaml_meta),
        'content': blocks,
    }

    Path(json_path).write_text(json.dumps(doc, ensure_ascii=False, indent=2), encoding='utf-8')

    # Print summary
    types: dict[str, int] = {}
    for b in blocks:
        types[b['type']] = types.get(b['type'], 0) + 1
    print(f'✅ {len(blocks)} blocks → {json_path}')
    for t, c in sorted(types.items()):
        print(f'   {t}: {c}')
    print(f'   文件标题: {doc_title}')
    print(f'   公文类型: {doc_type}')

    return json_path


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(f'用法: python3 {sys.argv[0]} <input.md> [output.json]')
        sys.exit(1)
    out = sys.argv[2] if len(sys.argv) > 2 else None
    convert(sys.argv[1], out)
