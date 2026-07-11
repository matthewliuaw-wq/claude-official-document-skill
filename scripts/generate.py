#!/usr/bin/env python3
"""公文 Word 文档生成脚本 - 自包含版本。

本脚本及所有依赖（builder、models、模板）均在同一 skill 目录下，
无需安装 docx_mcp MCP 服务器。

用法:
    python3 scripts/generate.py <input.json> <output.docx>
    python3 scripts/generate.py <input.json> <output.docx> --template official_document_compat

依赖:
    pip install python-docx pydantic pyyaml fonttools
"""

import sys
import json
import argparse
import yaml
from pathlib import Path

# 设置路径 — skill 根目录（包含 builder/、models/、references/）
SKILL_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = SKILL_DIR / 'references'

sys.path.insert(0, str(SKILL_DIR))

from models.content import DocumentContent
from builder import build_document
from builder.fonts import detect_system_fonts, resolve_fonts, print_font_report
from builder.validator import validate_document, format_validation_report

DEFAULT_TEMPLATE = 'official_document'


def generate(input_json_path: str, output_docx_path: str, template_name: str = DEFAULT_TEMPLATE):
    """从 JSON 内容文件生成 Word 公文文档。"""

    # 1. 加载内容
    content_data = json.load(open(input_json_path, encoding='utf-8'))

    # 2. 加载模板
    template_path = TEMPLATE_DIR / f'{template_name}.yaml'
    if not template_path.exists():
        print(f"❌ 模板文件不存在: {template_path}")
        print(f"   可用模板: official_document / official_document_compat")
        sys.exit(1)

    template_cfg = yaml.safe_load(open(template_path, encoding='utf-8'))
    print(f"📖 模板: {template_name} (v{template_cfg.get('version', '?')})")

    # 3. 字体自适配：检测系统已装字体，缺失则按平台自动降级到等效字体
    installed = detect_system_fonts()
    template_cfg, font_report = resolve_fonts(template_cfg, installed)
    print_font_report(font_report)

    # 4. 构建 Word 文档
    content = DocumentContent(**content_data)
    doc = build_document(content, template_cfg)

    # 5. 保存
    out = Path(output_docx_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))

    # 6. 验证
    validation = validate_document(doc, content, template_cfg)
    report = format_validation_report(validation)
    print(f"\n{report}")

    # 7. 打印标题样式验证
    from docx.oxml.ns import qn
    for sn in ['Heading 1', 'Heading 2', 'Heading 3']:
        try:
            s = doc.styles[sn]
            rPr = s.element.find(qn('w:rPr'))
            sz = rPr.find(qn('w:sz')).get(qn('w:val')) if rPr is not None and rPr.find(qn('w:sz')) is not None else '?'
            pt = float(sz) / 2 if sz != '?' else '?'
            print(f"  {sn}: {pt}pt, bold={s.font.bold}")
        except KeyError:
            pass

    print(f"\n✅ 已保存: {out}")
    return str(out)


def main():
    parser = argparse.ArgumentParser(description='生成公文 Word 文档（自包含）')
    parser.add_argument('input_json', help='内容 JSON 文件路径（由 md2json.py 生成）')
    parser.add_argument('output_docx', help='输出 .docx 路径')
    parser.add_argument('--template', default=DEFAULT_TEMPLATE,
                        help='模板名（references/<name>.yaml），默认 official_document；'
                             '开源字体预设用 official_document_compat')
    args = parser.parse_args()
    generate(args.input_json, args.output_docx, args.template)


if __name__ == '__main__':
    main()
