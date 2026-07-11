# 公文 Word 文档生成器

将符合公文 Markdown 写作规范的 `.md` 文件，自动转换为符合 GB/T 9704 标准的 Word 公文文档。

## 快速开始

```bash
# 1. 安装依赖（只需一次）
pip install -r requirements.txt

# 2. 转换 Markdown → Word
python3 scripts/md2json.py 你的公文.md /tmp/content.json
python3 scripts/generate.py /tmp/content.json 输出.docx

# 无方正小标宋？用开源字体预设（思源系列，免商业字体）：
python3 scripts/generate.py /tmp/content.json 输出.docx --template official_document_compat
```

## 字体准备（重要）

模板默认用**方正小标宋简体**（商业字体，多数电脑无）。**不装任何字体也能生成**——`generate.py` 会自动检测系统字体并降级到已装的等效字体（macOS/Windows 命名差异已处理），并打印「🔍 字体检测」报告。

- 想严格国标 → 参考 `references/字体安装说明.md` 安装字体
- 想完全免商业字体 → 用 `--template official_document_compat`（思源系列，SIL OFL）

## 目录结构

```
official-document-generator/
├── README.md                    ← 本文件
├── requirements.txt             ← Python 依赖
├── references/
│   ├── 公文Markdown写作规范.md   ← ★ 写作规范（给作者看）
│   ├── official_document.yaml   ← 国标模板（需方正小标宋等）
│   ├── official_document_compat.yaml ← 开源字体预设（思源）
│   └── 字体安装说明.md           ← 字体获取与安装指引
├── scripts/
│   ├── md2json.py               ← Markdown → JSON 转换器
│   └── generate.py              ← JSON → Word 生成器（含字体自适配）
├── builder/                     ← 文档构建引擎（含 fonts.py 字体自适配）
└── models/                      ← 数据模型
```

## 写作规范

写公文前请先阅读 `references/公文Markdown写作规范.md`，里面包含：
- YAML 元数据（发文机关、发文字号、公文类型）
- 层级规则（一、→（一）→ 1. → （1））
- 附件、落款、日期格式
- 申请表（表格）写法
- 四种公文类型差异（通知、请示、报告、函）

## 支持的公文类型

| 类型 | 标题后缀 | 结尾用语 |
|------|---------|---------|
| 通知 | …的**通知** | 特此通知。 |
| 请示 | …的**请示** | 妥否，请批示。 |
| 报告 | …的**报告** | 特此报告。 |
| 函 | …的**函** | 特此函达。 |

## 功能特性

- GB/T 9704 标准红头文件格式
- 方正小标宋标题、黑体/楷体/仿宋分层
- 正文固定28.95pt行距
- 自动引号转换
- 字体自适配：缺字体自动降级，分享给他人无需预装字体
- 生成后自动验证（必需元素、附件编号、标题层级）
