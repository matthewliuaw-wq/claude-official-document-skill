---
name: official-document-generator
description: 将符合公文Markdown写作规范的文档转换为Word公文。当用户要求生成公文、生成通知、生成请示、生成报告、生成函的Word文档，或将Markdown转为公文格式时使用此skill。触发场景：用户说"生成公文""转成Word公文""把这个通知生成公文""生成一份通知Word"等。
---

# 公文 Word 文档生成器

将符合公文 Markdown 写作规范的文档，自动转换为符合 GB/T 9704 标准的 Word 公文文档。

## 功能特性

| 特性 | 说明 |
|------|------|
| 红头封面 | 红色发文机关标志 + 发文字号 + 红色分隔线 |
| GB/T 9704 字体 | 方正小标宋标题、黑体一级标题、楷体二级标题、仿宋正文 |
| 行距规范 | 正文固定28.95pt（每面22行），标题单倍行距 |
| 自动引号转换 | 英文直引号 `"` `'` → 中文引号 `""` `''` |
| 三级标题加粗 | `1.` `2.` 开头的段落自动仿宋加粗（符合国标） |
| 字体自适配 | 生成前检测系统字体，缺方正小标宋等自动降级到等效字体，分享给他人无需预装字体 |
| 生成后验证 | 自动检查必需元素、引号残留、层级跳跃、附件编号 |

## 工作流程

```
判断输入来源
├─ 用户提供了 .md 文件 → 读取文件路径
├─ 对话中有 Markdown 内容 → 保存为临时 .md 文件
└─ 都没有 → 提示用户提供 Markdown 内容

       ↓

步骤 1：读取写作规范
  → 读取 references/公文Markdown写作规范.md

       ↓

步骤 2：转换 Markdown → JSON
  → 调用 scripts/md2json.py 自动解析
  → 自动处理：YAML 元数据提取、红头封面构建、附件识别、落款/日期右对齐
  → 用法：python3 scripts/md2json.py <input.md> <output.json>

       ↓

步骤 3：调用生成
  → 调用 skill 内的 scripts/generate.py（自包含，无需 docx_mcp）
  → 用法：python3 scripts/generate.py <output.json> <result.docx>
  → 自动执行：字体自适配检测、引号转换、三级标题加粗、标题行距设置

       ↓

步骤 4：读取验证报告
  → 脚本自动输出验证结果（errors / warnings）
  → 向用户展示检查结果

       ↓

步骤 5：返回结果
  → 告知用户文件路径
  → 用 open 命令打开文件
  → 如有警告或错误，提示用户确认
```

### 一键命令（步骤 2-3 可合并）

```bash
SKILL=~/.claude/skills/official-document-generator
python3 $SKILL/scripts/md2json.py input.md /tmp/content.json && \
python3 $SKILL/scripts/generate.py /tmp/content.json output.docx

# 用开源字体预设（思源系列，无需方正小标宋）：
# python3 $SKILL/scripts/generate.py /tmp/content.json output.docx --template official_document_compat
```

## 步骤 2 详细：Markdown → ContentBlock 映射规则

### 封面元数据

从 YAML 前置块提取，缺失则用占位符：

```
YAML org + "文件"  →  cover.title
YAML doc_no        →  cover.info_lines[0]
YAML title         →  meta.title
```

| 字段 | 默认值 |
|------|--------|
| `org` | `××××××` |
| `doc_no` | `×〔2026〕×号` |
| `type` | 从标题后缀推断：…的通知→通知，…的请示→请示，…的报告→报告，…的函→函 |

### 正文结构映射

按行解析 Markdown，使用以下规则：

| Markdown 元素 | 识别方式 | ContentBlock |
|--------------|---------|-------------|
| `# 标题` | 一级标题 | `{type:"heading", level:1, text:"标题"}` |
| `## 一、XXX` | 二级标题，含中文数字顿号 | `{type:"heading", level:2, text:"一、XXX"}` |
| `### （一）XXX` | 三级标题，含中文括号 | `{type:"heading", level:3, text:"（一）XXX"}` |
| 正文第一个以 `：` 结尾的行 | 在标题之后、其他段落之前 | `{type:"paragraph", text:"XXX：", alignment:"left"}` |
| 普通段落 | 非空行，非标题，非特殊元素 | `{type:"paragraph", text:"..."}` |
| `## 附件` | 二级标题，文字为"附件" | `{type:"paragraph", text:"附件：", space_before:60}`，后续列表项各自一个 paragraph |
| `---`（独立行） | 三个短横线 | `{type:"page_break"}` |
| `## 附件N：标题` | 二级标题，以"附件"开头 | `{type:"attachment_label", text:"附件N"}` + 空 paragraph + `{type:"heading", level:1, text:"标题"}` |
| Markdown 表格 | `\|...\|` 格式 | `{type:"table", table:{headers, rows, col_widths, row_heights}}` |

### 正文后的特殊元素

按位置识别（在 `## 附件` 之后、`---` 之前）：

**落款和日期**（正文末尾的非空短行）：
```
（发文机关名称）  →  {type:"paragraph", text:"（发文机关名称）", alignment:"right", space_before:120}
2026年6月5日     →  {type:"paragraph", text:"2026年6月5日", alignment:"right"}
```

识别规则：以 `XXXX年X月X日` 格式结尾的行 → 日期；日期上一行 → 落款。

**结尾用语**（根据公文类型自动判断，如果 Markdown 中有则保留）：
- 通知：特此通知。
- 请示：妥否，请批示。
- 报告：特此报告。
- 函：特此函达。

### 表格处理

Markdown 表格自动映射，默认列宽 `[2800, 6100]`（两列）或等分（多列）。

两列表格的行高规则（用于申请表等填写表）：
- 短字段（名称、电话等）：`500`
- 中等字段（含勾选项）：`500`
- 长文本字段（简介、方案等）：`1600`
- 意见栏：`1200`

表格下方的签字行：直接作为 paragraph，用全角空格 `　` 对齐。

## 步骤 3 详细：自动处理特性（由 builder 自动执行）

builder 在生成 Word 时会自动执行以下处理（无需手动配置）：

### 3.1 引号转换

当模板配置 `text_processing.convert_quotes: true` 时，所有 heading 和 paragraph 的文本自动执行：

```
英文直引号 " → 中文引号 ""（交替开闭）
英文直引号 ' → 中文引号 ''（交替开闭）
已是中文引号的不受影响
```

### 3.2 三级标题自动加粗

GB/T 9704 规定公文第三层标题（`1.` `2.` `3.` 开头的段落）应使用**仿宋加粗**。builder 自动检测以阿拉伯数字+半角句点开头的段落，对其加粗处理。

### 3.3 标题行距

一级标题（文件标题）使用独立行距，由模板 `paragraph.title_line_spacing` 配置：

| 配置值 | 效果 | 适用场景 |
|--------|------|---------|
| `"single"` | 单倍行距 | 标题不超过两行（默认） |
| `"35pt"` | 固定值35pt | 标题超过两行 |
| 不设置 | 沿用正文行距 | — |

### 3.4 调用方式

skill 自包含，直接调用 `scripts/generate.py` 即可（无需 MCP 服务器）：

```bash
python3 scripts/generate.py <input.json> <output.docx>
# 开源字体预设（思源系列，无需商业字体）：
python3 scripts/generate.py <input.json> <output.docx> --template official_document_compat
```

### 3.5 字体自适配（generate.py 自动执行）

`generate.py` 在构建文档前会检测系统已装字体。模板写死的「方正小标宋简体」「黑体-简」「楷体-简」「FangSong」若缺失，按平台自动降级到已装的等效字体（macOS/Windows 命名差异已处理），并打印「🔍 字体检测」报告：

- 本机字体齐全 → 全 ✓，不降级，行为不变
- 缺方正小标宋 → 自动降级（思源宋体粗 → 系统宋体），仍可生成可读公文
- 想严格国标 → 参考 `references/字体安装说明.md` 装字体，或用 `--template official_document_compat` 开源预设

降级映射规则见 `builder/fonts.py` 的 `FONT_FALLBACK_CHAIN`。

## 步骤 4 详细：验证检查

文档生成后，builder 自动运行以下检查：

| 检查项 | 类型 | 说明 |
|--------|------|------|
| 残留直引号 | ❌ 错误 | 扫描 .docx 全文（含表格），检查英文直引号残留 |
| 文件标题 | ❌ 错误 | 是否包含一级标题 |
| 主送机关 | ❌ 错误 | 正文开头是否有以全角冒号 `：` 结尾的段落 |
| 落款日期 | ❌ 错误 | 是否包含 `XXXX年X月X日` 格式的日期 |
| 发文机关 | ❌ 错误 | 是否包含 `（机关名称）` 格式的落款 |
| 标题层级跳跃 | ⚠️ 警告 | 是否出现标题层级跳跃（如从一级直接到三级） |
| 附件编号连续性 | ⚠️ 警告 | 附件编号是否连续（附件1、附件2…无缺失） |

**结果格式**：
- `passed: true/false` — 是否通过（有错误则为 false）
- `errors` — 错误列表（必须修复）
- `warnings` — 警告列表（建议检查）

收到结果后，向用户展示 `validation_report`，如有错误需提示修复。

## GB/T 9704 合规对照

| 要素 | 国标要求 | 模板实现 |
|------|---------|---------|
| 页面 | A4，上37mm/下35mm/左28mm/右26mm | ✅ `page.margins` |
| 正文 | 仿宋三号(16pt)，首行缩进2字符 | ✅ `fonts.body` + `paragraph.first_line_indent` |
| 文件标题 | 方正小标宋简体二号(22pt)，居中 | ✅ `fonts.heading.h1_*` |
| 一级标题（一、） | 黑体三号(16pt) | ✅ `fonts.heading.east_asia` |
| 二级标题（（一）） | 楷体三号(16pt) | ✅ `fonts.heading.h3_east_asia` |
| 三级标题（1.） | 仿宋加粗三号(16pt) | ✅ builder 自动检测加粗 |
| 四级标题（（1）） | 仿宋三号(16pt) | ✅ 沿用 body 字体 |
| 正文行距 | 固定28.95pt（每面22行） | ✅ `paragraph.line_spacing` |
| 标题行距 | 两行时单倍行距 | ✅ `paragraph.title_line_spacing` |
| 红头 | 红色方正小标宋 | ✅ `cover_page.title_font.color` |
| 红色分隔线 | 红线分隔红头与正文 | ✅ `cover_page.separator` |
| 页眉页脚 | 无页眉，有页码 | ✅ `header_footer` |

## 参考文档

开始工作前，**必须读取**以下文件了解完整规范：

- `references/公文Markdown写作规范.md` — 公文 Markdown 的完整写作规范，包含层级对照表和完整示例
- `references/official_document.yaml` — 公文模板配置（GB/T 9704），含所有参数及注释
- `references/公文模板实施计划.md` — 模板的设计决策记录和变更历史

## 文件结构

skill 完全自包含，builder/models/模板都在 skill 目录内，不依赖外部 MCP 服务器：

```
~/.claude/skills/official-document-generator/
├── skill.md                              ← 本文件
├── references/
│   ├── 公文Markdown写作规范.md            ← Markdown 写作规范
│   ├── official_document.yaml             ← 国标模板（方正小标宋等，需字体）
│   ├── official_document_compat.yaml      ← 开源字体预设（思源，免商业字体）
│   ├── 字体安装说明.md                     ← 字体获取与安装指引
│   └── 公文模板实施计划.md                 ← 设计决策与变更记录
├── scripts/
│   ├── md2json.py                        ← Markdown → JSON 转换器
│   └── generate.py                       ← JSON → Word 生成器（含字体自适配）
├── builder/                              ← 文档构建引擎（自包含）
│   ├── document_builder.py               ← 主编排器（引号转换、三级标题加粗）
│   ├── fonts.py                          ← 字体检测与自动降级
│   ├── validator.py                      ← 生成后验证检查
│   ├── styles.py                         ← 字体/行距/页面样式
│   ├── cover_page.py                     ← 红头封面 + 分隔线
│   ├── tables.py                         ← 表格渲染
│   ├── header_footer.py                  ← 页眉页脚
│   └── numbering.py                      ← 标题编号
└── models/
    └── content.py                        ← ContentBlock 数据模型
```

## 注意事项

1. **skill 自包含**：直接调 `scripts/generate.py` 即可，无需 docx_mcp MCP 服务器。模板直接读 `references/`，改 yaml 即时生效
2. **不要手动解析 Markdown**：使用 `scripts/md2json.py` 自动转换，避免手动构建 JSON 时遗漏映射规则
3. **不要发明内容**：如果 Markdown 中缺少某些公文要素（如发文字号），用占位符代替，不要自行编造
4. **保持层级顺序**：第一层（##）→ 第二层（###）→ 第三层（1.）→ 第四层（（1）），不可跳级
5. **附件编号规则**：单个附件写"附件：XXX"，多个附件写"附件1：""附件2："
6. **表格签字行**：用全角空格 `　` 对齐，不用普通空格
7. **如果 Markdown 不符合规范**：提示用户参考 `公文Markdown写作规范.md` 调整，不要强行解析
