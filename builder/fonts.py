"""字体自适配模块 —— 检测系统已装字体，对模板字体做可用性解析与自动降级。

公文模板（official_document.yaml）写死了「方正小标宋简体」「黑体-简」「楷体-简」「FangSong」
等字体名。这些字体在多数电脑上不存在：方正小标宋为商业字体（不可打包分发），
「黑体-简」「楷体-简」是 macOS 命名（Windows 不认）。本模块在生成文档前检测系统字体，
缺字体时按平台降级到已装的等效字体，保证分享给别人时无需对方安装任何字体也能生成可读公文。

字体检测依赖 fonttools（纯 Python，pip install fonttools）。未安装时跳过检测、保留原始字体名，
不会崩溃（此时退化为原有行为）。
"""

import sys
import subprocess
from pathlib import Path

# 支持的字体文件扩展名（.dfont 等 macOS 专有格式 fonttools 不支持，故排除）
_FONT_EXTS = {'.ttf', '.otf', '.ttc'}

# 模板中所有东亚字体字段的配置路径（覆盖 styles.py 与 cover_page.py 的全部字体注入点）
_FONT_FIELD_PATHS = [
    ('fonts', 'body', 'east_asia'),            # 正文：仿宋
    ('fonts', 'heading', 'east_asia'),         # 一级标题（一、）：黑体
    ('fonts', 'heading', 'h1_east_asia'),      # 文件标题：方正小标宋
    ('fonts', 'heading', 'h3_east_asia'),      # 二级标题（（一））：楷体
    ('cover_page', 'title_font', 'east_asia'),    # 红头发文机关：方正小标宋
    ('cover_page', 'subtitle_font', 'east_asia'), # 红头副标题：黑体
    ('cover_page', 'info_font', 'east_asia'),     # 发文字号：仿宋
]

# 每个原始字体名 → 降级候选列表（按优先级，已用平台实际命名）。
# 解析时：原始名已装则保留；否则取候选列表中第一个已装名。
FONT_FALLBACK_CHAIN = {
    # 文件标题 / 红头：方正小标宋（商业字体）→ 思源宋体粗（若装了）→ 系统宋体兜底
    # 标题本就加粗，普通宋体加粗后视觉可接受
    '方正小标宋简体': [
        'Source Han Serif SC Heavy', 'SourceHanSerifSC-Heavy',  # 思源宋体（SIL OFL，推荐替代）
        '思源宋体', '思源宋体 Heavy',
        'Songti SC', '宋体-简', 'STSong',                        # macOS
        'SimSun', '宋体', 'NSimSun',                              # Windows
    ],
    # 一级标题：黑体
    '黑体-简': [
        'SimHei', '黑体', 'Microsoft YaHei',                      # Windows
        'Heiti SC', 'STHeiti',                                    # macOS
        'Source Han Sans SC', 'SourceHanSansSC-Regular',          # 思源黑体
    ],
    # 二级标题：楷体（多数系统不自带，优先降级到仿宋——视觉最接近）
    '楷体-简': [
        'KaiTi', '楷体',                                           # Windows
        'STKaiti', 'Kaiti SC', '楷体-简',                          # macOS
        'STFangsong', '仿宋_GB2312', 'FangSong_GB2312', '仿宋',    # 仿宋（楷体最接近的常见替代）
        'Source Han Serif SC', 'SourceHanSerifSC-Regular',         # 思源宋体
        'STSong', 'Songti SC', '宋体-简', 'SimSun', '宋体',        # 宋体兜底
    ],
    # 正文：仿宋
    'FangSong': [
        '仿宋', '仿宋_GB2312', 'STFangsong', 'FangSong_GB2312',
        'SimSun', '宋体', 'Songti SC', 'STSong', '宋体-简',        # 兜底宋体（几乎所有中文系统都有）
    ],
    # === compat 模板的思源字体（SIL OFL 开源）→ 系统字体兜底 ===
    # 装了思源就用思源；没装则降级到系统宋体/黑体，保证 compat 在无思源环境也可用
    'Source Han Serif SC Heavy': [          # compat 文件标题/红头
        '方正小标宋简体', 'FZXiaoBiaoSong-B05S',
        'Songti SC', '宋体-简', 'STSong', 'SimSun', '宋体',
    ],
    'Source Han Sans SC': [                  # compat 一级标题/红头副标题
        '黑体-简', 'Heiti SC', 'STHeiti',
        'SimHei', '黑体', 'Microsoft YaHei',
    ],
    'Source Han Serif SC': [                 # compat 正文/二级标题
        'FangSong', '仿宋_GB2312', 'STFangsong', '仿宋',
        'Songti SC', '宋体-简', 'STSong', 'SimSun', '宋体',
    ],
}


def _font_dirs():
    """返回当前平台存在的字体目录列表。"""
    if sys.platform == 'darwin':
        home = Path.home()
        candidates = [
            '/System/Library/Fonts',
            '/Library/Fonts',
            str(home / 'Library/Fonts'),
            '/System/Library/Fonts/Supplemental',
        ]
    elif sys.platform == 'win32':
        local = __import__('os').environ.get('LOCALAPPDATA', '')
        candidates = [
            r'C:\Windows\Fonts',
            str(Path(local) / 'Microsoft' / 'Windows' / 'Fonts') if local else '',
        ]
    else:  # linux 及其他
        home = Path.home()
        candidates = [
            str(home / '.fonts'),
            str(home / '.local/share/fonts'),
            '/usr/share/fonts',
            '/usr/local/share/fonts',
        ]
    return [Path(d) for d in candidates if d and Path(d).exists()]


def _iter_ttf_fonts(path):
    """对一个字体文件，yield 其包含的所有 TTFont（.ttc 含多个 face）。失败则不 yield。"""
    from fontTools.ttLib import TTFont, TTCollection
    # .ttc 先按 collection 读（一次拿到全部 face）
    try:
        coll = TTCollection(str(path))
        for f in coll:
            yield f
        return
    except Exception:
        pass
    # 单字体文件（.ttf / .otf）
    try:
        yield TTFont(str(path))
    except Exception:
        # 极少数情况：TTCollection 解析失败的 .ttc，逐 fontNumber 遍历
        i = 0
        while True:
            try:
                yield TTFont(str(path), fontNumber=i)
                i += 1
            except Exception:
                break


def _read_family_names(path):
    """读出字体文件的所有 family name（nameID 1 与 16，不区分语言），返回名字集合。"""
    names = set()
    for tt in _iter_ttf_fonts(path):
        name_table = tt.get('name')
        if not name_table:
            continue
        for rec in name_table.names:
            if rec.nameID in (1, 16):
                try:
                    names.add(rec.toUnicode())
                except Exception:
                    pass
    return names


def _fc_list_fonts():
    """Linux 兜底：用 fc-list 列出字体族名。非 linux 或无 fc-list 时返回空集。"""
    names = set()
    try:
        out = subprocess.run(
            ['fc-list', ':', 'family'],
            capture_output=True, text=True, timeout=15,
        )
        for line in out.stdout.splitlines():
            nm = line.strip()
            if nm:
                names.add(nm)
    except Exception:
        pass
    return names


def detect_system_fonts():
    """枚举系统字体目录，返回所有已安装字体族名的集合。

    未安装 fonttools 时返回 None（调用方据此跳过降级，保留原始字体名）。
    """
    try:
        import fontTools.ttLib  # noqa: F401
    except ImportError:
        print('⚠️  未安装 fonttools，跳过字体检测（pip install fonttools 可启用字体自适配）')
        return None

    installed = set()
    for d in _font_dirs():
        for f in d.rglob('*'):
            if f.suffix.lower() in _FONT_EXTS:
                installed |= _read_family_names(f)

    if sys.platform.startswith('linux'):
        installed |= _fc_list_fonts()

    return installed


def _get_path(d, path):
    cur = d
    for k in path:
        if not isinstance(cur, dict) or k not in cur:
            return None
        cur = cur[k]
    return cur


def _set_path(d, path, val):
    cur = d
    for k in path[:-1]:
        cur = cur[k]
    cur[path[-1]] = val


def resolve_fonts(cfg, installed):
    """对 cfg 中所有字体字段做可用性解析与降级。

    Args:
        cfg: 已加载的模板配置 dict（会被原地修改）。
        installed: detect_system_fonts() 的返回值；None 表示无法检测，直接原样返回。

    Returns:
        (cfg, report)。report 为 list[dict]，每项含 field / orig / resolved / status：
          'ok'       原始字体已安装，保留
          'fallback' 原始字体缺失，已降级到等效字体
          'missing'  原始字体与全部候选均缺失，保留原始名（交由 Word 渲染时替换）
    """
    report = []
    if installed is None:
        return cfg, report

    for path in _FONT_FIELD_PATHS:
        orig = _get_path(cfg, path)
        if not orig:
            continue
        field = '.'.join(path)

        if orig in installed:
            report.append({'field': field, 'orig': orig, 'resolved': orig, 'status': 'ok'})
            continue

        candidates = FONT_FALLBACK_CHAIN.get(orig, [])
        resolved = next((c for c in candidates if c in installed), None)
        if resolved:
            _set_path(cfg, path, resolved)
            report.append({'field': field, 'orig': orig, 'resolved': resolved, 'status': 'fallback'})
        else:
            report.append({'field': field, 'orig': orig, 'resolved': orig, 'status': 'missing'})

    return cfg, report


def print_font_report(report):
    """向 stdout 打印字体检测结果。"""
    if not report:
        return
    print('🔍 字体检测:')
    for r in report:
        if r['status'] == 'ok':
            print(f"  ✓ {r['orig']}")
        elif r['status'] == 'fallback':
            print(f"  ✗ {r['orig']} → 已降级为 \"{r['resolved']}\"")
        else:  # missing
            print(f"  ⚠ {r['orig']} → 未安装且无等效替代，将由 Word 自动替换（可能不符国标）")

    fb = sum(1 for r in report if r['status'] == 'fallback')
    ms = sum(1 for r in report if r['status'] == 'missing')
    if fb or ms:
        msg = '  ℹ️  '
        parts = []
        if fb:
            parts.append(f'{fb} 项已自动降级')
        if ms:
            parts.append(f'{ms} 项无替代')
        msg += '，'.join(parts) + '。如需严格国标字体，参考 references/字体安装说明.md'
        print(msg)
