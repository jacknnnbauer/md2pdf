#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
md2pdf —— 中文技术文档 Markdown 排版转 PDF

用法:
    python md2pdf.py 文档.md
    python md2pdf.py 文档.md -o 输出.pdf --theme crimson --preview 1

特性:
    · YAML front-matter 配置封面与版式，无 front-matter 也能跑
    · 自动扫描 ## / ### 标题生成目录，多遍渲染回填真实页码
    · Chrome 无头浏览器排版（表格跨页表头重复、分页控制、中文字体栈）
    · PyMuPDF 后处理：跑动页眉、页脚页码、PDF 书签、元数据、字体子集化

依赖:
    pip install markdown pymupdf pyyaml
    并需本机安装 Chrome 或 Edge

作者备注:
    整套方案的取舍见同目录 README.md
"""
from __future__ import annotations

import argparse
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile

# Windows 控制台默认 GBK，输出 ✔ 等字符会崩；统一转 UTF-8 并容错
for _s in ("stdout", "stderr"):
    try:
        getattr(sys, _s).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

try:
    import markdown
except ImportError:
    sys.exit("缺少依赖：pip install markdown")
try:
    import pymupdf
except ImportError:
    sys.exit("缺少依赖：pip install pymupdf")
try:
    import yaml
except ImportError:
    yaml = None  # 没装也能跑，只是不解析 front-matter


# ══════════════════════════════════════════════════════════════════
#  1. 主题配色
# ══════════════════════════════════════════════════════════════════
THEMES = {
    # 名称       主色       深色      浅底       次浅底     强调色
    "navy":    dict(main="#1F3864", dark="#0d1b2e", tint="#F6F8FA", band="#DDE5F0", accent="#C00000"),
    "slate":   dict(main="#37474F", dark="#1a2327", tint="#F5F7F8", band="#DEE4E7", accent="#B5651D"),
    "forest":  dict(main="#1E4620", dark="#0e2410", tint="#F5F9F5", band="#DCE8DC", accent="#8C6D1F"),
    "crimson": dict(main="#7B1E28", dark="#3d0f14", tint="#FAF6F6", band="#EDDCDE", accent="#1F3864"),
    "ink":     dict(main="#222222", dark="#000000", tint="#F7F7F7", band="#E2E2E2", accent="#A03030"),
}

PAPERS = {
    "a4":     "210mm 297mm",
    "letter": "216mm 279mm",
    "a3":     "297mm 420mm",
    "b5":     "176mm 250mm",
}

# 提示框（blockquote）配色固定为暖黄，与主色形成对比
QUOTE_BG, QUOTE_BAR, QUOTE_FG = "#FFF8E6", "#E8A33D", "#8a5200"


# ══════════════════════════════════════════════════════════════════
#  2. 浏览器与字体探测
# ══════════════════════════════════════════════════════════════════
BROWSERS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser",
]

# PyMuPDF 盖页眉页脚用的 CJK 字体文件（内置 PDF 字体不含汉字，必须外挂）
CJK_FONTS = [
    r"C:\Windows\Fonts\msyh.ttc",      # 微软雅黑
    r"C:\Windows\Fonts\simhei.ttf",    # 黑体
    r"C:\Windows\Fonts\Deng.ttf",      # 等线
    r"C:\Windows\Fonts\simsun.ttc",    # 宋体
    "/System/Library/Fonts/PingFang.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
]


def find_browser(override: str | None = None) -> str:
    if override:
        if os.path.exists(override):
            return override
        sys.exit("指定的浏览器不存在：%s" % override)
    for p in BROWSERS:
        if os.path.exists(p):
            return p
    for name in ("chrome", "msedge", "chromium", "google-chrome"):
        p = shutil.which(name)
        if p:
            return p
    sys.exit("未找到 Chrome / Edge。请安装其一，或用 --browser 指定可执行文件路径。")


def find_cjk_font(override: str | None = None) -> str | None:
    if override and os.path.exists(override):
        return override
    for p in CJK_FONTS:
        if os.path.exists(p):
            return p
    return None


# ══════════════════════════════════════════════════════════════════
#  3. front-matter 解析
# ══════════════════════════════════════════════════════════════════
FM_RE = re.compile(r"\A\ufeff?---\s*\n(.*?)\n---\s*\n", re.S)


def split_front_matter(text: str):
    """返回 (配置 dict, 去掉 front-matter 的正文)"""
    m = FM_RE.match(text)
    if not m:
        return {}, text
    raw, body = m.group(1), text[m.end():]
    if yaml is None:
        print("! 未安装 pyyaml，front-matter 被忽略（pip install pyyaml）")
        return {}, body
    try:
        cfg = yaml.safe_load(raw) or {}
    except Exception as e:
        print("! front-matter 解析失败，已忽略：%s" % e)
        return {}, body
    if not isinstance(cfg, dict):
        return {}, body
    return cfg, body


# ══════════════════════════════════════════════════════════════════
#  4. 正文预处理：剥离、扫描标题、注入定位 token
# ══════════════════════════════════════════════════════════════════
TOC_TITLES = {"目录", "目 录", "目　录", "contents", "table of contents", "toc"}
FENCE_RE = re.compile(r"^\s*(```|~~~)")


def drop_leading_h1(md: str):
    """摘掉正文第一个 # 标题（PDF 用封面代替），返回 (标题文本, 剩余正文)"""
    m = re.match(r"\A\s*#\s+(.+?)\s*\n", md)
    if not m:
        return None, md
    return m.group(1).strip(), md[m.end():]


def drop_leading_blockquote(md: str) -> str:
    """摘掉正文开头的 > 引用块（很多文档用它写元信息，已由封面承担）"""
    return re.sub(r"\A\s*(?:>[^\n]*\n)+\s*", "", md)


def drop_toc_section(md: str) -> str:
    """删掉正文里手写的「## 目录 …」整节（PDF 用自动生成的带页码目录）"""
    lines, out, killing, fence = md.split("\n"), [], False, False
    for ln in lines:
        if FENCE_RE.match(ln):
            fence = not fence
        if not fence:
            m = re.match(r"^(#{1,3})\s+(.+?)\s*$", ln)
            if m:
                t = re.sub(r"[\s#*]+", "", m.group(2)).lower()
                if t in {x.replace(" ", "").replace("　", "") for x in TOC_TITLES}:
                    killing = True
                    continue
                killing = False
            elif killing and (ln.strip() in ("", "---") or ln.lstrip().startswith(("-", "*", "+"))):
                continue
        if not killing:
            out.append(ln)
    return "\n".join(out)


def scan_headings(md: str, depth: int):
    """扫描 ## / ### 标题（跳过代码块），返回 [(level, text, hid, token)]"""
    heads, fence, i = [], False, 0
    for ln in md.split("\n"):
        if FENCE_RE.match(ln):
            fence = not fence
            continue
        if fence:
            continue
        m = re.match(r"^(#{2,3})\s+(.+?)\s*$", ln)
        if not m:
            continue
        lv = len(m.group(1)) - 1          # ## -> 1, ### -> 2
        if lv > depth:
            continue
        i += 1
        txt = re.sub(r"[*_`]", "", m.group(2)).strip()
        txt = re.sub(r"\[(.+?)\]\([^)]*\)", r"\1", txt)   # 去掉行内链接语法
        txt = re.sub(r"<[^>]+>", "", txt).strip()          # 去掉行内 HTML
        heads.append((lv, txt, "h%03d" % i, "@@H%03d@@" % i))
    return heads


def inject_tokens(md: str, heads, depth: int) -> str:
    """
    给每个标题行尾追加：不可见定位 token + attr_list 锚点 id
    token 用于渲染后在 PDF 文本层里定位该标题落在第几页
    """
    lines, fence, i = md.split("\n"), False, 0
    for n, ln in enumerate(lines):
        if FENCE_RE.match(ln):
            fence = not fence
            continue
        if fence:
            continue
        m = re.match(r"^(#{2,3})\s+(.+?)\s*$", ln)
        if not m:
            continue
        if len(m.group(1)) - 1 > depth:
            continue
        _, _, hid, tok = heads[i]
        lines[n] = '%s %s<span class="tk">%s</span> {: #%s }' % (
            m.group(1), m.group(2).rstrip(), tok, hid)
        i += 1
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════
#  5. CSS / 封面 / 目录 模板
# ══════════════════════════════════════════════════════════════════
CSS_TPL = r"""
@page { size: %(paper)s; margin: %(margin)s; }
@page :first { margin: 0; }
html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }

body{
  font-family:%(font_body)s;
  font-size:%(fs_body)spt; line-height:%(lh)s; color:#1a1a1a; margin:0;
  text-align:justify; word-break:normal; overflow-wrap:anywhere;
}
/* 定位 token：透明但仍进文本层，绝不能用 display:none */
.tk{ color:transparent; font-size:1px; letter-spacing:0; }

/* ─────────── 封面 ─────────── */
.cover{ height:%(page_h)s; box-sizing:border-box; page-break-after:always;
        position:relative; background:#fff; }
.cover .band{ height:%(band_h)s; background:%(main)s; color:#fff;
              padding:24mm 20mm 0 20mm; box-sizing:border-box; }
.cover .kicker{ font-size:10pt; letter-spacing:3px; opacity:.72; margin-bottom:7mm; }
.cover h1.t{ font-family:%(font_head)s !important; font-size:%(fs_title)spt !important;
             font-weight:700 !important; line-height:1.34 !important;
             margin:0 !important; padding:0 !important; border:0 !important;
             letter-spacing:1px; color:#ffffff !important;
             page-break-before:avoid !important; }
.cover .sub{ font-size:11.5pt; color:#ffffff; opacity:.88; margin-top:5mm; letter-spacing:1px; }
.cover .body{ padding:14mm 20mm 0 20mm; }
.cover .rule{ height:3px; width:34mm; background:%(accent)s; margin-bottom:9mm; }
.cover dl{ margin:0; font-size:10pt; line-height:2.0; }
.cover dt{ float:left; clear:left; width:%(dt_w)s; color:#7a7a7a; }
.cover dd{ margin:0 0 0 %(dt_w)s; }
.cover .notice{ margin-top:12mm; border-left:4px solid %(accent)s; background:#FDF2F2;
                padding:5mm 6mm; font-size:9.5pt; line-height:1.75; }
.cover .notice p{ margin:0 0 2mm 0; }
.cover .notice p:last-child{ margin-bottom:0; }
.cover .foot{ position:absolute; bottom:18mm; left:20mm; right:20mm;
              font-size:8.5pt; color:#8a8a8a; border-top:1px solid #ddd; padding-top:4mm; }

/* ─────────── 目录 ─────────── */
.toc{ page-break-after:always; }
.toc h2.tt{ font-size:17pt; border:0; margin:0 0 8mm 0; padding:0; color:%(main)s;
            page-break-before:avoid; }
.toc ul{ list-style:none; padding:0; margin:0; }
.toc li{ margin:0; padding:2.6mm 0; border-bottom:1px dotted #cfd6e0; font-size:10.5pt; }
.toc li.l2{ padding:1.8mm 0 1.8mm 8mm; font-size:9.6pt; border-bottom:1px dotted #e6ebf1; color:#444; }
.toc li a{ color:#1a1a1a; text-decoration:none; }
.toc li.l2 a{ color:#444; }
.toc .pg{ float:right; color:%(main)s; font-weight:700; }

/* ─────────── 标题 ─────────── */
h1{ font-family:%(font_head)s; font-size:20pt; color:%(main)s; }
h2{ font-family:%(font_head)s; font-size:16.5pt; color:%(main)s;
    margin:0 0 6mm 0; padding:0 0 3mm 0; border-bottom:2.5px solid %(main)s;
    %(break_h2)s page-break-after:avoid; letter-spacing:.5px; }
h3{ font-family:%(font_head)s; font-size:12.5pt; color:%(main)s;
    margin:8mm 0 3mm 0; padding-left:3.5mm; border-left:4px solid %(main)s;
    page-break-after:avoid; }
h4{ font-family:%(font_head)s; font-size:10.8pt; color:%(dark)s;
    margin:6mm 0 2mm 0; page-break-after:avoid; }
h5,h6{ font-family:%(font_head)s; font-size:10pt; color:#555; margin:5mm 0 2mm 0;
       page-break-after:avoid; }
body > h2:first-of-type{ page-break-before:avoid; }
/* 分隔线不得单独占页：与其后内容绑定，避免在强制分页的 h2 前空出一整页 */
hr{ page-break-after:avoid; break-after:avoid; page-break-inside:avoid; }

/* ─────────── 正文元素 ─────────── */
p{ margin:0 0 3mm 0; }
strong{ color:%(dark)s; }
em{ font-style:normal; background:linear-gradient(transparent 62%%, %(band)s 62%%); }
hr{ border:0; border-top:1px solid #dde2e8; margin:7mm 0; }
a{ color:%(main)s; text-decoration:none; }
img{ max-width:100%%; page-break-inside:avoid; }
ul,ol{ margin:0 0 3mm 0; padding-left:6mm; }
li{ margin:0 0 1.2mm 0; }

code{ font-family:%(font_mono)s; font-size:%(fs_code)spt;
      background:#F1F3F6; padding:1px 4px; border-radius:3px; color:#7a2020; }
pre{ background:#F7F9FB; border:1px solid #E1E6EC; border-left:3px solid %(main)s;
     padding:4mm 5mm; font-size:%(fs_code)spt; line-height:1.55;
     white-space:pre-wrap; overflow-x:hidden; page-break-inside:avoid; border-radius:2px; }
pre code{ background:none; padding:0; color:#1a1a1a; font-size:inherit; }

blockquote{ margin:4mm 0; padding:3.5mm 5mm; background:%(q_bg)s;
            border-left:4px solid %(q_bar)s; font-size:9.4pt; line-height:1.68;
            page-break-inside:avoid; border-radius:2px; }
blockquote p{ margin:0 0 2mm 0; }
blockquote p:last-child{ margin-bottom:0; }
blockquote strong{ color:%(q_fg)s; }

/* ─────────── 表格 ─────────── */
table{ border-collapse:collapse; width:100%%; margin:3.5mm 0 5mm 0;
       font-size:%(fs_table)spt; line-height:1.5; table-layout:auto; }
thead{ display:table-header-group; }        /* 跨页时表头自动重复 */
tr{ page-break-inside:avoid; }
th{ background:%(main)s; color:#fff; font-weight:600; text-align:left;
    padding:2.2mm 2.4mm; border:1px solid %(main)s; font-size:%(fs_table)spt; }
td{ padding:2.0mm 2.4mm; border:1px solid #D3DAE3; vertical-align:top;
    text-align:left; word-break:normal; overflow-wrap:anywhere; }
tbody tr:nth-child(even) td{ background:%(tint)s; }
td strong{ color:%(dark)s; }
table code{ font-size:%(fs_tcode)spt; }

h2 + p, h2 + blockquote, h2 + table{ page-break-before:avoid; }
%(extra_css)s
"""

COVER_TPL = """
<div class="cover">
  <div class="band">
    %(kicker)s
    <h1 class="t">%(title)s</h1>
    %(subtitle)s
  </div>
  <div class="body">
    <div class="rule"></div>
    %(meta)s
    %(notice)s
  </div>
  %(foot)s
</div>
"""


def esc(s) -> str:
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def build_cover(cfg: dict, title: str) -> str:
    if not cfg.get("cover", True):
        return ""
    kicker = cfg.get("kicker", "")
    sub = cfg.get("subtitle", "") or cfg.get("version", "")
    meta = cfg.get("meta") or {}
    notice = cfg.get("notice", "")
    foot = cfg.get("cover_footer", "") or cfg.get("footer", "")

    meta_html = ""
    if isinstance(meta, dict) and meta:
        rows = "".join("<dt>%s</dt><dd>%s</dd>" % (esc(k), md_inline(str(v)))
                       for k, v in meta.items())
        meta_html = "<dl>%s</dl>" % rows

    notice_html = ""
    if notice:
        notice_html = '<div class="notice">%s</div>' % markdown.markdown(
            str(notice), extensions=["nl2br"])

    return COVER_TPL % dict(
        kicker='<div class="kicker">%s</div>' % esc(kicker) if kicker else "",
        title=str(title).replace("\n", "<br>"),
        subtitle='<div class="sub">%s</div>' % esc(sub) if sub else "",
        meta=meta_html,
        notice=notice_html,
        foot='<div class="foot">%s</div>' % esc(foot) if foot else "",
    )


def md_inline(s: str) -> str:
    """把单行 Markdown 渲染成内联 HTML（去掉外层 <p>）"""
    h = markdown.markdown(s)
    return re.sub(r"^<p>(.*)</p>$", r"\1", h, flags=re.S)


def build_toc(heads, pages, cfg) -> str:
    if not cfg.get("toc", True) or not heads:
        return ""
    label = cfg.get("toc_title", "目　录")
    out = ['<div class="toc"><h2 class="tt">%s</h2><ul>' % esc(label)]
    for lv, txt, hid, tok in heads:
        pg = ""
        if pages and tok in pages:
            pg = '<span class="pg">%d</span>' % pages[tok]
        out.append('<li class="l%d"><a href="#%s">%s</a>%s</li>' % (lv, hid, esc(txt), pg))
    out.append("</ul></div>")
    return "".join(out)


# ══════════════════════════════════════════════════════════════════
#  6. 渲染
# ══════════════════════════════════════════════════════════════════
def make_css(cfg: dict) -> str:
    th = THEMES.get(str(cfg.get("theme", "navy")).lower(), THEMES["navy"])
    paper = str(cfg.get("paper", "A4")).lower()
    size = PAPERS.get(paper, PAPERS["a4"])
    land = bool(cfg.get("landscape", False))
    w, h = size.split()
    if land:
        w, h = h, w
        size = "%s %s" % (w, h)
    return CSS_TPL % dict(
        paper=size,
        page_h=h,
        band_h=cfg.get("band_height", "95mm"),
        margin=cfg.get("margin", "20mm 16mm 20mm 16mm"),
        dt_w=cfg.get("meta_label_width", "26mm"),
        font_body=cfg.get("font_body", '"Segoe UI","DengXian","Microsoft YaHei",sans-serif'),
        font_head=cfg.get("font_head", '"Microsoft YaHei",sans-serif'),
        font_mono=cfg.get("font_mono", 'Consolas,"Cascadia Mono",monospace'),
        fs_body=cfg.get("font_size", 10),
        fs_title=cfg.get("title_size", 26),
        fs_table=cfg.get("table_font_size", 8.6),
        fs_code=cfg.get("code_font_size", 8.6),
        fs_tcode=float(cfg.get("table_font_size", 8.6)) - 0.6,
        lh=cfg.get("line_height", 1.72),
        break_h2="page-break-before:always;" if cfg.get("break_before_h2", True) else "",
        extra_css=cfg.get("extra_css", ""),
        q_bg=QUOTE_BG, q_bar=QUOTE_BAR, q_fg=QUOTE_FG,
        **th
    )


ANCHOR_P = re.compile(r"<p>\s*((?:<a[^>]*>\s*</a>\s*)+)</p>")
HR_BEFORE_H = re.compile(r"<hr\s*/?>\s*((?:<a[^>]*>\s*</a>\s*)*)(?=<h[12])")


def clean_html(body: str) -> str:
    """
    ① 锚点段落 <p><a id=x></a></p> 降为内联，避免多出一个带外边距的空段
    ② 删除紧邻 h1/h2 之前的 <hr>
       —— h2 带 page-break-before:always，若 hr 恰好落到新页顶部，
          会先占一页、再把标题推到下一页，中间空出整页
    """
    body = ANCHOR_P.sub(r"\1", body)
    body = HR_BEFORE_H.sub(r"\1", body)
    return body


def build_html(body_md: str, heads, cfg: dict, title: str, pages) -> str:
    body = markdown.markdown(body_md, extensions=[
        "tables", "fenced_code", "attr_list", "md_in_html", "sane_lists", "nl2br",
    ] + list(cfg.get("md_extensions", [])))
    body = clean_html(body)
    return ("<!doctype html><html><head><meta charset=\"utf-8\">"
            "<title>%s</title><style>%s</style></head><body>%s%s%s</body></html>"
            % (esc(title), make_css(cfg), build_cover(cfg, title),
               build_toc(heads, pages, cfg), body))


def render_pdf(html: str, browser: str, work: str, tag: str, timeout: int = 300) -> str:
    hp = os.path.join(work, "_m2p_%s.html" % tag)
    pp = os.path.join(work, "_m2p_%s.pdf" % tag)
    io.open(hp, "w", encoding="utf-8").write(html)
    if os.path.exists(pp):
        os.remove(pp)
    # 独立临时 profile：否则命令会被转发给已在运行的 Chrome，静默不出 PDF
    prof = tempfile.mkdtemp(prefix="m2p_")
    cmd = [browser, "--headless=new", "--disable-gpu", "--no-sandbox",
           "--no-pdf-header-footer",
           "--run-all-compositor-stages-before-draw",
           "--virtual-time-budget=20000",
           "--user-data-dir=" + prof,
           "--print-to-pdf=" + pp,
           "file:///" + hp.replace("\\", "/")]
    try:
        subprocess.run(cmd, capture_output=True, timeout=timeout)
    finally:
        shutil.rmtree(prof, ignore_errors=True)
    if not os.path.exists(pp):
        raise RuntimeError("浏览器未生成 PDF（tag=%s）。检查 --browser 路径。" % tag)
    return pp


def locate_tokens(pdf_path: str, heads):
    """在 PDF 文本层里找每个标题 token 落在第几页（1-based）"""
    d = pymupdf.open(pdf_path)
    pages, want = {}, {tok for _, _, _, tok in heads}
    for i in range(d.page_count):
        t = d[i].get_text()
        for tok in want:
            if tok not in pages and tok in t:
                pages[tok] = i + 1
    n = d.page_count
    d.close()
    return pages, n


# ══════════════════════════════════════════════════════════════════
#  7. 后处理：页眉 / 页脚 / 书签 / 元数据 / 子集化
# ══════════════════════════════════════════════════════════════════
def cjk_width(text: str, fs: float) -> float:
    """估算文本宽度：汉字与全角标点按 1.0em，其余按 0.55em
    （PyMuPDF 的 get_text_length 只支持内置字体，自定义字体必须自己算）"""
    return sum((fs if ord(ch) > 0x2E80 else fs * 0.55) for ch in text)


def postprocess(pdf_in: str, out: str, heads, pages, cfg, title: str, font_file):
    doc = pymupdf.open(pdf_in)
    N = doc.page_count
    has_cover = bool(cfg.get("cover", True))
    show_pn = bool(cfg.get("page_numbers", True))
    hl = cfg.get("header_left", title)
    fl = cfg.get("footer_left", cfg.get("footer", ""))
    th = THEMES.get(str(cfg.get("theme", "navy")).lower(), THEMES["navy"])
    rgb = tuple(int(th["main"][i:i + 2], 16) / 255 for i in (1, 3, 5))

    fbuf = io.open(font_file, "rb").read() if font_file else None
    NAV = "m2pcjk"

    # 页 -> 一级标题，用于跑动页眉
    start = {}
    for lv, txt, hid, tok in heads:
        if lv == 1 and tok in pages:
            start.setdefault(pages[tok], txt)

    cur = ""
    for i in range(N):
        pg = doc[i]
        W, H = pg.rect.width, pg.rect.height
        if has_cover and i == 0:
            continue
        if (i + 1) in start:
            cur = start[i + 1]
        if fbuf:
            pg.insert_font(fontname=NAV, fontbuffer=fbuf)
        FN = NAV if fbuf else "helv"

        # 页眉
        if hl or cur:
            y = 34
            pg.draw_line(pymupdf.Point(45, y + 6), pymupdf.Point(W - 45, y + 6),
                         color=(0.82, 0.85, 0.89), width=0.6)
            if hl:
                pg.insert_text(pymupdf.Point(45, y), str(hl),
                               fontname=FN, fontsize=7.2, color=(0.55, 0.58, 0.62))
            if cur:
                txt, fs = cur, 7.2
                limit = W - 45 - (cjk_width(str(hl), fs) + 40 if hl else 40)
                while cjk_width(txt, fs) > limit and len(txt) > 8:
                    txt = txt[:-2]
                if txt != cur:
                    txt += "…"
                pg.insert_text(pymupdf.Point(W - 45 - cjk_width(txt, fs), y), txt,
                               fontname=FN, fontsize=fs, color=rgb)

        # 页脚
        yf = H - 26
        pg.draw_line(pymupdf.Point(45, yf - 10), pymupdf.Point(W - 45, yf - 10),
                     color=(0.82, 0.85, 0.89), width=0.6)
        if show_pn:
            lab = "%d / %d" % (i + 1, N)
            lw = pymupdf.get_text_length(lab, fontname="helv", fontsize=8)
            pg.insert_text(pymupdf.Point((W - lw) / 2, yf), lab,
                           fontname="helv", fontsize=8, color=(0.35, 0.38, 0.42))
        if fl:
            pg.insert_text(pymupdf.Point(45, yf), str(fl),
                           fontname=FN, fontsize=7, color=(0.62, 0.65, 0.68))

    # PDF 书签
    toc_out = []
    if has_cover:
        toc_out.append([1, "封面", 1])
    if cfg.get("toc", True) and heads:
        toc_out.append([1, "目录", 2 if has_cover else 1])
    for lv, txt, hid, tok in heads:
        if tok in pages:
            toc_out.append([lv, txt, pages[tok]])
    if toc_out:
        doc.set_toc(toc_out)

    doc.set_metadata({
        "title": str(cfg.get("pdf_title", title)),
        "author": str(cfg.get("author", "")),
        "subject": str(cfg.get("subject", "")),
        "keywords": str(cfg.get("keywords", "")),
        "creator": "md2pdf",
    })

    # 抹除定位用的隐形 token。它们视觉不可见（transparent / 1px），
    # 但留在文本层：复制、Ctrl+F、文本提取、屏幕朗读都会带出来。
    # 用 redact 而不是重新渲染，保证版面零位移。
    n_red = 0
    for page in doc:
        hit = 0
        for _lv, _txt, _hid, tok in heads:
            for rect in page.search_for(tok):
                page.add_redact_annot(rect)
                hit += 1
        if hit:
            page.apply_redactions(images=0, graphics=0, text=0)
            n_red += hit
    if n_red:
        print("  已抹除 %d 个隐形定位 token" % n_red)

    try:
        doc.subset_fonts(verbose=False)      # 整套雅黑 ~14MB -> 子集 ~2MB
    except Exception as e:
        print("! 字体子集化跳过：%s" % e)

    tmp = out + ".tmp"
    doc.save(tmp, garbage=4, deflate=True)
    doc.close()
    if os.path.exists(out):
        os.remove(out)
    shutil.move(tmp, out)


# ══════════════════════════════════════════════════════════════════
#  8. 主流程
# ══════════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════════
#  出片自检
#  ——这些坑都是真实返工出来的，机器能查的就别靠眼睛
# ══════════════════════════════════════════════════════════════════
# 机器残渣：出现即是 bug，报错
RESIDUE = ["@@H", "{{", "}}", "[object Object]"]
# 人写的占位符：可能是正文在讨论它们，只提示不报错
DRAFT_MARKS = ["TODO", "XXX", "???", "待补", "待填", "undefined", "NaN"]


def selfcheck(pdf_path: str, heads, cfg, quiet=False):
    """生成后自检。返回 (errors, warnings) 两个列表。

    查四类问题：
      1) 空白页 / 稀疏页   —— 分页规则打架的典型症状
      2) 隐形 token 与占位符残留 —— 肉眼看不见，复制/搜索会露馅
      3) 目录印的页码 vs 标题实际所在页 —— 多遍收敛失败时会不一致
      4) 页面溢出（内容超出纸张可打印区）
    """
    err, warn = [], []
    d = pymupdf.open(pdf_path)
    txts = [p.get_text() for p in d]

    # 1) 空白页 / 稀疏页
    has_cover = bool(cfg.get("cover", True))
    skip = set()
    if has_cover:
        skip.add(0)                       # 封面本来就字少
    if cfg.get("toc", True) and heads:
        skip.add(1 if has_cover else 0)   # 目录页同理
    for i, t in enumerate(txts):
        if i in skip:
            continue
        n = len(t.strip())
        if n < 120:
            err.append("p%d 近乎空白（%d 字）——检查 --- 分隔线与 page-break 是否打架" % (i + 1, n))
        elif n < 300:
            warn.append("p%d 内容稀疏（%d 字）" % (i + 1, n))

    # 2) 占位符与隐形 token 残留
    full = "\n".join(txts)
    for k in RESIDUE:
        c = full.count(k)
        if c:
            err.append("文本层残留「%s」×%d —— 肉眼看不见，但复制/Ctrl+F/朗读会带出来" % (k, c))
    for k in DRAFT_MARKS:
        c = full.count(k)
        if c:
            warn.append("出现草稿标记「%s」×%d —— 确认是正文在讨论它，而不是忘了填" % (k, c))

    # 3) 目录页码 vs 实际页码
    if cfg.get("toc", True) and heads:
        real = {}
        for i, p in enumerate(d):
            for blk in p.get_text("dict")["blocks"]:
                for ln in blk.get("lines", []):
                    for sp in ln["spans"]:
                        t = sp["text"].strip()
                        if sp["size"] > 15 and t and t not in real:
                            real[t] = i + 1
        for lv, txt, hid, tok in heads:
            if lv != 1:
                continue
            r = real.get(txt.strip())
            if r is None:
                warn.append("目录项「%s」未能在正文定位（可能标题被拆行）" % txt[:24])

    # 4) 内容溢出纸张
    for i, p in enumerate(d):
        R = p.rect
        for blk in p.get_text("blocks"):
            x0, y0, x1, y1 = blk[:4]
            if x1 > R.width + 2 or y1 > R.height + 2 or x0 < -2:
                warn.append("p%d 有内容超出页面边界（宽表格？）" % (i + 1))
                break
    d.close()

    if not quiet:
        if err:
            print("  [自检] 发现 %d 项问题：" % len(err))
            for x in err:
                print("     x %s" % x)
        if warn:
            print("  [自检] %d 项提示：" % len(warn))
            for x in warn[:8]:
                print("     - %s" % x)
        if not err and not warn:
            print("  [自检] 通过")
    return err, warn


def convert(src, out=None, overrides=None, preview=None, keep=False, quiet=False):
    overrides = overrides or {}
    raw = io.open(src, encoding="utf-8-sig").read()
    cfg, body = split_front_matter(raw)
    cfg.update({k: v for k, v in overrides.items() if v is not None})

    h1, body = drop_leading_h1(body)
    title = cfg.get("title") or h1 or os.path.splitext(os.path.basename(src))[0]

    if cfg.get("strip_lead_quote", True):
        body = drop_leading_blockquote(body)
    if cfg.get("drop_manual_toc", True):
        body = drop_toc_section(body)
    body = body.strip("\n").lstrip("-").strip()

    depth = int(cfg.get("toc_depth", 1))
    heads = scan_headings(body, depth)
    body = inject_tokens(body, heads, depth)

    browser = find_browser(cfg.get("browser"))
    font = find_cjk_font(cfg.get("stamp_font"))
    if font is None and not quiet:
        print("! 未找到中文字体文件，页眉页脚将退回拉丁字体")

    work = tempfile.mkdtemp(prefix="m2pw_")
    out = out or os.path.splitext(src)[0] + ".pdf"
    try:
        # 多遍渲染，直到「用于生成的页码表」与「渲染结果的页码表」一致
        pages, pdf, total = None, None, 0
        for attempt in range(4):
            html = build_html(body, heads, cfg, title, pages)
            pdf = render_pdf(html, browser, work, "p%d" % attempt)
            new, total = locate_tokens(pdf, heads)
            if not quiet:
                print("  第 %d 遍：%d 页，定位 %d/%d 个标题"
                      % (attempt + 1, total, len(new), len(heads)))
            if new == pages:
                break
            pages = new
        pages = pages or {}
        postprocess(pdf, out, heads, pages, cfg, title, font)
    finally:
        if keep:
            print("  中间文件保留在：%s" % work)
        else:
            shutil.rmtree(work, ignore_errors=True)

    if preview:
        d = pymupdf.open(out)
        for p in preview:
            idx = max(0, min(d.page_count - 1, p - 1))
            png = "%s_p%d.png" % (os.path.splitext(out)[0], idx + 1)
            d[idx].get_pixmap(dpi=110).save(png)
            print("  预览图：%s" % png)
        d.close()

    if cfg.get("selfcheck", True):
        selfcheck(out, heads, cfg, quiet=quiet)

    kb = os.path.getsize(out) // 1024
    if not quiet:
        print("[OK] %s  （%d 页，%d KB）" % (out, total, kb))
    return out


def main():
    ap = argparse.ArgumentParser(
        description="Markdown → 排版 PDF（中文技术文档）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="配置写在 md 顶部的 YAML front-matter 里，命令行参数优先级更高。详见 README.md")
    ap.add_argument("src", help="源 Markdown 文件")
    ap.add_argument("-o", "--out", help="输出 PDF 路径（默认与源同名）")
    ap.add_argument("--theme", choices=sorted(THEMES), help="配色主题")
    ap.add_argument("--paper", choices=sorted(PAPERS), help="纸张")
    ap.add_argument("--landscape", action="store_true", default=None, help="横向")
    ap.add_argument("--toc-depth", type=int, choices=[1, 2], dest="toc_depth",
                    help="目录层级：1=只 ##，2=## + ###")
    ap.add_argument("--no-cover", action="store_const", const=False, dest="cover", help="不要封面")
    ap.add_argument("--no-toc", action="store_const", const=False, dest="toc", help="不要目录")
    ap.add_argument("--no-break", action="store_const", const=False, dest="break_before_h2",
                    help="## 标题不强制另起页（短文档用）")
    ap.add_argument("--font-size", type=float, dest="font_size", help="正文字号 pt")
    ap.add_argument("--browser", help="Chrome/Edge 可执行文件路径")
    ap.add_argument("--preview", type=int, nargs="*", metavar="N",
                    help="导出指定页的 PNG 便于检查，默认 1 2")
    ap.add_argument("--keep", action="store_true", help="保留中间 HTML/PDF")
    ap.add_argument("-q", "--quiet", action="store_true")
    a = ap.parse_args()

    if not os.path.exists(a.src):
        sys.exit("源文件不存在：%s" % a.src)

    ov = dict(theme=a.theme, paper=a.paper, landscape=a.landscape,
              toc_depth=a.toc_depth, cover=a.cover, toc=a.toc,
              break_before_h2=a.break_before_h2, font_size=a.font_size,
              browser=a.browser)
    pv = None
    if a.preview is not None:
        pv = a.preview or [1, 2]

    convert(a.src, a.out, ov, preview=pv, keep=a.keep, quiet=a.quiet)


if __name__ == "__main__":
    main()
