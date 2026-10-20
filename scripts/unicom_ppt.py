"""
unicom_ppt.py —— 基于中国联通官方母版生成「可编辑」PPTX 的工具库

    import sys; sys.path.insert(0, "<skill目录>/scripts")
    from unicom_ppt import *

    deck = UnicomDeck()
    deck.cover("2025年三季度经营分析", org="XX分公司 市场部", date="2025年10月")
    deck.toc(["经营概览", "问题诊断", "下一步举措"], active=0)
    pg = deck.page("一、经营概览")
    y = pg.lead("收入同比" + growth(12.6, 11.98) + "，但宽带渗透率较年初" + pp(41.2, 42.5))
    pg.bullets(L, y, W, 2, ["<b>收入：</b>…", "<b>用户：</b>…"])
    pg.source("数据来源：经分系统；统计口径：2025年1-9月")
    deck.ending()
    deck.save("out.pptx")

约定
- 坐标单位一律为英寸（页面 13.333 x 7.5）。内容安全区：L / T / W / H / R / B 常量。
- 所有文字参数都支持轻量标记：
    <b>加粗</b>  <good>好(红)</good>  <bad>差(绿)</bad>
    <red>强调红</red>  <green>绿</green>  <blue>蓝</blue>  <gray>灰</gray>
  换行用 "\n"，或直接传字符串列表（每项一个段落）。
- 文本框默认自动缩字号以避免溢出（不低于 min_size），仍溢出时在 save() 时给出警告。
"""
import math
import os
import re
import sys

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION, XL_TICK_MARK
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

# ---------------------------------------------------------------- 路径
_HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(os.path.dirname(_HERE), "assets")
TEMPLATE = os.path.join(ASSETS, "unicom_template.pptx")
LOGO = os.path.join(ASSETS, "unicom_logo.png")
ENDING_IMG = os.path.join(ASSETS, "ending.jpg")

# 母版中的版式名称（更换母版时只需改这里）
LAYOUT_COVER = "标题幻灯片"
LAYOUT_TOC = "1_联通版-分隔页1"
LAYOUT_CONTENT = "冬奥联通版范例一-空白"

# ---------------------------------------------------------------- 品牌规范
FONT = "微软雅黑"

RED = "C00000"         # 母版主红：标题条、项目符号、目录当前项；数据分析中表示「好」
GREEN = "008650"       # 数据分析中表示「差」
GOOD, BAD = RED, GREEN
COVER_RED = "EE0000"   # 仅封面竖线
BLACK = "000000"       # 页面标题
TEXT = "262626"        # 正文
DARK = "595959"        # 次级文字
GRAY = "7F7F7F"        # 注释、来源
MID = "A6A6A6"         # 目录非当前项、中性图表系列
LINE = "D9D9D9"        # 边框、分隔线
LIGHT = "F2F2F2"       # 浅灰底板
WARM = "F1EBE9"        # 母版暖灰（分隔页底色）
LIGHT_RED = "FBE9E9"   # 浅红底板
PINK = "F4C7C7"        # 浅红强调/分层
BLUE = "2F5496"        # 辅助说明、策略信息（传统报告类）
LIGHT_BLUE = "E3EAF5"
WHITE = "FFFFFF"

SERIES = [RED, MID, DARK, PINK, BLUE]   # 多系列图表默认配色

TAG_STYLE = {
    "b": {"bold": True},
    "good": {"color": GOOD, "bold": True},
    "bad": {"color": BAD, "bold": True},
    "red": {"color": RED, "bold": True},
    "green": {"color": GREEN, "bold": True},
    "blue": {"color": BLUE},
    "gray": {"color": GRAY},
}

# ---------------------------------------------------------------- 版面网格（英寸）
SLIDE_W, SLIDE_H = 13.333, 7.5
L = 0.30            # 内容左边界
R = 13.03           # 内容右边界
T = 1.05            # 内容上边界（红色标题线下方）
B = 6.85            # 内容下边界（页码/来源上方）
W = R - L
H = B - T

EMU = 914400
_LH = 1.30          # 微软雅黑单倍行距 ≈ 1.3 × 字号


def _e(v):
    return Emu(int(round(v * EMU)))


def _rgb(hexstr):
    return RGBColor.from_string(hexstr.replace("#", "").upper())


# ================================================================ 数字格式化（数据分析类核心规则）
def mark(text, good):
    """按好坏给文本上色：good=True → 红，False → 绿，None → 仅加粗。"""
    if good is None:
        return f"<b>{text}</b>"
    tag = "good" if good else "bad"
    return f"<{tag}>{text}</{tag}>"


_BETTER = ("up", "down", None)


def num(v, unit="", digits=1, sign=False):
    """纯文本数字（不着色）。负号用数学减号 −；sign=True 时正数带 +。
    num(-1.25, "pp", sign=True) → '−1.3pp'；num(38.4, "%") → '38.4%'"""
    r = round(v, digits)
    if r == 0:
        r = 0.0                                          # 避免出现 -0.0
    s = "+" if (sign and r > 0) else ("\u2212" if r < 0 else "")   # 数学减号 −，PowerPoint 不会在其后断行
    return f"{s}{abs(r):.{digits}f}{unit}"


def delta(v, unit="%", better="up", digits=1, prefix=""):
    """格式化一个「变化量」并按好坏着色（返回带标记的文本）。

    v       变化值，例如 5.23 表示 +5.23%（unit='%'）或 +5.23pp（unit='pp'）
    better  'up' 越高越好；'down' 越低越好（离网率、成本、投诉、拆机）；None 不判断好坏，仅加粗
    比率类指标（渗透率、占比、转化率、完成率、离网率、份额、增速）之间的差值一律用 pp。
    四舍五入后为 0 时输出不带符号的「0.0」并仅加粗（持平不判好坏）。
    """
    if better not in _BETTER:
        raise ValueError(f"better 只能是 'up' / 'down' / None，收到 {better!r}")
    if v is None:
        return "—"
    r = round(v, digits)
    s = prefix + num(v, unit, digits, sign=True)
    if better is None or r == 0:
        return f"<b>{s}</b>"
    return mark(s, (r > 0) == (better == "up"))


def pp(new, old, better="up", digits=1, prefix=""):
    """两个百分比之差，单位 pp。pp(35.4, 32.1) → '<good>+3.3pp</good>'"""
    return delta(new - old, "pp", better, digits, prefix)


def growth(new, old, better="up", digits=1, prefix=""):
    """绝对量的变化率，单位 %。growth(12.6, 11.98) → '<good>+5.2%</good>'"""
    if not old:
        return "—"
    return delta((new - old) / abs(old) * 100, "%", better, digits, prefix)


def vs_target(actual, target, text=None, better="up", digits=1, unit=""):
    """给「实际值本身」按是否达到目标/时序进度着色：达到或优于 → 红，落后 → 绿。
    vs_target(75.2, 75, unit="%") → '<good>75.2%</good>'
    （要展示「差距」时用 pp(actual, target)；两者可同时使用。）"""
    if better not in ("up", "down"):
        raise ValueError(f"better 只能是 'up' / 'down'，收到 {better!r}")
    good = actual >= target if better == "up" else actual <= target
    return mark(text if text is not None else num(actual, unit, digits), good)


# ================================================================ 标记解析
_TOKEN = re.compile(r"(</?(?:%s)>)" % "|".join(TAG_STYLE))


def parse_markup(text):
    """'<b>A</b>B' → [('A', {'bold':True}), ('B', {})]"""
    runs, stack = [], []
    for tok in _TOKEN.split(text):
        if not tok:
            continue
        m = re.fullmatch(r"<(/?)(\w+)>", tok)
        if m and m.group(2) in TAG_STYLE:
            if m.group(1):
                for i in range(len(stack) - 1, -1, -1):
                    if stack[i] == m.group(2):
                        del stack[i]
                        break
            else:
                stack.append(m.group(2))
            continue
        style = {}
        for t in stack:
            style.update(TAG_STYLE[t])
        runs.append((tok, style))
    return runs


def strip_markup(text):
    return _TOKEN.sub("", text)


# ================================================================ 文字宽度估算（用于自动缩字号与溢出检查）
def _char_em(ch):
    o = ord(ch)
    if o >= 0x2E80 or 0xFF00 <= o <= 0xFFEF or ch in "“”‘’—…·":
        return 1.0
    if ch == " ":
        return 0.3
    if ch in "il.,:;|!'`()[]{}":
        return 0.32
    if ch in "mwMW%@":
        return 0.9
    if ch.isupper() or ch.isdigit() or ch in "+-=#$&?":
        return 0.6
    return 0.52


def _tokens(text):
    """CJK 单字成词，西文连续字符成词。"""
    return re.findall(r"[A-Za-z0-9.,%+\-_/:'()]+|\s|.", text)


def _layout(runs, avail_in, safety=0.96):
    """runs: [(text, size_pt)] → [(行高pt, 该行文字)]。safety<1 时预留渲染差异余量。"""
    avail_in *= safety
    lines, cur_w, cur_max, cur_t = [], 0.0, 0.0, ""
    for text, size in runs:
        for tok in _tokens(text):
            if tok == "\n":
                lines.append((cur_max or size, cur_t))
                cur_w, cur_max, cur_t = 0.0, 0.0, ""
                continue
            w = sum(_char_em(c) for c in tok) * size / 72.0
            if cur_w + w > avail_in and cur_w > 0 and not tok.isspace():
                lines.append((cur_max or size, cur_t))
                cur_w, cur_max, cur_t = 0.0, 0.0, ""
            while w > avail_in and avail_in > 0:       # 超长词强制断开
                lines.append((size, tok))
                w -= avail_in
            cur_w += w
            cur_max = max(cur_max, size)
            cur_t += tok
    lines.append((cur_max or (runs[-1][1] if runs else 12), cur_t))
    return lines


def _count_lines(runs, avail_in):
    return [h for h, _ in _layout(runs, avail_in)]


_PUNCT = set("，。、；：！？）》」』”’,.;:!?)%")


def _para_runs(p):
    runs = []
    for child in p:
        if child.tag in (qn("a:r"), qn("a:fld")):
            rPr = child.find(qn("a:rPr"))
            sz = rPr.get("sz") if rPr is not None else None
            t = child.find(qn("a:t"))
            runs.append((t.text if t is not None and t.text else "", int(sz) / 100 if sz else 18))
        elif child.tag == qn("a:br"):
            runs.append(("\n", runs[-1][1] if runs else 12))
    return runs


def find_orphans(txBody, width_emu):
    """找出末行只剩 1–2 个字的段落（中文排版中的「孤字」），返回段落开头文字列表。估算结果，仅作提示。"""
    bodyPr = txBody.find(qn("a:bodyPr"))
    lIns = _attr_emu(bodyPr, "lIns", 91440)
    rIns = _attr_emu(bodyPr, "rIns", 91440)
    out = []
    for p in txBody.findall(qn("a:p")):
        runs = _para_runs(p)
        if not runs:
            continue
        pPr = p.find(qn("a:pPr"))
        avail = (width_emu - lIns - rIns - _attr_emu(pPr, "marL", 0)) / EMU
        lines = _layout(runs, avail, safety=1.0)
        if len(lines) >= 2 and "\n" not in "".join(t for t, _ in runs):
            last = [c for c in lines[-1][1].strip() if c not in _PUNCT]
            if 0 < len(last) <= 2:
                out.append("".join(t for t, _ in runs)[:14])
    return out


def _attr_emu(el, name, default):
    v = el.get(name) if el is not None else None
    return int(v) if v is not None else default


def estimate_tf_height(txBody, width_emu):
    """估算一个 txBody（文本框或表格单元格）所需高度（英寸）。"""
    bodyPr = txBody.find(qn("a:bodyPr"))
    lIns = _attr_emu(bodyPr, "lIns", 91440)
    rIns = _attr_emu(bodyPr, "rIns", 91440)
    tIns = _attr_emu(bodyPr, "tIns", 45720)
    bIns = _attr_emu(bodyPr, "bIns", 45720)
    total = (tIns + bIns) / EMU
    for p in txBody.findall(qn("a:p")):
        pPr = p.find(qn("a:pPr"))
        marL = _attr_emu(pPr, "marL", 0)
        avail = (width_emu - lIns - rIns - marL) / EMU
        runs = _para_runs(p)
        if not runs:
            end = p.find(qn("a:endParaRPr"))
            sz = end.get("sz") if end is not None else None
            runs = [("", int(sz) / 100 if sz else 18)]
        factor, fixed = 1.0, None
        if pPr is not None:
            ln = pPr.find(qn("a:lnSpc"))
            if ln is not None:
                pct = ln.find(qn("a:spcPct"))
                pts = ln.find(qn("a:spcPts"))
                if pct is not None:
                    factor = int(pct.get("val")) / 100000
                elif pts is not None:
                    fixed = int(pts.get("val")) / 100
        for lh in _count_lines(runs, avail):
            total += (fixed if fixed else lh * _LH * factor) / 72
        if pPr is not None:
            for tag in ("a:spcBef", "a:spcAft"):
                sp = pPr.find(qn(tag))
                if sp is not None and sp.find(qn("a:spcPts")) is not None:
                    total += int(sp.find(qn("a:spcPts")).get("val")) / 100 / 72
    return total


# ================================================================ 底层 XML 工具
def _strip_style(shape):
    """去掉 python-pptx 自动形状自带的主题样式（避免阴影/白字/蓝边）。"""
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def _shadow(shape):
    spPr = shape._element.spPr
    eff = etree.SubElement(spPr, qn("a:effectLst"))
    sh = etree.SubElement(eff, qn("a:outerShdw"), blurRad="50800", dist="38100",
                          dir="8100000", algn="tr", rotWithShape="0")
    clr = etree.SubElement(sh, qn("a:prstClr"), val="black")
    etree.SubElement(clr, qn("a:alpha"), val="40000")


def _set_run(run, text, size, color, bold, font=FONT):
    run.text = text
    rPr = run._r.get_or_add_rPr()
    for child in list(rPr):
        rPr.remove(child)
    rPr.set("lang", "zh-CN")
    rPr.set("altLang", "en-US")
    rPr.set("sz", str(int(round(size * 100))))
    rPr.set("b", "1" if bold else "0")
    fill = etree.SubElement(rPr, qn("a:solidFill"))
    etree.SubElement(fill, qn("a:srgbClr"), val=color)
    for tag in ("a:latin", "a:ea", "a:cs"):
        etree.SubElement(rPr, qn(tag), typeface=font)


def _bullet(p, char="p", font="Wingdings", color=RED, marL=0.22, indent=-0.22):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(marL * EMU)))
    pPr.set("indent", str(int(indent * EMU)))
    etree.SubElement(pPr, qn("a:buClr")).append(etree.Element(qn("a:srgbClr"), val=color))
    etree.SubElement(pPr, qn("a:buSzPct"), val="100000")
    etree.SubElement(pPr, qn("a:buFont"), typeface=font, pitchFamily="2", charset="2")
    etree.SubElement(pPr, qn("a:buChar"), char=char)


def _no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    etree.SubElement(pPr, qn("a:buNone"))


_ALIGNS = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT, "j": PP_ALIGN.JUSTIFY}
_ANCHORS = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}
_ALIAS = {"left": "l", "center": "c", "centre": "c", "middle": "m", "right": "r", "justify": "j",
          "top": "t", "bottom": "b"}


def _align(a):
    k = _ALIAS.get(str(a).lower(), str(a).lower())
    if k not in _ALIGNS:
        raise ValueError(f"align 只能是 l/c/r/j（或 left/center/right/justify），收到 {a!r}")
    return _ALIGNS[k]


def _anchor(a):
    k = _ALIAS.get(str(a).lower(), str(a).lower())
    k = {"c": "m"}.get(k, k)
    if k not in _ANCHORS:
        raise ValueError(f"anchor 只能是 t/m/b（或 top/middle/bottom），收到 {a!r}")
    return _ANCHORS[k]


def _paras(text):
    if isinstance(text, (list, tuple)):
        out = []
        for t in text:
            out.extend(str(t).split("\n"))
        return out
    return str(text).split("\n")


def write_text(tf, text, size=12, color=TEXT, bold=False, align="l", line_spacing=1.15,
               space_after=0, bullets=None, sizes=None):
    """把（带标记的）文本写入 text_frame。

    bullets: None 无符号；'square' 红色□（母版样式）；'dot' 灰色•；也可传列表逐段指定。
             以两个空格或 '- ' 开头的段落自动作为二级（灰色•、缩进更深）。
    sizes:   可选，逐段字号列表（用于 KPI 卡片等混排）。
    """
    paras = _paras(text)
    tf.clear()
    for i, raw in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        level2 = raw.startswith("  ") or raw.startswith("- ")
        line = raw[2:] if level2 else raw
        p.alignment = _align(align)
        p.line_spacing = line_spacing
        if space_after:
            p.space_after = Pt(space_after)
        b = bullets[i] if isinstance(bullets, (list, tuple)) else bullets
        if b and level2:
            _bullet(p, char="•", font="Arial", color=DARK, marL=0.42, indent=-0.16)
        elif b == "square":
            _bullet(p)
        elif b == "dot":
            _bullet(p, char="•", font="Arial", color=DARK, marL=0.18, indent=-0.18)
        else:
            _no_bullet(p)
        sz = sizes[i] if sizes else (size - 1 if (b and level2) else size)
        runs = parse_markup(line) or [("", {})]
        for txt, st in runs:
            _set_run(p.add_run(), txt, sz, st.get("color", color), st.get("bold", bold))
        end = p._p.get_or_add_endParaRPr()
        end.set("sz", str(int(sz * 100)))


def _frame(tf, margin=0.06, anchor="t", wrap=True):
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    if isinstance(margin, (int, float)):
        margin = (margin, margin * 0.6, margin, margin * 0.6)
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = (_e(m) for m in margin)
    tf.vertical_anchor = _anchor(anchor)


def _shrink_to_fit(shape, min_size, h_in=None):
    """按估算结果逐步缩小字号直到放得下；返回 (是否放得下, 所需高度, 缩小的磅数)。"""
    txBody = shape._element.find(qn("p:txBody"))
    avail = h_in if h_in is not None else shape.height / EMU
    need = estimate_tf_height(txBody, shape.width)
    shrunk = 0.0
    while need > avail + 0.01:
        rprs = txBody.findall(".//" + qn("a:rPr")) + txBody.findall(".//" + qn("a:endParaRPr"))
        sizes = [int(r.get("sz")) for r in rprs if r.get("sz")]
        if not sizes or max(sizes) - 50 < min_size * 100:
            return False, need, shrunk
        for r in rprs:
            if r.get("sz"):
                r.set("sz", str(max(int(r.get("sz")) - 50, 600)))
        shrunk += 0.5
        need = estimate_tf_height(txBody, shape.width)
    return True, need, shrunk


def _set_cell_border(cell, color=LINE, width=0.75, sides="LRTB"):
    tcPr = cell._tc.get_or_add_tcPr()
    fills = [c for c in tcPr if c.tag in (qn("a:solidFill"), qn("a:noFill"))]
    for c in list(tcPr):
        if c.tag in (qn("a:lnL"), qn("a:lnR"), qn("a:lnT"), qn("a:lnB")):
            tcPr.remove(c)
    for f in fills:
        tcPr.remove(f)
    for s in "LRTB":
        ln = etree.SubElement(tcPr, qn("a:ln" + s), w=str(int(width * 12700)), cap="flat", cmpd="sng", algn="ctr")
        if s in sides:
            sf = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr"), val=color)
            etree.SubElement(ln, qn("a:prstDash"), val="solid")
        else:
            etree.SubElement(ln, qn("a:noFill"))
    for f in fills:
        tcPr.append(f)


def _chart_font(chart, size, color=TEXT):
    chart.font.size = Pt(size)
    chart.font.name = FONT
    chart.font.color.rgb = _rgb(color)
    defRPr = chart._chartSpace.find(qn("c:txPr")).find(".//" + qn("a:defRPr"))
    if defRPr.find(qn("a:ea")) is None:
        etree.SubElement(defRPr, qn("a:ea"), typeface=FONT)


# ================================================================ 页面
class Page:
    """一页幻灯片。所有绘图方法坐标单位为英寸，返回创建的 shape（或下一个可用 y）。"""

    def __init__(self, deck, slide, kind):
        self.deck = deck
        self.slide = slide
        self.kind = kind
        self.shapes = slide.shapes
        self.top = T            # 下一个可用 y（lead() 后自动下移）

    # ---------- 基础
    def _warn(self, msg):
        self.deck.warnings.append(f"第{self.index}页：{msg}")

    @property
    def index(self):
        return list(self.deck.prs.slides).index(self.slide) + 1

    def text(self, x, y, w, h, text, size=12, color=TEXT, bold=False, align="l", anchor="t",
             line_spacing=1.15, space_after=0, bullets=None, fill=None, line=None, margin=0.05,
             fit=True, min_size=10, sizes=None):
        """文本框（支持标记、项目符号、自动缩字号）。sizes 可逐段指定字号。"""
        shp = self.shapes.add_textbox(_e(x), _e(y), _e(w), _e(h))
        _frame(shp.text_frame, margin, anchor)
        write_text(shp.text_frame, text, size, color, bold, align, line_spacing, space_after, bullets, sizes)
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = _rgb(fill)
        if line:
            shp.line.color.rgb = _rgb(line)
            shp.line.width = Pt(0.75)
        if fit:
            ok, need, shrunk = _shrink_to_fit(shp, min_size)
            self._fit_report(ok, need, h, shrunk, size, text)
        return shp

    def _fit_report(self, ok, need, h, shrunk, size, text):
        head = strip_markup(_paras(text)[0])[:18]
        if not ok:
            self._warn(f"文字可能溢出（需约 {need:.2f}in，框高 {h:.2f}in）：{head}…")
        elif shrunk >= 1:
            self._warn(f"字号已从 {size:g} 自动缩到 {size - shrunk:g}，可能与同组文字不一致"
                       f"（减字、加大框，或给同组统一设字号）：{head}…")

    def bullets(self, x, y, w, h, items, size=13, color=TEXT, line_spacing=1.25, space_after=4,
                style="square", **kw):
        """项目符号列表（母版样式：红色□）。以 '- ' 开头的项为二级。"""
        return self.text(x, y, w, h, items, size=size, color=color, line_spacing=line_spacing,
                         space_after=space_after, bullets=style, **kw)

    def box(self, x, y, w, h, text="", fill=None, line=LINE, line_w=0.75, color=TEXT, size=12,
            bold=False, align="l", anchor="t", shape="rect", dash=False, margin=0.08,
            line_spacing=1.15, space_after=0, bullets=None, fit=True, min_size=9, shadow=False):
        """矩形/圆角矩形容器，可带文字。shape: rect | round | pill"""
        kind = {"rect": MSO_SHAPE.RECTANGLE, "round": MSO_SHAPE.ROUNDED_RECTANGLE,
                "pill": MSO_SHAPE.ROUNDED_RECTANGLE}[shape]
        shp = self.shapes.add_shape(kind, _e(x), _e(y), _e(w), _e(h))
        _strip_style(shp)
        if shape == "round":
            shp.adjustments[0] = min(0.12, 0.08 / max(min(w, h), 0.1))
        elif shape == "pill":
            shp.adjustments[0] = 0.5
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = _rgb(fill)
        else:
            shp.fill.background()
        if line:
            shp.line.color.rgb = _rgb(line)
            shp.line.width = Pt(line_w)
            if dash:
                shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        else:
            shp.line.fill.background()
        if shadow:
            _shadow(shp)
        _frame(shp.text_frame, margin, anchor)
        if text != "":
            write_text(shp.text_frame, text, size, color, bold, align, line_spacing, space_after, bullets)
            if fit:
                ok, need, shrunk = _shrink_to_fit(shp, min_size)
                self._fit_report(ok, need, h, shrunk, size, text)
        return shp

    def label(self, x, y, w, h, text, fill=RED, color=WHITE, size=13, bold=True, align="c", shape="rect", **kw):
        """色块标题条（如：体系名称、方案名称、栏目名）。"""
        return self.box(x, y, w, h, text, fill=fill, line=None, color=color, size=size, bold=bold,
                        align=align, anchor="m", shape=shape, margin=0.06, **kw)

    def heading(self, x, y, w, text, size=15, color=BLACK, rule=False):
        """模块小标题：母版样式红色□ + 加粗文字。返回下方可用 y。"""
        h = size * _LH / 72 + 0.1
        self.text(x, y, w, h, text, size=size, color=color, bold=True, bullets="square", anchor="m", fit=False)
        if rule:
            self.line(x, y + h + 0.02, x + w, y + h + 0.02, color=LINE, width=0.75)
        return y + h + (0.1 if rule else 0.06)

    def lead(self, text, x=L, y=None, w=W, size=15, bold=True, fill=LIGHT, h=None):
        """页首结论条（金字塔塔尖：一句话结论，关键数字着色）。返回下方内容起始 y。"""
        y = self.top if y is None else y
        if h is None:
            probe = self.shapes.add_textbox(_e(x + 0.12), 0, _e(w - 0.12), _e(1))
            _frame(probe.text_frame, (0.12, 0.08, 0.12, 0.08), "m")
            write_text(probe.text_frame, text, size, TEXT, bold, "l", 1.2)
            need = estimate_tf_height(probe._element.find(qn("p:txBody")), probe.width)
            probe._element.getparent().remove(probe._element)
            h = max(0.52, need + 0.04)
        self.box(x, y, w, h, fill=fill, line=None)
        self.box(x, y, 0.07, h, fill=RED, line=None)
        self.text(x + 0.12, y, w - 0.12, h, text, size=size, bold=bold, anchor="m",
                  line_spacing=1.2, margin=(0.12, 0.08, 0.12, 0.08), min_size=12)
        self.top = y + h + 0.18
        return self.top

    def kpi(self, x, y, w, h, label, value, unit="", change=None, note=None, judge=None,
            value_size=28, fill=LIGHT, align="l"):
        """指标卡：标签 / 大数字+单位 / 变化（标记文本，如 '同比 ' + growth(a, b)） / 备注。
        value: 纯数字文本（建议用 num()），也可直接传 growth()/pp() 等带标记的文本。
        unit:  '%'、'pp' 紧贴数字且同字号；其他单位（亿元、万户）以小字号跟在后面。
        judge: 'good' | 'bad' | None —— 大数字本身是否按好坏着色。
               只有当大数字本身就是评价（完成率、达标率、排名）时才设 judge；水平值保持 None。"""
        if judge not in ("good", "bad", None):
            raise ValueError(f"judge 只能是 'good' / 'bad' / None，收到 {judge!r}")
        shp = self.box(x, y, w, h, fill=fill, line=None)
        tf = shp.text_frame
        _frame(tf, (0.14, 0.1, 0.14, 0.08), "m")
        tf.clear()
        vcolor = {"good": GOOD, "bad": BAD}.get(judge, BLACK)
        rows = ([("label", label)] if label else []) + [("value", value)]
        if change:
            rows.append(("change", change))
        if note:
            rows.append(("note", note))
        for i, (kind, content) in enumerate(rows):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = _align(align)
            p.line_spacing = 1.0
            _no_bullet(p)
            if kind == "label":
                _set_run(p.add_run(), content, 12, DARK, True)
                p.space_after = Pt(2)
            elif kind == "value":
                for txt, st in parse_markup(str(content)):
                    _set_run(p.add_run(), txt, value_size, st.get("color", vcolor), True)
                if unit:
                    tight = unit in ("%", "pp", "‰")
                    _set_run(p.add_run(), unit if tight else " " + unit, value_size if tight else 12,
                             vcolor if tight else DARK, tight)
                p.space_after = Pt(2)
            else:
                sz = 11 if kind == "change" else 10
                for txt, st in parse_markup(content):
                    _set_run(p.add_run(), txt, sz, st.get("color", DARK if kind == "change" else GRAY),
                             st.get("bold", False))
        ok, need, _ = _shrink_to_fit(shp, 9)
        if not ok:
            self._warn(f"指标卡「{label}」内容放不下（需约 {need:.2f}in）")
        return shp

    def conclusions(self, items, x=L, y=None, w=W, h=None, gap=0.18, value_w=2.4, head_size=16, body_size=13):
        """核心结论 / 关键要点列表（金字塔第二层）。每行 = 序号徽标 + 判断标题 + 证据句 [+ 右侧关键数字]。

        items: [(标题, 证据, 数字, 数字标签, judge), ...]，后三项可省略或为 None。
               证据、数字均可传带标记文本；数字建议用 num()/growth()/pp()。
        默认占满 lead() 之后到页面底部的空间。返回下方可用 y。"""
        y = self.top if y is None else y
        h = (B - y) if h is None else h
        rows = self.rows(len(items), y=y, h=h, gap=gap)
        for i, ((ry, rh), it) in enumerate(zip(rows, items)):
            head, body, value, lab, judge = (list(it) + [None] * 5)[:5]
            self.box(x, ry, w, rh, fill=None, line=LINE)
            d = min(0.44, rh - 0.2)
            self.badge(x + 0.22, ry + rh / 2 - d / 2, d, i + 1, size=13)
            tx = x + 0.85
            tw = (w - 0.85 - value_w - 0.35) if value is not None else (w - 1.05)
            paras = [f"<b>{head}</b>"] + ([body] if body else [])
            self.text(tx, ry + 0.06, tw, rh - 0.12, paras, size=body_size, color=TEXT, anchor="m",
                      line_spacing=1.25, space_after=4, sizes=[head_size] + [body_size] * (len(paras) - 1))
            if value is not None:
                vx = x + w - value_w - 0.15
                self.line(vx - 0.15, ry + 0.2, vx - 0.15, ry + rh - 0.2, color=LINE)
                self.kpi(vx, ry + 0.1, value_w, rh - 0.2, lab or "", value, judge=judge, fill=WHITE,
                         align="c", value_size=28)
        return y + h

    def badge(self, x, y, d, text, fill=RED, color=WHITE, size=12):
        """圆形序号徽标。"""
        shp = self.shapes.add_shape(MSO_SHAPE.OVAL, _e(x), _e(y), _e(d), _e(d))
        _strip_style(shp)
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(fill)
        shp.line.fill.background()
        _frame(shp.text_frame, 0, "m", wrap=False)
        write_text(shp.text_frame, str(text), size, color, True, "c", 1.0)
        return shp

    def line(self, x1, y1, x2, y2, color=MID, width=1.0, dash=False, head=None, tail=None):
        """直线/箭头。head/tail: None | 'triangle' | 'arrow' | 'oval'（tail 为终点端）。"""
        c = self.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, _e(x1), _e(y1), _e(x2), _e(y2))
        _strip_style(c)
        c.line.color.rgb = _rgb(color)
        c.line.width = Pt(width)
        if dash:
            c.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        ln = c.line._get_or_add_ln()
        if head:
            etree.SubElement(ln, qn("a:headEnd"), type=head, w="med", len="med")
        if tail:
            etree.SubElement(ln, qn("a:tailEnd"), type=tail, w="med", len="med")
        return c

    def arrow(self, x1, y1, x2, y2, color=MID, width=1.25, dash=False, both=False, text=None,
              text_size=10, text_color=DARK, text_pos="side"):
        """带箭头连线（终点有箭头；both=True 双向）。起止点应正好落在框边上。
        text 为连线标签（数据交换 / 服务支撑 / 业务协同 / 结果反馈）。
        text_pos='side'：竖线标签放在线右侧、横线标签放在线上方，不遮挡箭头；'center'：白底压在线中间。"""
        c = self.line(x1, y1, x2, y2, color, width, dash, head="triangle" if both else None, tail="triangle")
        if text:
            tw = max(0.5, sum(_char_em(ch) for ch in strip_markup(text)) * text_size / 72 + 0.12)
            th = text_size * _LH / 72 + 0.06
            mx, my = (x1 + x2) / 2, (y1 + y2) / 2
            vertical = abs(x2 - x1) < abs(y2 - y1)
            if text_pos == "side" and vertical:
                tx, ty, al = mx + 0.06, my - th / 2, "l"
            elif text_pos == "side":
                tx, ty, al = mx - tw / 2, min(y1, y2) - th - 0.01, "c"
            else:
                tx, ty, al = mx - tw / 2, my - th / 2, "c"
            self.text(tx, ty, tw, th, text, size=text_size, color=text_color,
                      align=al, anchor="m", fill=None if text_pos == "side" else WHITE, margin=0.02, fit=False)
        return c

    def loop(self, src, dst, top, text=None, depth=0.3, color=RED, width=1.25, text_size=11):
        """反馈闭环：从 src 列底部引出虚线，沿下方横线回到 dst 列底部（带箭头）。
        src / dst 为 (x, w) 列（如 chevrons() 的返回值）；top 为两列内容的底边 y；depth 为横线离 top 的距离。
        只有材料里确有反馈机制时才画；dst 指向「被反馈优化的对象」（策略、模型或规则所在的列）。
        返回横线下方可用的 y。"""
        xs, xd = src[0] + src[1] / 2, dst[0] + dst[1] / 2
        ly = top + depth
        self.line(xs, top, xs, ly, color, width, dash=True)
        self.line(xs, ly, xd, ly, color, width, dash=True)
        self.line(xd, ly, xd, top, color, width, dash=True, tail="triangle")
        if text:
            tw = sum(_char_em(ch) for ch in strip_markup(text)) * text_size / 72 + 0.3
            th = text_size * _LH / 72 + 0.08
            self.text((xs + xd) / 2 - tw / 2, ly - th / 2, tw, th, text, size=text_size, color=color, bold=True,
                      align="c", anchor="m", fill=WHITE, margin=0.02, fit=False)
        return ly + 0.2

    def block_arrow(self, x, y, w, h, direction="right", fill=PINK, text="", color=TEXT, size=11):
        """粗箭头（阶段推进、因果指向）。direction: right | down | left | up"""
        kind = {"right": MSO_SHAPE.RIGHT_ARROW, "down": MSO_SHAPE.DOWN_ARROW,
                "left": MSO_SHAPE.LEFT_ARROW, "up": MSO_SHAPE.UP_ARROW}[direction]
        shp = self.shapes.add_shape(kind, _e(x), _e(y), _e(w), _e(h))
        _strip_style(shp)
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(fill)
        shp.line.fill.background()
        if text:
            _frame(shp.text_frame, 0.02, "m")
            write_text(shp.text_frame, text, size, color, True, "c", 1.0)
        return shp

    def chevrons(self, x, y, w, h, steps, fill=RED, color=WHITE, size=14, gap=0.25, active=None,
                 inactive=MID, ratios=None):
        """横向阶段链（首个为五边形，其余为燕尾），节点与等宽内容列居中对齐。
        active=None 时全部用 fill，否则仅 active 节点用 fill。
        ratios 可指定各列宽度比例（如 [1, 1, 1.4, 1]），默认等宽。
        返回各内容列 (x, w)，可直接用于在节点下方排布对应内容。"""
        n = len(steps)
        notch = min(h * 0.35, 0.3)
        if ratios is not None and len(ratios) != n:
            raise ValueError(f"ratios 长度（{len(ratios)}）应与 steps 数量（{n}）一致")
        columns = self.cols(ratios or n, x, w, gap)
        for i, (s, (cx, cw)) in enumerate(zip(steps, columns)):
            left = cx if i == 0 else cx - gap / 2 - notch / 2 + 0.03
            right = cx + cw if i == n - 1 else cx + cw + gap / 2 + notch / 2 - 0.03
            kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
            shp = self.shapes.add_shape(kind, _e(left), _e(y), _e(right - left), _e(h))
            _strip_style(shp)
            shp.adjustments[0] = notch / h
            shp.fill.solid()
            shp.fill.fore_color.rgb = _rgb(fill if active is None or i == active else inactive)
            shp.line.fill.background()
            _frame(shp.text_frame, (notch + 0.04 if i else 0.08, 0.02, notch + 0.04, 0.02), "m")
            write_text(shp.text_frame, s, size, color, True, "c", 1.0)
            _shrink_to_fit(shp, 9)
        return columns

    def table(self, x, y, w, h, rows, col_widths=None, header=True, size=11, header_size=None,
              header_fill=RED, header_color=WHITE, zebra=True, align=None, first_col_bold=False,
              highlight_rows=None, highlight_fill=LIGHT_RED):
        """原生表格。rows 为二维列表（单元格支持标记）；align 逐列 'l'/'c'/'r'。
        行高按 h 平均分配，PowerPoint 会按内容自动撑高——内容过多时 save() 会给出警告。"""
        nr, nc = len(rows), max(len(r) for r in rows)
        gf = self.shapes.add_table(nr, nc, _e(x), _e(y), _e(w), _e(h))
        tbl = gf.table
        tblPr = tbl._tbl.tblPr
        tblPr.set("firstRow", "0")
        tblPr.set("bandRow", "0")
        sid = tblPr.find(qn("a:tableStyleId"))
        if sid is None:
            sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"   # No Style, No Grid
        ws = col_widths or [1] * nc
        tot = sum(ws)
        for i, cw in enumerate(ws):
            tbl.columns[i].width = _e(w * cw / tot)
        for r in range(nr):
            tbl.rows[r].height = _e(h / nr)
        align = align or (["l"] + ["c"] * (nc - 1))
        highlight_rows = set(highlight_rows or [])
        for r in range(nr):
            for c in range(nc):
                cell = tbl.cell(r, c)
                val = rows[r][c] if c < len(rows[r]) else ""
                is_head = header and r == 0
                cell.margin_left = cell.margin_right = _e(0.06)
                cell.margin_top = cell.margin_bottom = _e(0.025)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                if is_head:
                    fillc = header_fill
                elif r in highlight_rows:
                    fillc = highlight_fill
                elif zebra and (r - (1 if header else 0)) % 2 == 1:
                    fillc = "F7F7F7"
                else:
                    fillc = WHITE
                cell.fill.solid()
                cell.fill.fore_color.rgb = _rgb(fillc)
                write_text(cell.text_frame, str(val), header_size or size if is_head else size,
                           header_color if is_head else TEXT,
                           bold=is_head or (first_col_bold and c == 0),
                           align="c" if is_head else align[c], line_spacing=1.0)
                _set_cell_border(cell, LINE, 0.75)
        need = sum(self._row_need(tbl, r) for r in range(nr))
        if need > h + 0.05:
            self._warn(f"表格内容需约 {need:.2f}in 高，超过预留 {h:.2f}in，行会被撑高（底部到 {y + need:.2f}in）")
        return gf

    @staticmethod
    def _row_need(tbl, r):
        best = 0
        for c, cell in enumerate(tbl.rows[r].cells):
            best = max(best, estimate_tf_height(cell._tc.find(qn("a:txBody")), tbl.columns[c].width)
                       - 0.1 + 0.05)   # 单元格边距与文本框不同，粗略修正
        return max(best, tbl.rows[r].height / EMU)

    # ---------- 图表（原生、可编辑）
    def _series(self, series):
        if isinstance(series, dict):
            return list(series.items())
        if series and not isinstance(series[0], (list, tuple)):
            return [("系列1", series)]
        return list(series)

    def bar(self, x, y, w, h, categories, series, colors=None, highlight=None, number_format="0.0",
            horizontal=False, stacked=False, labels=True, legend=None, size=10, gap_width=70,
            overlap=None, axis=False, label_color=TEXT):
        """柱状/条形图。
        series:    [1,2,3] | {'本期':[...], '同期':[...]} | [('本期',[...]), ...]
        highlight: {类别索引或名称: 'good'|'bad'|'neutral'|色值}，仅作用于第一个系列
                   （单系列默认浅灰；'neutral' 用深灰表示「重要但不评好坏」，如目标值、测算值）
        """
        ser = self._series(series)
        for k in (highlight or {}):                      # 先校验，避免报错时页面上留下半成品图表
            if not isinstance(k, int) and k not in categories:
                raise ValueError(f"highlight 中的类别 {k!r} 不在 categories {list(categories)} 中")
        cd = CategoryChartData()
        cd.categories = categories
        for name, vals in ser:
            cd.add_series(name, vals)
        if horizontal:
            ct = XL_CHART_TYPE.BAR_STACKED if stacked else XL_CHART_TYPE.BAR_CLUSTERED
        else:
            ct = XL_CHART_TYPE.COLUMN_STACKED if stacked else XL_CHART_TYPE.COLUMN_CLUSTERED
        gf = self.shapes.add_chart(ct, _e(x), _e(y), _e(w), _e(h), cd)
        ch = gf.chart
        ch.has_title = False
        _chart_font(ch, size)
        colors = colors or ([MID] if len(ser) == 1 else SERIES)
        plot = ch.plots[0]
        plot.gap_width = gap_width
        if overlap is not None:
            plot.overlap = overlap
        elif stacked:
            plot.overlap = 100
        pos = XL_LABEL_POSITION.CENTER if stacked else XL_LABEL_POSITION.OUTSIDE_END
        for i, s in enumerate(plot.series):
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = _rgb(colors[i % len(colors)])
            s.invert_if_negative = False
            if labels:                       # 系列级数据标签（点级覆盖不会让其他点丢标签）
                dl = s.data_labels
                dl.show_value = True
                dl.number_format = number_format
                dl.number_format_is_linked = False
                dl.font.size = Pt(size)
                dl.font.color.rgb = _rgb(WHITE if stacked else label_color)
                dl.position = pos
        if highlight:
            s0 = plot.series[0]
            for k, v in highlight.items():
                idx = k if isinstance(k, int) else list(categories).index(k)
                clr = {"good": GOOD, "bad": BAD, "neutral": DARK}.get(v, v)
                pt = s0.points[idx]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = _rgb(clr)
                if labels and not stacked:
                    pdl = pt.data_label
                    pdl.font.color.rgb = _rgb(clr)
                    pdl.font.bold = True
                    pdl.font.size = Pt(size)
                    pdl.position = pos
                    dLbl = pdl._dLbl
                    nf = dLbl.find(qn("c:numFmt"))
                    if nf is None:
                        nf = etree.Element(qn("c:numFmt"))
                        dLbl.find(qn("c:idx")).addnext(nf)
                    nf.set("formatCode", number_format)
                    nf.set("sourceLinked", "0")
        va, ca = ch.value_axis, ch.category_axis
        va.has_major_gridlines = False
        va.visible = axis
        va.tick_labels.font.size = Pt(size)
        all_vals = [v for _, vals in ser for v in vals if v is not None]
        if all_vals and min(all_vals) >= 0:
            va.minimum_scale = 0             # 柱状图必须从 0 起，否则夸大差异
        elif all_vals:
            ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW   # 有负值时类别标签放到轴外侧
        ca.tick_labels.font.size = Pt(size)
        ca.major_tick_mark = XL_TICK_MARK.NONE
        ca.format.line.color.rgb = _rgb(LINE)
        if horizontal:
            ca.reverse_order = True
        show_legend = legend if legend is not None else len(ser) > 1
        ch.has_legend = show_legend
        if show_legend:
            ch.legend.position = XL_LEGEND_POSITION.TOP
            ch.legend.include_in_layout = False
            ch.legend.font.size = Pt(size)
        return gf

    def line_chart(self, x, y, w, h, categories, series, colors=None, number_format="0.0", labels=True,
                   legend=None, size=10, smooth=False, axis=False, label_last_only=False):
        """折线图（趋势）。多系列默认 红=本期/重点、灰=对比。"""
        ser = self._series(series)
        cd = CategoryChartData()
        cd.categories = categories
        for name, vals in ser:
            cd.add_series(name, vals)
        gf = self.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, _e(x), _e(y), _e(w), _e(h), cd)
        ch = gf.chart
        ch.has_title = False
        _chart_font(ch, size)
        colors = colors or ([DARK] if len(ser) == 1 else SERIES)
        for i, s in enumerate(ch.plots[0].series):
            clr = _rgb(colors[i % len(colors)])
            s.format.line.color.rgb = clr
            s.format.line.width = Pt(2)
            s.smooth = smooth
            s.marker.style = 8   # circle
            s.marker.size = 6
            s.marker.format.fill.solid()
            s.marker.format.fill.fore_color.rgb = clr
            s.marker.format.line.color.rgb = clr
            if labels:
                if label_last_only:
                    dl = s.points[len(categories) - 1].data_label
                    dl.has_text_frame = False
                    dl.position = XL_LABEL_POSITION.ABOVE
                    dl.number_format = number_format
                    dl.number_format_is_linked = False
                    dl.font.size = Pt(size)
                    dl.font.color.rgb = clr
                    dl.show_value = True
                else:
                    s.data_labels.show_value = True
                    s.data_labels.number_format = number_format
                    s.data_labels.number_format_is_linked = False
                    s.data_labels.position = XL_LABEL_POSITION.ABOVE if i == 0 else XL_LABEL_POSITION.BELOW
                    s.data_labels.font.size = Pt(size)
                    s.data_labels.font.color.rgb = clr
        va, ca = ch.value_axis, ch.category_axis
        va.has_major_gridlines = False
        va.visible = axis
        ca.major_tick_mark = XL_TICK_MARK.NONE
        ca.format.line.color.rgb = _rgb(LINE)
        ca.tick_labels.font.size = Pt(size)
        show_legend = legend if legend is not None else len(ser) > 1
        ch.has_legend = show_legend
        if show_legend:
            ch.legend.position = XL_LEGEND_POSITION.TOP
            ch.legend.include_in_layout = False
            ch.legend.font.size = Pt(size)
        return gf

    def doughnut(self, x, y, w, h, categories, values, colors=None, number_format="0%", size=10,
                 legend=True, hole=55):
        """环形图（结构/占比）。values 传占比小数或绝对值均可。"""
        cd = CategoryChartData()
        cd.categories = categories
        cd.add_series("占比", values)
        gf = self.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, _e(x), _e(y), _e(w), _e(h), cd)
        ch = gf.chart
        ch.has_title = False
        _chart_font(ch, size)
        colors = colors or [RED, MID, DARK, PINK, BLUE, LINE]
        s = ch.plots[0].series[0]
        for i in range(len(categories)):
            pt = s.points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = _rgb(colors[i % len(colors)])
            pt.format.line.color.rgb = _rgb(WHITE)
        plot = ch.plots[0]
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.show_percentage = number_format.endswith("%") and max(values) > 1
        dl.show_value = not dl.show_percentage
        dl.number_format = number_format
        dl.number_format_is_linked = False
        dl.font.size = Pt(size)
        dl.font.color.rgb = _rgb(WHITE)
        dl.font.bold = True
        hs = ch._chartSpace.find(".//" + qn("c:holeSize"))
        if hs is not None:
            hs.set("val", str(hole))
        ch.has_legend = legend
        if legend:
            ch.legend.position = XL_LEGEND_POSITION.RIGHT
            ch.legend.include_in_layout = False
            ch.legend.font.size = Pt(size)
        return gf

    # ---------- 其他
    def image(self, path, x, y, w=None, h=None):
        return self.shapes.add_picture(path, _e(x), _e(y), _e(w) if w else None, _e(h) if h else None)

    def source(self, text, size=9):
        """页脚来源/口径说明（左下角，避开页码）。"""
        return self.text(L, 6.93, 11.6, 0.32, text, size=size, color=GRAY, anchor="m", fit=True, min_size=8)

    def notes(self, text):
        """演讲者备注。"""
        self.slide.notes_slide.notes_text_frame.text = text

    def cols(self, ratios, x=L, w=W, gap=0.25):
        """按比例切分列：cols([1,2,1]) → [(x0,w0), (x1,w1), (x2,w2)]"""
        if isinstance(ratios, int):
            ratios = [1] * ratios
        tot = sum(ratios)
        avail = w - gap * (len(ratios) - 1)
        out, cx = [], x
        for r in ratios:
            cw = avail * r / tot
            out.append((cx, cw))
            cx += cw + gap
        return out

    def rows(self, ratios, y=None, h=None, gap=0.2):
        """按比例切分行：rows([1,1], y, h) → [(y0,h0), (y1,h1)]；默认用 lead() 之后的剩余空间。"""
        y = self.top if y is None else y
        h = (B - y) if h is None else h
        if isinstance(ratios, int):
            ratios = [1] * ratios
        tot = sum(ratios)
        avail = h - gap * (len(ratios) - 1)
        out, cy = [], y
        for r in ratios:
            ch = avail * r / tot
            out.append((cy, ch))
            cy += ch + gap
        return out


# ================================================================ 整套演示文稿
class UnicomDeck:
    def __init__(self, template=TEMPLATE):
        self.prs = Presentation(template)
        self.warnings = []
        sld = self.prs.slides._sldIdLst
        for s in list(sld):                       # 清空母版自带的示例页
            self.prs.part.drop_rel(s.rId)
            sld.remove(s)
        self._layouts = {l.name: l for l in self.prs.slide_layouts}

    def _new(self, layout_name, kind):
        slide = self.prs.slides.add_slide(self._layouts[layout_name])
        for ph in list(slide.placeholders):       # 不使用占位符，避免出现「单击此处…」
            ph._element.getparent().remove(ph._element)
        return Page(self, slide, kind)

    # ---------- 封面
    def cover(self, title, org="", date="", subtitle=None):
        """封面：右上 Logo + 红色竖线 + 大标题 + 单位/日期。"""
        pg = self._new(LAYOUT_COVER, "cover")
        pg.image(LOGO, 8.87, 0.60, w=3.725)
        pg.line(1.652, 2.905, 1.652, 6.18, color=COVER_RED, width=4.5).line._get_or_add_ln().set("cmpd", "thinThick")
        size = 44
        n = len(strip_markup(title))
        if n > 18:
            size = 40 if n <= 22 else 36
        tb = pg.text(1.98, 2.88, 10.9, 1.7, title, size=size, color=BLACK, bold=True, anchor="t",
                     line_spacing=1.05, min_size=28)
        need = estimate_tf_height(tb._element.find(qn("p:txBody")), tb.width)
        tb.height = _e(need + 0.02)
        y = 2.88 + need + 0.1
        one = dict(line_spacing=1.0, margin=(0.05, 0.02, 0.05, 0.02), fit=True)   # 单行文字
        if subtitle:
            pg.text(1.98, y, 10.9, 0.52, subtitle, size=24, color=DARK, bold=False, min_size=16, **one)
        if org:
            pg.text(2.07, 5.19, 9, 0.47, org, size=22, color="808080", bold=True, min_size=14, **one)
        if date:
            pg.text(2.07, 5.66, 9, 0.47, date, size=22, color="808080", bold=True, min_size=14, **one)
        return pg

    # ---------- 目录 / 章节过渡页
    def toc(self, items, active=0):
        """目录/章节过渡页（母版红色 V 形背景）。active 为当前章节序号（0 起），高亮为红色。
        惯例：封面后放 toc(active=0)，每进入新章节前再放一次 toc(active=i)。"""
        n = len(items)
        if not 1 <= n <= 6:
            raise ValueError(f"toc() 支持 1–6 个章节，收到 {n} 个；章节过多时请合并")
        if not 0 <= active < n:
            raise ValueError(f"active={active} 超出章节范围 0–{n - 1}")
        pg = self._new(LAYOUT_TOC, "toc")
        nums = "一二三四五六七八九十"
        pitch = 1.256 if n <= 4 else (1.0 if n == 5 else 0.85)
        bh = 0.70 if n <= 4 else 0.62
        cy0 = 3.725 - pitch * (n - 1) / 2
        longest = max(len(strip_markup(t)) for t in items)
        size = 28 if longest <= 8 else (24 if longest <= 10 else (20 if longest <= 13 else 18))
        lw = max(3.77, min(4.55, longest * size / 72 * 1.05 + 0.5))
        for i, t in enumerate(items):
            y = cy0 + i * pitch - bh / 2
            clr = RED if i == active else MID
            for kind, x, w, txt in ((MSO_SHAPE.RECTANGLE, 7.358, 0.862, nums[i]),
                                    (MSO_SHAPE.PENTAGON, 8.712, lw, t)):
                shp = pg.shapes.add_shape(kind, _e(x), _e(y), _e(w), _e(bh))
                _strip_style(shp)
                if kind == MSO_SHAPE.PENTAGON:
                    shp.adjustments[0] = 0.003
                shp.fill.solid()
                shp.fill.fore_color.rgb = _rgb(clr)
                shp.line.color.rgb = _rgb("F4B9C0")
                shp.line.width = Pt(1)
                _shadow(shp)
                _frame(shp.text_frame, 0.05, "m", wrap=False)
                write_text(shp.text_frame, txt, size if kind == MSO_SHAPE.PENTAGON else 28, WHITE, True, "c", 1.0)
        return pg

    # ---------- 内容页
    def page(self, title, size=30):
        """标准内容页：母版自带红色标题线、右上 Logo、右下页码。返回 Page。"""
        pg = self._new(LAYOUT_CONTENT, "content")
        n = sum(_char_em(c) for c in strip_markup(title))
        s = size
        while s > 22 and n * s / 72 > 10.75:
            s -= 1
        if n * s / 72 > 10.75:
            self.warnings.append(f"第{pg.index}页：标题过长（建议 ≤ 26 字），当前 {len(title)} 字")
        pg.text(0.27, 0.19, 10.9, 0.66, title, size=s, color=BLACK, bold=True, anchor="m",
                fit=False, margin=(0.1, 0.03, 0.1, 0.03))
        return pg

    # ---------- 结束页
    def ending(self, text=None):
        """结束页。默认使用母版的「做难事，必有所得」整页图片；传入 text（如 '谢谢！'）则生成简洁版。"""
        pg = self._new(LAYOUT_CONTENT if text is None else LAYOUT_COVER, "ending")
        if text is None:
            pic = pg.image(ENDING_IMG, 0, 0, w=SLIDE_W, h=SLIDE_H)
            pic.crop_top = 0.045      # 裁去原图顶部的英文路牌
        else:
            pg.image(LOGO, 8.87, 0.60, w=3.725)
            pg.line(1.652, 2.905, 1.652, 4.6, color=COVER_RED, width=4.5).line._get_or_add_ln().set("cmpd", "thinThick")
            pg.text(1.98, 3.2, 9, 1.1, text, size=48, color=BLACK, bold=True, anchor="m", fit=False)
        return pg

    # ---------- 输出
    def save(self, path, check=True):
        self.prs.save(path)
        problems = list(self.warnings)
        if check:
            problems += lint(path)
        if problems:
            print(f"[unicom_ppt] 已保存 {path}，发现 {len(problems)} 处需要人工复核：")
            for p in problems:
                print("  - " + p)
        else:
            print(f"[unicom_ppt] 已保存 {path}，自动检查未发现问题（仍需渲染后目检）。")
        return problems


# ================================================================ 自动检查（也可用于任何 pptx）
def _boxes_overlap(a, b, tol=0.03):
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    ix = min(ax2, bx2) - max(ax1, bx1)
    iy = min(ay2, by2) - max(ay1, by1)
    if ix <= tol or iy <= tol:
        return False
    contains = lambda o, i: o[0] - tol <= i[0] and o[1] - tol <= i[1] and o[2] + tol >= i[2] and o[3] + tol >= i[3]
    return not (contains(a, b) or contains(b, a))


def lint(path):
    """静态检查：越界、文字溢出（估算）、字号过小、残留占位文字、文字框相互压叠。"""
    prs = Presentation(path)
    sw, sh = prs.slide_width / EMU, prs.slide_height / EMU
    out = []
    bad_words = re.compile(r"请输入|单击此处|X{2,}|xxx|lorem|ipsum|待补充|TODO")
    for si, slide in enumerate(prs.slides, 1):
        boxes = []
        for shp in slide.shapes:
            if shp.left is None or shp.width is None:
                continue
            x, y = shp.left / EMU, shp.top / EMU
            w, h = shp.width / EMU, shp.height / EMU
            name = shp.name
            if x < -0.01 or y < -0.01 or x + w > sw + 0.01 or y + h > sh + 0.01:
                out.append(f"第{si}页：「{name}」超出页面边界")
            if shp.has_text_frame and shp.text_frame.text.strip():
                txt = shp.text_frame.text.strip()
                if bad_words.search(txt):
                    out.append(f"第{si}页：残留占位/待补充文字「{txt[:20]}」")
                if re.search(r"</?(?:b|good|bad|red|green|blue|gray)>", txt):
                    out.append(f"第{si}页：「{txt[:20]}」中出现未解析的标记标签（该参数可能不支持标记）")
                for r in shp._element.iter(qn("a:rPr")):
                    if r.get("sz") and int(r.get("sz")) < 900:
                        out.append(f"第{si}页：「{txt[:12]}」字号小于 9pt")
                        break
                body = shp._element.find(qn("p:txBody"))
                bodyPr = body.find(qn("a:bodyPr"))
                if bodyPr is not None and bodyPr.get("wrap") == "none":
                    continue
                need = estimate_tf_height(body, shp.width)
                for o in find_orphans(body, shp.width):
                    out.append(f"第{si}页：「{o}…」末行可能只剩1–2个字（孤字），可微调措辞或框宽（估算，以渲染为准）")
                auto = bodyPr is not None and bodyPr.find(qn("a:spAutoFit")) is not None
                if auto:
                    if y + need > sh:
                        out.append(f"第{si}页：「{txt[:12]}」自动扩展后超出页面底部")
                elif need > h + 0.08:
                    out.append(f"第{si}页：「{txt[:12]}」文字可能溢出（需 {need:.2f}in / 框 {h:.2f}in）")
                geom = shp._element.find(".//" + qn("a:prstGeom"))
                nest = geom is not None and geom.get("prst") in ("chevron", "homePlate")
                boxes.append(((x, y, x + w, y + h), txt[:12], nest))
            elif shp.has_table or shp.has_chart:
                boxes.append(((x, y, x + w, y + h), name, False))
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                if boxes[i][2] and boxes[j][2]:      # 阶段链燕尾形本身就是嵌套排列
                    continue
                if _boxes_overlap(boxes[i][0], boxes[j][0]):
                    out.append(f"第{si}页：「{boxes[i][1]}」与「{boxes[j][1]}」可能重叠")
    return out


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python unicom_ppt.py <deck.pptx>   # 对已有文件做静态检查")
        sys.exit(1)
    res = lint(sys.argv[1])
    print("\n".join(res) if res else "未发现问题")
