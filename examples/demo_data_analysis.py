"""
数据分析类示例：某市分公司经营分析（全部为示例数据）

要点演示：
1. 金字塔：封面后先给「核心结论」页，每页 lead() 一句话结论，主体为支撑论据
2. 好红差绿：用 delta / pp / growth / vs_target 生成着色数字，而不是手写颜色
3. 比率类指标变化一律用 pp；绝对量变化用 %
4. 所有派生数字由原始数据计算，避免手算错误；页脚写清来源与口径

运行：python examples/demo_data_analysis.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
from unicom_ppt import *  # noqa: E402,F401

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demo_data_analysis.pptx")

# ---------------------------------------------------------------- 原始数据（示例）
TIME_PROGRESS = 75.0                     # 9月末时序进度 %
PROVINCE_GROWTH = 3.9                    # 全省收入同比 %
months = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月"]
monthly_rev = [4.12, 4.05, 4.31, 4.28, 4.33, 4.36, 4.29, 4.40, 4.48]          # 亿元
pen_trend = [42.4, 42.3, 42.1, 42.0, 41.8, 41.7, 41.5, 41.3, 41.2]           # 宽带渗透率 %
PEN_START = 42.5                          # 年初
# 区县：名称, 1-9月收入, 去年同期, 全年目标, 渗透率(9月), 渗透率(年初)
districts = [
    ("A区", 9.85, 9.14, 13.10, 46.8, 47.2),
    ("B区", 7.62, 7.15, 10.20, 43.5, 44.1),
    ("C区", 6.48, 6.18, 8.95, 40.9, 42.6),
    ("D县", 5.37, 5.18, 7.40, 38.2, 39.8),
    ("E县", 4.96, 4.82, 6.90, 37.6, 39.9),
    ("F县", 4.34, 4.37, 6.28, 35.1, 36.8),
]
mobile_users, mobile_users_start = 512.3, 499.7                              # 万户
# 宽带渗透率下降归因（pp），三项之和须等于总变化
contrib = {"新装不足": -0.5, "拆机上升": -0.6, "住户基数扩大": -0.2}

# ---------------------------------------------------------------- 派生计算
rev = sum(monthly_rev)
rev_ly = sum(d[2] for d in districts)
target = sum(d[3] for d in districts)
assert abs(rev - sum(d[1] for d in districts)) < 0.005, "月度合计与区县合计不一致"
assert abs(sum(contrib.values()) - (pen_trend[-1] - PEN_START)) < 0.05, "归因之和与总变化不一致"
rev_growth = (rev / rev_ly - 1) * 100
completion = rev / target * 100
q4_monthly = (target - rev) / 3
q4_lift = (q4_monthly / (rev / 9) - 1) * 100
mom = (monthly_rev[-1] / monthly_rev[-2] - 1) * 100
leading = [d[0] for d in districts if d[1] / d[3] * 100 >= TIME_PROGRESS]
SRC = "数据来源：经分系统（示例数据）；统计口径：2025年1-9月累计，收入为不含税出账收入"

deck = UnicomDeck()

# 1 封面 -------------------------------------------------------------
deck.cover("2025年1-9月经营分析报告", org="示例市分公司 市场部", date="2025年10月")

# 2 核心结论（金字塔塔尖：整份报告的结论先行）--------------------------
pg = deck.page("核心结论")
pg.lead(f"前三季度收入稳中有进，但完成率落后时序{pp(completion, TIME_PROGRESS)}、宽带渗透率持续下滑，"
        f"四季度须聚焦存量维系与落后区县追赶")
items = [
    ("收入稳中有进，全年目标承压",
     f"1-9月收入<b>{rev:.2f}亿元</b>，同比{growth(rev, rev_ly)}，领先全省{pp(rev_growth, PROVINCE_GROWTH)}；"
     f"完成率{completion:.1f}%，落后时序进度{pp(completion, TIME_PROGRESS)}，四季度月均需达<b>{q4_monthly:.2f}亿元</b>"
     f"（较前三季度月均{delta(q4_lift, better=None)}）",
     growth(rev, rev_ly), "收入同比"),
    ("宽带渗透率下滑，拆机上升是首因",
     f"9月宽带渗透率{pen_trend[-1]:.1f}%，较年初{pp(pen_trend[-1], PEN_START)}；"
     f"其中拆机上升贡献{delta(contrib['拆机上升'], 'pp')}，新装不足贡献{delta(contrib['新装不足'], 'pp')}",
     pp(pen_trend[-1], PEN_START), "渗透率较年初"),
    ("区县分化明显，E县、F县是追赶重点",
     f"仅{'、'.join(leading)}完成率跑赢时序进度；E县、F县分别落后"
     f"{pp(districts[4][1] / districts[4][3] * 100, TIME_PROGRESS)}、{pp(districts[5][1] / districts[5][3] * 100, TIME_PROGRESS)}，"
     f"且F县收入同比{growth(districts[5][1], districts[5][2])}",
     f"{len(leading)}/{len(districts)}", "跑赢时序区县", "bad"),
]
pg.conclusions(items)   # 每行：判断标题 + 证据句 + 右侧关键数字（与正文同色）
pg.source(SRC)

# 3 目录 -------------------------------------------------------------
sections = ["经营概览", "问题诊断", "下一步举措"]
deck.toc(sections, active=0)

# 4 经营概览：KPI + 趋势 + 洞察 --------------------------------------
pg = deck.page("一、经营概览：收入稳中有进，全年目标承压")
y = pg.lead(f"1-9月收入同比{growth(rev, rev_ly)}、领先全省{pp(rev_growth, PROVINCE_GROWTH)}，"
            f"但完成率落后时序{pp(completion, TIME_PROGRESS)}，四季度月均收入需达<b>{q4_monthly:.2f}亿元</b>")
cards = [
    ("营业收入（1-9月）", f"{rev:.2f}", "亿元", "同比 " + growth(rev, rev_ly), None),
    ("收入完成率", num(completion), "%", f"时序进度 {TIME_PROGRESS:.0f}%  " + pp(completion, TIME_PROGRESS), "bad"),
    ("移动用户", f"{mobile_users:.1f}", "万户", "较年初 " + growth(mobile_users, mobile_users_start), None),
    ("宽带渗透率", num(pen_trend[-1]), "%", "较年初 " + pp(pen_trend[-1], PEN_START), None),
]   # 大数字只在其本身就是评价（完成率、达标率、排名）时着色；水平值保持黑色，由变化值着色
for (cx, cw), (lab, val, unit, chg, judge) in zip(pg.cols(4, gap=0.2), cards):
    pg.kpi(cx, y, cw, 1.25, lab, val, unit, chg, judge=judge)
y += 1.25 + 0.25
(lx, lw), (rx, rw) = pg.cols([3, 2], gap=0.35)
pg.heading(lx, y, lw, "月度收入（亿元）")
pg.bar(lx, y + 0.4, lw, B - y - 0.4, months, monthly_rev, highlight={"9月": "good"}, number_format="0.00")
pg.heading(rx, y, rw, "关键发现")
pg.bullets(rx, y + 0.45, rw, B - y - 0.45, [
    f"<b>单月创新高：</b>9月收入{monthly_rev[-1]:.2f}亿元，环比{delta(mom)}，为年内单月最高水平",
    f"<b>增速领先：</b>累计同比{growth(rev, rev_ly)}，高于全省{PROVINCE_GROWTH:.1f}%，领先{pp(rev_growth, PROVINCE_GROWTH)}",
    f"<b>进度落后：</b>完成率{completion:.1f}%，落后时序{pp(completion, TIME_PROGRESS)}；"
    f"完成全年目标{target:.2f}亿元，四季度月均需{q4_monthly:.2f}亿元，较前三季度月均{delta(q4_lift, better=None)}",
], size=14)
pg.source(SRC)

# 5 区县对标：表格 + 结论 --------------------------------------------
pg = deck.page("一、经营概览：区县分化明显")
e_gap = districts[4][1] / districts[4][3] * 100
f_gap = districts[5][1] / districts[5][3] * 100
worst = min(districts, key=lambda d: d[4] - d[5])          # 渗透率降幅最大的区县
y = pg.lead(f"仅{'、'.join(leading)}跑赢时序进度；E县、F县完成率分别落后{pp(e_gap, TIME_PROGRESS)}、"
            f"{pp(f_gap, TIME_PROGRESS)}，{worst[0]}渗透率降幅全市最大（{pp(worst[4], worst[5])}），两县是追赶重点")
rows = [["区县", "收入（亿元）", "同比", "完成率", "较时序进度", "宽带渗透率", "较年初"]]
for name, r, ly, tg, pen, pen0 in districts:
    c = r / tg * 100
    rows.append([name, f"{r:.2f}", growth(r, ly), f"{c:.1f}%", pp(c, TIME_PROGRESS), f"{pen:.1f}%", pp(pen, pen0)])
rows.append(["<b>全市</b>", f"<b>{rev:.2f}</b>", growth(rev, rev_ly), f"<b>{completion:.1f}%</b>",
             pp(completion, TIME_PROGRESS), f"<b>{pen_trend[-1]:.1f}%</b>", pp(pen_trend[-1], PEN_START)])
(tx, tw), (cx, cw) = pg.cols([7, 3], gap=0.35)
pg.table(tx, y, tw, B - y, rows, col_widths=[1.1, 1.3, 1, 1, 1.2, 1.2, 1],
         size=13, highlight_rows=[len(rows) - 1], highlight_fill=LIGHT)
pg.heading(cx, y, cw, "对标结论")
notes = [
    (f"<b>领跑：A区</b>\n收入同比{growth(9.85, 9.14)}，唯一跑赢时序的区县，可复制其存量经营做法", LIGHT_RED),
    (f"<b>追赶：E县、F县</b>\n完成率落后时序3pp以上，渗透率降幅{pp(37.6, 39.9)}、{pp(35.1, 36.8)}", LIGHT),
    (f"<b>关注：C区</b>\n收入体量第三，但渗透率较年初{pp(40.9, 42.6)}，降幅高于全市", LIGHT),
]
for (by, bh), (txt, fc) in zip(pg.rows(3, y=y + 0.48, h=B - y - 0.48, gap=0.18), notes):
    pg.box(cx, by, cw, bh, txt, fill=fc, line=None, size=13, line_spacing=1.3, margin=0.15, anchor="m")
pg.source(SRC + "；时序进度按9/12计")

# 6 目录 -------------------------------------------------------------
deck.toc(sections, active=1)

# 7 问题诊断：议题树（问题 → 原因 → 证据）----------------------------
pg = deck.page("二、问题诊断：宽带渗透率下滑的原因")
y = pg.lead(f"宽带渗透率较年初{pp(pen_trend[-1], PEN_START)}，拆机上升是首要原因（贡献{delta(contrib['拆机上升'], 'pp')}），"
            f"应优先抓存量维系")
(ax, aw), (bx, bw), (ex, ew) = pg.cols([2.3, 3, 5.2], gap=0.5)
top, bottom = y + 0.05, B - 0.75
pg.label(ax, top, aw, 0.4, "问题", fill=DARK, size=12)
pg.label(bx, top, bw, 0.4, "原因拆解（贡献度）", fill=DARK, size=12)
pg.label(ex, top, ew, 0.4, "数据证据", fill=DARK, size=12)
ry0 = top + 0.6
rh = (bottom - ry0 - 0.3) / 3
mid_y = ry0 + (3 * rh + 0.3) / 2
pg.box(ax, mid_y - 0.95, aw, 1.9,
       f"<b>宽带渗透率下滑</b>\n\n9月 {pen_trend[-1]:.1f}%\n较年初 {pp(pen_trend[-1], PEN_START)}",
       fill=LIGHT_RED, line=None, size=14, align="c", anchor="m", line_spacing=1.2)
evidence = {
    "新装不足": f"月均新装<b>1.82万户</b>，同比{growth(1.82, 2.15)}；新建小区覆盖率仅62%",
    "拆机上升": f"月均拆机<b>1.64万户</b>，同比{growth(1.64, 1.38, better='down')}；合约到期用户拆机占58%",
    "住户基数扩大": f"新增住户<b>3.6万户</b>，主要集中在C区、E县新建小区，资源覆盖滞后",
}
for i, (k, v) in enumerate(contrib.items()):
    cy = ry0 + i * (rh + 0.15)
    is_top = v == min(contrib.values())
    pg.box(bx, cy, bw, rh, f"<b>{k}</b>　{delta(v, 'pp')}", fill=LIGHT_RED if is_top else LIGHT,
           line=RED if is_top else None, size=14, anchor="m", margin=0.15)
    pg.box(ex, cy, ew, rh, evidence[k], fill=None, line=LINE, size=12, anchor="m", margin=0.15, line_spacing=1.25)
    pg.line(ax + aw, mid_y, bx, cy + rh / 2, color=MID, width=1)
    pg.line(bx + bw, cy + rh / 2, ex, cy + rh / 2, color=MID, width=1, tail="triangle")
pg.box(L, B - 0.55, W, 0.5, f"<b>启示：</b>拆机贡献超过新装不足，四季度资源应向合约到期用户续约、"
       f"装维质量提升倾斜，同时补齐C区、E县新建小区覆盖", fill=WARM, line=None, size=13, anchor="m", margin=0.15)
pg.source("数据来源：宽带运营系统（示例数据）；渗透率=宽带在网用户/住户数，归因按因素分解法测算")

# 8 目录 -------------------------------------------------------------
deck.toc(sections, active=2)

# 9 下一步举措 -------------------------------------------------------
pg = deck.page("三、下一步举措")
y = pg.lead(f"围绕「保存量、补短板、促追赶」三条主线并强化保障，力争年底完成率达100%、渗透率止跌回升至<b>41.5%</b>")
rows = [
    ["主线", "关键举措", "量化目标", "牵头部门", "完成时间"],
    ["<b>保存量</b>", "合约到期用户提前90天续约；装维质量专项整治", "月均拆机降至<b>1.40万户</b>以下", "市场部、网络部", "11月底"],
    ["<b>补短板</b>", "C区、E县新建小区资源覆盖攻坚，同步驻点营销", "新建小区覆盖率提升至<b>85%</b>", "网络部", "12月底"],
    ["<b>促追赶</b>", "E县、F县「一县一策」，A区存量经营经验复制推广", "两县完成率追平时序进度", "市场部、各区县", "12月底"],
    ["<b>强保障</b>", "周通报、月复盘；四季度激励向存量维系倾斜", "收入月均达<b>{:.2f}亿元</b>".format(q4_monthly), "综合部、财务部", "持续"],
]
pg.table(L, y, W, 0.7 * len(rows), rows, col_widths=[1.1, 4.2, 2.8, 1.7, 1.2], size=13,
         align=["c", "l", "l", "c", "c"])
y += 0.7 * len(rows) + 0.3
pg.heading(L, y, W, "推进节奏")
pg.chevrons(L, y + 0.5, W, 0.62, ["10月：方案部署、目标分解到区县", "11月：存量维系与覆盖攻坚冲刺",
                                   "12月：收官冲刺、全年复盘"], size=14)
pg.source("注：目标值为示例，需与年度预算口径对齐后确认")

# 10 结束页 ----------------------------------------------------------
deck.ending()
deck.save(OUT)
