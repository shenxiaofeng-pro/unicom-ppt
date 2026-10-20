"""
传统报告类示例：三种页型各一页（业务内容均为示例，不代表真实系统或数据）

  类型一  体系架构与业务价值页：基础支撑 → 核心能力及协同 → 业务价值（25:50:25 三栏）
  类型二  专项攻坚与策略落地页：问题与措施 → 对象×策略映射 → 重点方案 → 执行要点
  类型三  能力贯通与运营闭环页：阶段链 → 各阶段分栏展开 → 结果回流（有反馈机制才画闭环）

逐页生成：每页一个函数，按顺序放进 PAGES，由 build() 运行
  python examples/demo_business_report.py --upto N   只生成前 N 页，检查并渲染第 N 页
  python examples/demo_business_report.py            全部页面 + 全量检查
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
from unicom_ppt import *  # noqa: E402,F401

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demo_business_report.pptx")
def p_cover(deck):
    deck.cover("经营管理数字化建设汇报", org="示例分公司", date="2025年10月", subtitle="体系建设 · 专项攻坚 · 数据赋能（示例）")


def p_architecture(deck):
    """类型一：体系架构与业务价值"""
    pg = deck.page("一、薪酬激励一体化平台")
    y0 = pg.lead("依托<b>人力、财务、经分</b>三大系统数据，建设<b>薪酬激励一体化平台</b>，"
                 "打通「指标—考核—薪酬—分析」全链条，实现激励规则<red>线上配置</red>、"
                 "薪酬结果<red>自动核算</red>、激励效果<red>按月评价</red>", bold=False, size=14)
    (lx, lw), (mx, mw), (rx, rw) = pg.cols([1, 2, 1], gap=0.3)
    pg.heading(lx, y0, lw, "基础支撑")
    pg.heading(mx, y0, mw, "平台架构")
    y1 = pg.heading(rx, y0, rw, "业务价值") + 0.04       # heading() 返回下方可用 y

    # 左：基础支撑（依托什么）
    supports = [("人力资源系统", ["组织架构、岗位序列、人员档案", "数据按日同步"]),
                ("财务系统", ["薪酬总额、成本中心、发放台账"]),
                ("经营分析系统", ["收入、用户、业务量等考核指标", "数据按月抽取"])]
    for (sy, sh), (name, desc) in zip(pg.rows(3, y=y1, h=B - y1, gap=0.15), supports):
        pg.box(lx, sy, lw, sh, fill=LIGHT_RED, line=None)
        pg.label(lx, sy, lw, 0.42, name, fill=PINK, color=RED, size=13, align="l")
        pg.text(lx + 0.05, sy + 0.48, lw - 0.1, sh - 0.52, desc, size=12, line_spacing=1.25)

    # 中：平台架构（能力如何组织）
    # 纵向切分：顶部平台名称条 + rows() 按比例切出「服务对象 / 核心能力 / 数据底座」三层，层间留出放连线的间距
    tag_w = 0.95                                          # 左侧层名标签宽度
    bx, bw = mx + tag_w + 0.1, mw - tag_w - 0.1           # 层内容区
    pg.label(mx, y1, mw, 0.42, "薪酬激励一体化平台", size=14)
    (svc_y, svc_h), (cap_y, cap_h), (base_y, base_h) = pg.rows([0.6, 1.6, 0.8], y=y1 + 0.55, h=B - y1 - 0.55, gap=0.42)
    # 服务对象层
    pg.label(mx, svc_y, tag_w, svc_h, "服务对象", fill=DARK, size=11)
    for (cx, cw), t in zip(pg.cols(4, x=bx, w=bw, gap=0.12),
                           ["省公司管理层", "地市人力部门", "区县/网格长", "一线员工"]):
        pg.box(cx, svc_y + 0.08, cw, svc_h - 0.16, t, fill=WHITE, line=MID, size=11, align="c", anchor="m", shape="round")   # 同排统一字号
    # 核心能力层
    pg.arrow(bx + bw / 2, cap_y, bx + bw / 2, svc_y + svc_h - 0.08, color=MID, text="服务支撑")
    pg.label(mx, cap_y, tag_w, cap_h, "核心能力", fill=DARK, size=11)
    modules = [("考核指标管理", "指标库配置\n目标分解下达\n权重动态调整"),
               ("薪酬测算发放", "绩效自动核算\n薪酬模拟测算\n发放校验留痕"),
               ("激励效果分析", "激励投入产出\n异常波动预警\n人均效能对标")]
    mcols = pg.cols(3, x=bx, w=bw, gap=0.15)
    for (cx, cw), (t, items) in zip(mcols, modules):
        pg.box(cx, cap_y, cw, cap_h, fill=WHITE, line=RED, line_w=1)
        pg.text(cx, cap_y + 0.06, cw, 0.36, f"<red>{t}</red>", size=13, align="c", anchor="m")
        pg.line(cx + 0.15, cap_y + 0.45, cx + cw - 0.15, cap_y + 0.45, color=PINK)
        pg.text(cx, cap_y + 0.5, cw, cap_h - 0.55, items, size=12, align="c", line_spacing=1.3)
    # 模块协同箭头 —— 依据是本示例总述中的「指标—考核—薪酬—分析」全链条；材料没有先后关系时不要画
    for i in range(2):
        a = mcols[i][0] + mcols[i][1]
        pg.arrow(a, cap_y + cap_h / 2 + 0.25, mcols[i + 1][0], cap_y + cap_h / 2 + 0.25, color=RED, width=1.25)
    # 数据底座层 —— 本示例材料只说明数据从各系统同步到平台，所以画单向「数据支撑」；
    # 只有材料写明平台会回写数据时，才用 both=True 画双向「数据交换」
    pg.arrow(bx + bw / 2, base_y, bx + bw / 2, cap_y + cap_h, color=MID, text="数据支撑")
    pg.label(mx, base_y, tag_w, base_h, "数据底座", fill=DARK, size=11)
    pg.box(bx, base_y, bw, base_h, fill=WARM, line=None)
    for (cx, cw), t in zip(pg.cols(4, x=bx + 0.12, w=bw - 0.24, gap=0.12), ["组织信息", "人员信息", "考核指标", "薪酬规则"]):
        pg.box(cx, base_y + 0.1, cw, base_h - 0.2, t, fill=WHITE, line=None, size=12, align="c", anchor="m")

    # 右：业务价值（带来什么改善——写改善，不写功能名）
    values = [("规则透明", "考核与薪酬规则线上公示，员工可查可算，减少口径争议"),
              ("核算提效", "月度薪酬核算由人工汇总改为系统自动生成，减少跨系统取数"),
              ("激励精准", "按投入产出评价激励效果，资源向高效能团队倾斜")]
    vrows = pg.rows(3, y=y1, h=B - y1, gap=0.0)
    for i, ((vy, vh), (t, d)) in enumerate(zip(vrows, values)):
        pg.badge(rx, vy + 0.1, 0.4, i + 1, size=13)
        pg.text(rx + 0.5, vy + 0.08, rw - 0.5, 0.44, f"<b>{t}</b>", size=15, color=BLACK, anchor="m")
        pg.text(rx + 0.5, vy + 0.55, rw - 0.5, vh - 0.65, d, size=12, line_spacing=1.3)
        if i < 2:
            pg.line(rx, vy + vh, rx + rw, vy + vh, color=LINE, dash=True)


def p_campaign(deck):
    """类型二：专项攻坚与策略落地"""
    pg = deck.page("二、存量宽带融合迁转专项攻坚")
    pg.lead_text = "面向三类存量客群分群施策，以家庭融合升级为主线，四季度融合迁转6.2万户"   # 本页用问题/措施代替结论条，手动登记结论
    y = T
    (ax, aw), (cx2, cw2) = pg.cols(2, gap=0.3)
    for (bx_, bw_), tag, txt in (
            ((ax, aw), "现状和问题", "单宽带用户占比偏高，融合率低于全省平均，到期拆机风险集中；"
                                     "现有套餐与家庭多成员需求<b>不匹配</b>"),
            ((cx2, cw2), "解决措施", "面向<b>三类存量客群</b>分群施策，以家庭融合升级为主线，"
                                    "四季度融合迁转<red>6.2万户</red>")):
        pg.label(bx_, y, 1.25, 0.82, tag, size=13)
        pg.box(bx_ + 1.25, y, bw_ - 1.25, 0.82, txt, fill=LIGHT, line=None, size=12, anchor="m",
               margin=0.12, line_spacing=1.25)
    y += 0.82 + 0.25
    (mx, mw), (px, pw) = pg.cols([2.2, 1], gap=0.35)
    mid_bottom = 5.1
    # 中部：对象 × 策略 × 方案 映射
    sub = pg.cols([2.2, 2.6, 2.3, 1.3], x=mx, w=mw, gap=0.3)
    for (sx, sw), t in zip(sub, ["目标客群（规模）", "分群条件 / 策略", "迁转方案", "四季度目标"]):
        pg.label(sx, y, sw, 0.36, t, fill=DARK, size=12)
    segments = [
        ("单宽带用户", "32.6万户", "有移动需求、在网≥12个月", "宽带+移动融合入网", "3.0万户"),
        ("低值融合用户", "18.4万户", "套餐档位低、流量饱和度高", "融合套餐升档", "2.2万户"),
        ("家庭成员异网", "9.8万户", "宽带家庭内有异网成员", "家庭共享副卡", "1.0万户"),
    ]
    ry = y + 0.5
    rh = (mid_bottom - ry - 0.2) / 3
    for i, (seg, scale, cond, plan, goal) in enumerate(segments):
        yy = ry + i * (rh + 0.1)
        (s0x, s0w), (s1x, s1w), (s2x, s2w), (s3x, s3w) = sub
        pg.box(s0x, yy, s0w, rh, f"<b>{seg}</b>\n{scale}", fill=LIGHT_RED, line=None, size=12, align="c",
               anchor="m", line_spacing=1.2)
        pg.box(s1x, yy, s1w, rh, cond, fill=WHITE, line=LINE, size=12, anchor="m", align="c")
        pg.box(s2x, yy, s2w, rh, f"<b>{plan}</b>", fill=WHITE, line=RED, size=12, anchor="m", align="c")
        pg.text(s3x, yy, s3w, rh, f"<red>{goal}</red>", size=14, align="c", anchor="m")
        pg.arrow(s0x + s0w, yy + rh / 2, s1x, yy + rh / 2, color=MID)
        pg.arrow(s1x + s1w, yy + rh / 2, s2x, yy + rh / 2, color=MID)
    # 右侧：重点方案
    pg.heading(px, y - 0.04, pw, "重点方案")
    pg.label(px, y + 0.42, pw, 0.42, "主推：家庭融合升级包", size=13)
    pg.box(px, y + 0.84, pw, mid_bottom - y - 0.84, [
        "<b>适用对象：</b>单宽带及低值融合家庭",
        "<b>差异点：</b>宽带提速与家庭副卡打包，一次办理",
        "<b>预期作用：</b>提升融合率，降低到期拆机",
        "<gray>（资费与权益以正式方案为准）</gray>",
    ], fill=LIGHT_RED, line=None, size=13, line_spacing=1.4, space_after=6, margin=0.16)
    # 底部：执行要点
    pg.heading(L, mid_bottom + 0.2, W, "执行要点")
    ky = mid_bottom + 0.64
    for (kx, kw), (t, d) in zip(pg.cols(4, gap=0.2), [
            ("口径统一", "统一客群清单与迁转口径，周更新"),
            ("渠道执行", "厅店、装维、外呼三渠道同步触达"),
            ("组织分工", "市场部牵头，区县落实到网格"),
            ("过程监测", "日跟踪、周通报，落后网格专项督导")]):
        pg.box(kx, ky, kw, B - ky, f"<b>{t}</b>\n{d}", fill=LIGHT, line=None, size=12, line_spacing=1.25,
               margin=0.12, anchor="m")
    pg.source("注：客群规模与目标为示例数据")


def p_capability(deck):
    """类型三：能力贯通与运营闭环"""
    pg = deck.page("三、数据能力精准助销")
    y = pg.lead("以<b>用户标签与预测模型</b>为基础，形成可配置的营销策略，下发至厅店、装维、外呼等一线场景，"
                "执行结果<red>按日回流</red>，驱动绩效核算与策略迭代", bold=False, size=14)
    # chevrons() 返回与各阶段节点对齐的内容列；场景列内容最多，用 ratios 加宽
    cols4 = pg.chevrons(L, y, W, 0.55, ["数据与基础能力", "策略与规则", "场景与一线动作", "结果反馈与激励"],
                        size=14, ratios=[1, 1, 1.25, 1])
    cy = y + 0.55 + 0.18
    col_bottom = B - 0.5                                   # 底部留出闭环线的位置
    # 第一列：数据能力
    (c1x, c1w), (c2x, c2w), (c3x, c3w), (c4x, c4w) = cols4
    for (yy, hh), (t, d) in zip(pg.rows(3, y=cy, h=col_bottom - cy, gap=0.3), [
            ("多源数据汇聚", "B域、O域、M域数据统一接入"),
            ("用户标签体系", "基础属性、消费、终端、家庭关系"),
            ("预测模型", "离网预警、升档倾向、融合潜力")]):
        pg.box(c1x, yy, c1w, hh, f"<b>{t}</b>\n{d}", fill=LIGHT, line=None, size=12, line_spacing=1.25,
               margin=0.12, anchor="m")
    # 第二列：策略与规则（纵向步骤）
    steps = ["客户识别与分群", "策略匹配与规则配置", "策略输出与分发"]
    srows = pg.rows(3, y=cy, h=col_bottom - cy, gap=0.3)   # 与其他三列相同的行切分
    for i, ((yy, hh), t) in enumerate(zip(srows, steps)):
        pg.box(c2x, yy, c2w, hh, f"<b>{t}</b>", fill=WHITE, line=RED, size=13, align="c", anchor="m")
        if i < 2:
            pg.arrow(c2x + c2w / 2, yy + hh + 0.02, c2x + c2w / 2, srows[i + 1][0] - 0.02, color=RED)
    # 第三列：场景（场景 + 对象/工具 + 动作）
    for (yy, hh), (scene, tool, act) in zip(pg.rows(3, y=cy, h=col_bottom - cy, gap=0.3), [
            ("厅店", "店员、助销工具", "到店识别，精准推荐"),
            ("智家装维", "装维工程师", "上门服务时融合推荐"),
            ("集约外呼", "外呼坐席", "高潜客户批量外呼")]):
        pg.label(c3x, yy, 0.9, hh, scene, fill=PINK, color=RED, size=12)
        pg.box(c3x + 0.9, yy, c3w - 0.9, hh, f"<gray>{tool}</gray>\n<b>{act}</b>", fill=WHITE, line=LINE,
               size=12, anchor="m", line_spacing=1.25, margin=0.1)
    # 第四列：结果反馈与激励
    for (yy, hh), (t, d) in zip(pg.rows(3, y=cy, h=col_bottom - cy, gap=0.3), [
            ("结果回传", "办理结果按日回流至平台"),
            ("进度监测", "执行进度看板，异常自动提醒"),
            ("绩效激励", "按成效核算绩效，激励按月兑现")]):
        pg.box(c4x, yy, c4w, hh, f"<b>{t}</b>\n{d}", fill=LIGHT, line=None, size=12, line_spacing=1.25,
               margin=0.12, anchor="m")
    # 闭环：仅在材料确有反馈机制时绘制。本示例材料写的是「驱动策略迭代」，所以指向策略列（第 2 列）；
    # 如果材料写的是「按月优化模型」，就应指向数据/模型所在的第 1 列
    pg.loop(cols4[3], cols4[1], top=col_bottom, text="执行结果回流，驱动策略迭代")


def p_ending(deck):
    deck.ending("谢谢！")


# 内容只有 3 页，按规范不放目录页（目录页用法见 demo_data_analysis.py）
PAGES = [p_cover, p_architecture, p_campaign, p_capability, p_ending]

if __name__ == "__main__":
    build(PAGES, OUT)
