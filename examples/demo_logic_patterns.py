"""
逻辑关系示例：用一份「提升宽带续约率的方案建议」演示金字塔结构与五种逻辑关系（全部为示例数据）

故事线（金字塔）：
  塔尖   以「提速+权益」续约为主方案，三个月内把到期续约率从 62% 提升到 70%
  支撑1  为什么要做 —— 续约率下滑沿因果链传导，最终拉低渗透率          → causal_chain()
  支撑2  选什么方案 —— 三个方案对比，「提速+权益」综合最优               → compare()
  支撑3  先做什么   —— 价值/难度矩阵，先推进高价值、低难度的三项举措     → matrix()
  支撑4  怎么持续   —— 建立月度续约运营闭环                             → cycle()
  核心结论页把四个支撑并列展示                                          → pyramid()

逐页生成（每页一个函数）：
  python examples/demo_logic_patterns.py --upto 1     生成并校验第 1 页
  python examples/demo_logic_patterns.py --upto 2     第 1 页没问题后，再做第 2 页 ……
  python examples/demo_logic_patterns.py              全部完成后做全量检查
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
from unicom_ppt import *  # noqa: E402,F401

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "demo_logic_patterns.pptx")
SRC = "数据来源：宽带运营系统、续约专项调研（示例数据）；统计口径：2025年1-9月"

# ---------------------------------------------------------------- 原始数据（示例）
RENEW_NOW, RENEW_GOAL = 62.0, 70.0          # 到期续约率 %
REACH_NOW, REACH_LY = 41.0, 53.0            # 到期前客户经理触达率 %
CHURN_M, CHURN_M_LY = 1.64, 1.38            # 月均拆机（万户）
PEN_NOW, PEN_START = 41.2, 42.5             # 宽带渗透率 %
CHURN_CONTRIB = -0.6                        # 拆机对渗透率变化的贡献（pp）


def p_cover(deck):
    deck.cover("提升宽带续约率的方案建议", org="示例市分公司 市场部", date="2025年10月",
               subtitle="金字塔结构与逻辑关系示例")


def p_summary(deck):
    """金字塔支撑：结论条即塔尖，下面并列四个支撑论点（为什么 / 选什么 / 先做什么 / 怎么持续）。"""
    pg = deck.page("核心结论")
    pg.lead(f"建议以「提速+权益」续约为主方案，三个月内把到期续约率从{num(RENEW_NOW, '%', 0)}"
            f"提升到<b>{num(RENEW_GOAL, '%', 0)}</b>")
    y = pg.pyramid(None, [
        ("续约率下滑正在拉低渗透率", [f"到期续约率仅{num(RENEW_NOW, '%', 0)}",
                                    f"月均拆机同比{growth(CHURN_M, CHURN_M_LY, better='down')}",
                                    f"拆机拖累渗透率{delta(CHURN_CONTRIB, 'pp')}"]),
        ("「提速+权益」综合最优", ["续约率提升接近最优", "收入基本不受损", "代价：需协调提速资源"]),
        ("先推进三项高价值、低难度举措", ["到期前90天预警", "提速权益包上架", "客户经理续约激励"]),
        ("建立月度续约运营闭环", ["预警→派单→触达→回传", "按月评估并优化权益与话术"]),
    ])
    # 金字塔下方：请领导决策的事项（汇报的「行动」落点）
    pg.heading(L, y + 0.3, W, "需决策事项")
    pg.box(L, y + 0.75, W, B - y - 0.75, [
        "<b>批准方案：</b>以「提速+权益」作为四季度到期续约的主方案，降价续约仅作为兜底",
        "<b>协调资源：</b>由网络部保障到期用户的提速资源，11月底前完成",
        "<b>考核导向：</b>把到期续约率纳入客户经理四季度考核",
    ], fill=WARM, line=None, bullets="square", size=14, line_spacing=1.3, space_after=4, anchor="m", margin=0.2)
    pg.source(SRC)


def p_causal(deck):
    """因果链：起因 → 传导 → 结果。每一环都要有证据；尚未证实的一环画虚线并标「待验证」。"""
    pg = deck.page("一、为什么要做：续约率下滑正在拉低渗透率")
    y = pg.lead(f"续约问题沿「提醒滞后 → 触达不足 → 竞对挖转 → 拆机上升」传导，拆机拖累渗透率"
                f"{delta(CHURN_CONTRIB, 'pp')}")
    top = pg.takeaway("<b>启示：</b>链条起点是提醒窗口过短，把预警提前到到期前90天、补齐触达，是阻断传导成本最低的环节")
    pg.causal_chain(
        [("到期提醒滞后", "仅提前30天"),
         ("续约触达不足", f"触达率{num(REACH_NOW, '%', 0)}"),
         ("竞对低价挖转", "异网集中优惠"),
         ("到期拆机上升", f"月均{num(CHURN_M, '万户', 2)}"),
         ("渗透率下滑", f"较年初{pp(PEN_NOW, PEN_START)}")],
        links=["触达窗口过短", f"触达率{pp(REACH_NOW, REACH_LY)}", "抽样：挖转占45%",
               f"贡献{delta(CHURN_CONTRIB, 'pp')}"],
        verified=[True, True, False, True],      # 第 3 环只有抽样调研，尚未被全量数据证实
        notes=[["系统提前30天提醒", "竞对提前60天"],                 # 窄卡片里每行一个短句，避免孤字
               ["去年同期触达率", num(REACH_LY, '%', 0)],
               ["回访抽样300户", "需全量核实"],
               ["月均拆机同比", growth(CHURN_M, CHURN_M_LY, better='down')],
               ["口径：在网用户", "÷ 住户数"]],
        y=y + 0.1, h=1.9, bottom=top - 0.2)       # 证据卡片填到底部结论条上方
    pg.source(SRC + "；虚线环节为待验证假设")


def p_compare(deck):
    """对比权衡：同一组维度逐项比较，推荐方案整列高亮，结论写清「选什么、为什么、代价是什么」。"""
    pg = deck.page("二、选什么方案：「提速+权益」续约综合最优")
    pg.lead("三个方案中，「提速+权益」续约在续约率与收入之间取得最好平衡，代价是需协调提速资源")
    pg.compare(
        ["降价续约", "提速+权益续约", "融合捆绑续约"],
        ["续约率提升", "收入影响", "实施难度", "见效周期"],
        [[stars(4) + "\n预计 +9pp", stars(4) + "\n预计 +8pp", stars(5) + "\n预计 +11pp"],
         [stars(1) + "\nARPU 约降 8%", stars(4) + "\nARPU 基本持平", stars(3) + "\n短期让利"],
         [stars(5) + "\n资费调整即可", stars(4) + "\n需协调提速资源", stars(2) + "\n需改造产品与系统"],
         [stars(5) + "\n1 个月", stars(4) + "\n1–2 个月", stars(2) + "\n3 个月以上"]],
        recommend=1, crit_header="评价维度\n（★越多越优）", crit_w=1.6,
        verdict="<b>结论：</b>推荐「提速+权益续约」——续约率提升接近最优，且收入基本不受损。"
                "<b>代价：</b>需网络部协调提速资源，见效比降价续约晚约 1 个月；融合捆绑可作为下一阶段方案。")
    pg.source(SRC + "；★为三个方案之间的相对评价（初评），预计值为测算，需试点验证")


def p_matrix(deck):
    """优先级矩阵：横轴实施难度、纵轴业务价值；左上（高价值、低难度）优先推进。评分依据必须写出来。"""
    pg = deck.page("三、先做什么：优先推进三项高价值、低难度举措")
    y = pg.lead("七项候选举措按「业务价值 × 实施难度」评估，左上象限的三项本月启动，右上两项纳入四季度规划")
    (mx, mw), (rx, rw) = pg.cols([2.2, 1], gap=0.35)
    pg.matrix(mx, y, mw, B - y, "实施难度", "业务价值",
              [("优先推进", "高价值、低难度"), ("重点规划", "高价值、高难度"),
               ("顺手完成", "低价值、低难度"), ("暂缓", "低价值、高难度")],
              items=[("90天预警", 1.6, 4.6), ("续约激励", 2.2, 3.9), ("权益包上架", 2.5, 4.4),   # 点名用短名，全称写在右侧
                     ("装维质量整治", 3.9, 4.2), ("融合产品改造", 4.5, 3.6),
                     ("竞对资费监测", 1.7, 2.2), ("续约专属客服线", 3.8, 1.8)],
              highlight=0)
    # 右侧：与四个象限一一对应的行动安排（矩阵本身不是结论）
    y2 = pg.heading(rx, y, rw, "行动安排")
    actions = [("<red>优先推进 · 本月启动</red>", "到期前90天预警、提速权益包上架、客户经理续约激励", LIGHT_RED),
               ("<b>重点规划 · 四季度</b>", "装维质量整治、融合产品改造", LIGHT),
               ("<b>顺手完成</b>", "竞对资费监测纳入周报", LIGHT),
               ("<b>暂缓</b>", "续约专属客服线", LIGHT)]
    for (ay, ah), (t, d, fc) in zip(pg.rows([1.4, 1, 1, 1], y=y2, h=B - y2, gap=0.15), actions):
        pg.box(rx, ay, rw, ah, [t, d], fill=fc, line=None, size=13, anchor="m", line_spacing=1.3, margin=0.15)
    pg.source("评分依据（1–5分）：业务价值=对续约率的预计贡献；实施难度=资源投入与系统改造量；评分为初评（示例），需业务部门会审确认")


def p_cycle(deck):
    """流程闭环：只有材料写明「回到起点、持续迭代」时才用环形；单向流程用 chevrons()。"""
    pg = deck.page("四、怎么持续：建立月度续约运营闭环")
    pg.lead("从到期预警到策略优化形成闭环，每月复盘一次，用续约结果反过来调整权益与话术")
    pg.cycle([("到期预警", "到期前90天生成名单"), ("名单下发", "按网格派单到客户经理"),
              ("触达续约", "厅店、外呼、上门"), ("结果回传", "办理结果按日回流"),
              ("效果评估", "续约率、ARPU、成本"), ("策略优化", "调整权益包与话术")],
             center="按月复盘，持续迭代")
    pg.source("注：流程与频次为示例，需结合各区县实际确认")


def p_ending(deck):
    deck.ending()


PAGES = [p_cover, p_summary, p_causal, p_compare, p_matrix, p_cycle, p_ending]

if __name__ == "__main__":
    build(PAGES, OUT)
