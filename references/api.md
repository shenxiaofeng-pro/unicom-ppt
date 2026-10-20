# unicom_ppt API 参考

```python
import sys; sys.path.insert(0, "SKILL_DIR/scripts")
from unicom_ppt import *
```

坐标单位一律为**英寸**（页面 13.333 × 7.5）。所有方法返回创建的 shape，另有说明的除外。

## 版面常量

| 常量 | 值 | 含义 |
|---|---|---|
| `L` `R` | 0.30 / 13.03 | 内容区左、右边界 |
| `T` `B` | 1.05 / 6.85 | 内容区上边界（标题线下方）、下边界（页码和来源上方） |
| `W` `H` | 12.73 / 5.80 | 内容区宽、高 |

颜色常量：`RED`(=`GOOD`) `GREEN`(=`BAD`) `BLACK` `TEXT` `DARK` `GRAY` `MID` `LINE` `LIGHT` `WARM` `LIGHT_RED` `PINK` `BLUE` `LIGHT_BLUE` `WHITE`，色值见 SKILL.md。
`SERIES = [RED, MID, DARK, PINK, BLUE]` 是多系列图表的默认配色。

## 文字标记

以下参数支持标记：`text / bullets / box / label / lead / conclusions` 的文字，`table` 的单元格，`kpi` 的 `value / change / note`。
**不支持**标记的参数：`kpi` 的 `label`、图表的类别名和系列名、`deck.page()` 的标题、`toc()` 的章节名。

```
<b>加粗</b>   <good>好·红</good>   <bad>差·绿</bad>   <red>强调红</red>
<green>绿</green>   <blue>蓝</blue>   <gray>灰</gray>
```

- 标记可以嵌套。分段用 `"\n"`，或直接传字符串列表。
- 在 `bullets()` 中，以 `"- "` 开头的项为二级条目。
- `parse_markup(text)` 把标记文本解析为 `[(文字, 样式)]`；`strip_markup(text)` 去掉所有标记，返回纯文本。

## 数字格式化

| 函数 | 作用 | 例子 → 输出 |
|---|---|---|
| `num(v, unit="", digits=1, sign=False)` | 纯文本数字，不着色 | `num(-1.25, "pp", sign=True)` → `−1.3pp`；`num(38.4, "%")` → `38.4%` |
| `growth(new, old, better="up", digits=1, prefix="")` | 绝对量的变化率（%），着色 | `growth(12.6, 11.98)` → `<good>+5.2%</good>` |
| `pp(new, old, better="up", digits=1, prefix="")` | 两个百分比之差（pp），着色 | `pp(41.2, 42.5)` → `<bad>−1.3pp</bad>` |
| `delta(v, unit="%", better="up", digits=1, prefix="")` | 已知变化值，直接格式化并着色 | `delta(-0.6, "pp")` → `<bad>−0.6pp</bad>` |
| `vs_target(actual, target, text=None, better="up", digits=1, unit="")` | 给实际值本身按是否达标着色 | `vs_target(75.2, 75, unit="%")` → `<good>75.2%</good>` |
| `mark(text, good)` | 手动着色：`True` 红、`False` 绿、`None` 只加粗 | `mark("第1名", True)` |

- `better`：`"up"` 表示越高越好；`"down"` 表示越低越好（离网、投诉、成本、拆机）；`None` 表示不判断好坏，只加粗。传入其他值会报错。
- 变化值四舍五入后为 0 时输出「0.0」，不带符号、只加粗。
- 负号统一用数学减号「−」，PowerPoint 不会在它后面断行。**不要用 `f"{x:+.1f}"` 自己拼正负号。**

## UnicomDeck

| 方法 | 说明 |
|---|---|
| `UnicomDeck(template=TEMPLATE)` | 加载母版并清空示例页 |
| `cover(title, org="", date="", subtitle=None)` | 封面。标题超过 18 字时自动缩小字号并换行 |
| `toc(items, active=0)` | 目录 / 章节页，2–6 项，`active` 项标红 |
| `page(title, size=30)` | 内容页，返回 `Page`。标题过长时自动缩到最小 22pt，并给出警告（建议不超过 26 字） |
| `ending(text=None)` | 结束页：默认用母版图片；传入文字（如 `"谢谢！"`）则生成「Logo + 红线 + 文字」样式 |
| `save(path, check=True)` | 保存并做静态检查，打印并返回问题列表 |

## Page：布局辅助

| 方法 | 说明 |
|---|---|
| `lead(text, x=L, y=None, w=W, size=15, bold=True, fill=LIGHT, h=None)` | 页首结论条（左侧红色竖条 + 浅灰底，高度随文字自适应）。**返回下方内容的起始 y**。传统报告类的「总述」用 `bold=False, size=14` |
| `top` | 属性：下一个可用的 y（调用 `lead()` 后会自动下移） |
| `cols(ratios, x=L, w=W, gap=0.25)` | 按比例切分列，例如 `cols([1, 2, 1])` → `[(x, w), …]`；传整数 n 表示 n 等分 |
| `rows(ratios, y=None, h=None, gap=0.2)` | 按比例切分行 → `[(y, h), …]`；默认从 `top` 一直分到 `B` |

## Page：文字与容器

| 方法 | 说明 |
|---|---|
| `text(x, y, w, h, text, size=12, color=TEXT, bold=False, align="l", anchor="t", line_spacing=1.15, space_after=0, bullets=None, fill=None, line=None, margin=0.05, fit=True, min_size=10, sizes=None)` | 文本框。`fit=True` 时放不下会自动缩小字号，最小到 `min_size`；`sizes` 可逐段指定字号 |
| `bullets(x, y, w, h, items, size=13, color=TEXT, line_spacing=1.25, space_after=4, style="square", **kw)` | 项目符号列表，用母版的红色□；`style="dot"` 为灰色圆点。`**kw` 会传给 `text()` |
| `heading(x, y, w, text, size=15, color=BLACK, rule=False)` | 模块小标题（红色□ + 加粗），**返回下方可用 y** |
| `box(x, y, w, h, text="", fill=None, line=LINE, line_w=0.75, color=TEXT, size=12, bold=False, align="l", anchor="t", shape="rect", dash=False, margin=0.08, line_spacing=1.15, space_after=0, bullets=None, fit=True, min_size=9, shadow=False)` | 容器框。`shape` 可选 rect / round / pill；`line=None` 表示无边框 |
| `label(x, y, w, h, text, fill=RED, color=WHITE, size=13, bold=True, align="c", shape="rect", **kw)` | 色块标题条（无边框、垂直居中） |
| `kpi(x, y, w, h, label, value, unit="", change=None, note=None, judge=None, value_size=28, fill=LIGHT, align="l")` | 指标卡。`value` 用 `num()` 生成；`unit` 为 `%` / `pp` 时紧贴数字、字号相同，其他单位（亿元、万户）以小字号跟在后面；`change` 如 `"同比 " + growth(a, b)`；`judge` 取 `"good"` / `"bad"` / `None` |
| `conclusions(items, x=L, y=None, w=W, h=None, gap=0.18, value_w=2.4, head_size=16, body_size=13)` | 核心结论 / 要点列表。`items = [(标题, 证据, 数字, 数字标签, judge), …]`，后三项可省略。默认占满 lead 以下的空间，**返回下方 y** |
| `badge(x, y, d, text, fill=RED, color=WHITE, size=12)` | 圆形序号徽标 |
| `source(text, size=9)` | 左下角的来源 / 口径说明 |
| `notes(text)` | 演讲者备注 |
| `image(path, x, y, w=None, h=None)` | 插入图片（截图证据）；只给 w 或只给 h 时按原比例缩放 |

`align` 可传 `l / c / r / j`，也可传 `left / center / right / justify`；`anchor` 可传 `t / m / b`，也可传 `top / middle / bottom`。传入其他值会报出明确的错误。

## Page：关系与流程

| 方法 | 说明 |
|---|---|
| `line(x1, y1, x2, y2, color=MID, width=1.0, dash=False, head=None, tail=None)` | 直线；`head` / `tail` 可设为 triangle / arrow / oval（tail 是终点端） |
| `arrow(x1, y1, x2, y2, color=MID, width=1.25, dash=False, both=False, text=None, text_size=10, text_color=DARK, text_pos="side")` | 箭头连线，终点带箭头，`both=True` 为双向。`text` 为连线标签：`text_pos="side"` 时竖线标签放在线右侧、横线标签放在线上方，不遮挡箭头；`"center"` 时白底压在线中间。起止点要正好落在框边上 |
| `loop(src, dst, top, text=None, depth=0.3, color=RED, width=1.25, text_size=11)` | 反馈闭环：从 `src` 列底部引出红色虚线，沿下方横线回到 `dst` 列底部（带箭头）。`src` / `dst` 为 `(x, w)`，通常取自 `chevrons()` 的返回值；`top` 为内容底边。**返回下方 y** |
| `block_arrow(x, y, w, h, direction="right", fill=PINK, text="", color=TEXT, size=11)` | 粗箭头，方向可选 right / down / left / up |
| `chevrons(x, y, w, h, steps, fill=RED, color=WHITE, size=14, gap=0.25, active=None, inactive=MID, ratios=None)` | 横向阶段链。`ratios` 可指定各列宽度比例（默认等宽）。**返回与各节点对齐的内容列** `[(x, w), …]`，可直接用来在下方排布内容 |

## Page：表格与图表（原生、可编辑）

| 方法 | 说明 |
|---|---|
| `table(x, y, w, h, rows, col_widths=None, header=True, size=11, header_size=None, header_fill=RED, header_color=WHITE, zebra=True, align=None, first_col_bold=False, highlight_rows=None, highlight_fill=LIGHT_RED)` | `rows` 为二维列表，单元格支持标记；`col_widths` 是相对比例；`align` 逐列指定（默认首列左对齐、其余居中）。行高按 h 平均分配，内容多时 PowerPoint 会把行撑高，`save()` 会提示 |
| `bar(x, y, w, h, categories, series, colors=None, highlight=None, number_format="0.0", horizontal=False, stacked=False, labels=True, legend=None, size=10, gap_width=70, overlap=None, axis=False, label_color=TEXT)` | 柱状 / 条形图。`series` 可传列表、字典或 `[(name, vals)]`；`highlight={类别名或索引: "good" / "bad" / "neutral" / 色值}`，只作用于第一个系列；全为正值时纵轴从 0 开始；有负值时类别标签自动放到轴外侧 |
| `line_chart(x, y, w, h, categories, series, colors=None, number_format="0.0", labels=True, legend=None, size=10, smooth=False, axis=False, label_last_only=False)` | 折线图。多系列时第 1 条线的标签在上方、其余在下方 |
| `doughnut(x, y, w, h, categories, values, colors=None, number_format="0%", size=10, legend=True, hole=55)` | 环形图，类别不超过 5 个；`values` 传占比小数或绝对值均可 |

`legend=None` 表示多系列时自动显示图例。`number_format` 是 Excel 格式码，例如 `"0.0"`、`"0.00"`、`"#,##0"`、`"0%"`。

## 检查与渲染

```bash
python SKILL_DIR/scripts/unicom_ppt.py deck.pptx      # 对任意 pptx 做静态检查（save() 也会自动运行）
python SKILL_DIR/scripts/render.py deck.pptx [--out 目录] [--dpi 110]   # 导出 PDF 和 slide-XX.png
```

静态检查的项目：
- 构建时：文字放不下、字号被自动缩小 1pt 以上（同组文字可能大小不一）、标题过长、表格超高
- 保存后：越界、文字溢出（估算）、段落末行孤字（估算）、字号小于 9pt、残留占位文字（请输入 / 待补充 / XX / xxx）、未解析的标记标签、文字框或图表之间部分重叠
检查是估算，**最终以渲染图为准**；行内留白过多这类问题只能靠目检发现。

## 更换母版

联通发布新母版时，替换 `assets/unicom_template.pptx`，然后核对：

1. `unicom_ppt.py` 顶部的三个版式名称常量 `LAYOUT_COVER / LAYOUT_TOC / LAYOUT_CONTENT`。
2. `cover()` / `toc()` / `page()` 中的坐标是否与新母版一致。
3. `assets/unicom_logo.png`、`assets/ending.jpg` 是否需要同步更新。
