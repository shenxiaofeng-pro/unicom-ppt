# 联通 PPT Skill（unicom-ppt）

让 Claude 按**中国联通官方母版**直接生成可编辑的 PPT：封面、目录页、红色标题线、Logo、页码都与母版一致，文字、表格、图表在 PowerPoint 和 WPS 里都能直接修改。

## 支持两类报告

| 类型 | 适用场景 | 特点 |
|---|---|---|
| **数据分析类** | 经营分析、月报 / 季报 / 年报、指标通报、对标排名 | 金字塔结构、结论先行；做得好的数字标**红** `#C00000`，做得差的标**绿** `#008650`；比率类指标的变化用 **pp** 表示 |
| **传统报告类** | 平台建设、专项攻坚、产品调整、数据赋能、管理汇报 | 三种页型：体系架构与业务价值 / 专项攻坚与策略落地 / 能力贯通与运营闭环 |

效果参考：`examples/preview/` 下的图片，或直接打开 `examples/*.pptx`。

## 安装

1. 把整个 `unicom-ppt` 文件夹复制到以下任一位置：
   - 个人使用（所有项目可用）：`~/.claude/skills/unicom-ppt/`
   - 只在某个项目里使用：`<项目目录>/.claude/skills/unicom-ppt/`
2. 安装依赖：`pip install python-pptx`
3. （可选，用于生成预览图）安装 Microsoft PowerPoint 或 LibreOffice；macOS 上可再执行 `brew install poppler`。

## 使用

在 Claude Code 里直接描述需求，并附上材料即可，例如：

- 「用联通模板，根据这份 Excel 做一份 9 月经营分析 PPT，8 页左右」
- 「把这段平台建设材料做成一页联通风格的体系架构页」
- 「做一份存量融合专项攻坚的汇报，包含问题、策略和执行保障」

Claude 会先判断报告类型、列出大纲，再生成 PPTX，并自动检查排版。

## 目录结构

```
unicom-ppt/
├── SKILL.md                  # Claude 读取的主说明
├── references/
│   ├── data-analysis.md      # 数据分析类规范（金字塔、好红差绿、pp 规则）
│   ├── business-report.md    # 传统报告类规范（三种页型）
│   └── api.md                # 工具库 API
├── scripts/
│   ├── unicom_ppt.py         # 生成工具库（基于母版）+ 静态检查
│   └── render.py             # 渲染为 PDF / PNG，用于检查
├── assets/                   # 联通母版、Logo、结束页图片
└── examples/                 # 两类报告的示例脚本、生成结果和预览图
```

## 说明

- 示例中的数据和业务内容均为虚构，仅用于演示版式。
- 母版更新时，替换 `assets/unicom_template.pptx`，并参照 `references/api.md` 末尾的「更换母版」一节核对。
