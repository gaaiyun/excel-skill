# Excel Skill

不是 Excel 教程，是 **AI 友好的 Excel 知识库 + 4 大行业模板 + Python 自动化工具**。

让 Cursor / Claude Code / Codex 等 AI agent 能根据你的需求**生成可直接用的 Excel 文件 / 公式 / 透视表 / Dashboard**，而不是查文档查到崩溃。

---

## 这个项目要解决什么

我（项目作者）和很多金融/快消/电商/互联网从业者一样，工作离不开 Excel，但又不是 Excel 重度老用户：

- 知道有公式但记不住具体语法（`SUMIFS` 还是 `SUMIF`？`INDEX/MATCH` 还是 `XLOOKUP`？）
- 知道 Power Query 牛但学习曲线陡（界面操作 + M 语言）
- 想用 Python `openpyxl` 自动化但每次都查文档
- 行业里大家都在用某些"约定俗成"的报表样式（金融三表、电商漏斗、快消 RTM 等），但 Google 出来的模板大多不对版

让 AI 帮我处理这些是最快的——但 AI 经常：
- 编错的公式参数（"`VLOOKUP` 第三个参数应该是 0 还是 FALSE"）
- 用 `pandas.read_excel + to_excel` 时**把所有公式变成字符串**（最经典的坑）
- 生成的 openpyxl 代码在「formula 在 openpyxl 里只是字符串，需要重算」这点上踩坑

这个项目就是给 AI 一份**结构化知识库 + 行业模板库 + 一组可复用脚本**，让它一次给对。

---

## 它和已有项目的关系

| 项目 | Stars | 我们怎么用 |
|---|---|---|
| [Anthropic 官方 xlsx skill](https://github.com/anthropics/skills) | - | 通用 xlsx 操作的标准实现，我们 reference 它 |
| [`claude-office-skills/skills`](https://github.com/claude-office-skills/skills) | - | 通用 openpyxl 知识，我们做行业化补充 |
| [`davepoon/buildwithclaude`](https://github.com/davepoon/buildwithclaude) | 3K | 含 `recalc.py` 公式重算，我们 fork 思路 |
| [`sartrus/modelling-team-skill`](https://github.com/sartrus/modelling-team-skill) | - | 三 agent 团队建财务模型，我们参考它的 Architect/Coder/Challenger 分层 |
| [`LondonMarket/Financial-Model-Excel-Template`](https://github.com/LondonMarket/Financial-Model-Excel-Template) | 9 | DCF 模板，我们的金融行业部分 reference |
| 熊猫办公 / 各模板站 | - | 中文社区主流模板，我们梳理"哪些场景用哪些模板" |

**核心差异**：
- **行业垂直**——金融、快消、电商、互联网 4 大领域，每个都有专属模板和 KPI 字典
- **中文优先**——文档、注释、模板字段都是中文
- **AI 友好**——不是给人读的 Excel 教程，是给 LLM 读的 progressive disclosure 知识库

---

## 4 大行业覆盖

### 💰 金融（Finance）
- 三表（资产负债 / 利润 / 现金流）联动模板
- DCF 估值模型
- 股票投资组合跟踪
- 财务比率分析（ROE 杜邦分解）

### 🛒 快消（FMCG / CPG）
- 销售达成 vs 目标分析
- RTM (Route-to-Market) 渠道分析
- 经销商进销存表
- SKU ABC 分类 + Pareto 分析
- 促销活动 ROI 计算

### 🛍️ 电商（E-commerce）
- GMV / DAU / 转化率 / 客单价 仪表盘
- 销售漏斗（曝光 → 点击 → 加购 → 下单 → 支付）
- ROI / ROAS 投放分析
- 用户分层（RFM 模型）
- SKU 销售排名 + 长尾分析

### 🌐 互联网（Internet / SaaS）
- DAU / MAU / 留存曲线
- 漏斗分析（注册 → 激活 → 留存 → 付费）
- LTV / CAC / 回本周期
- A/B 测试结果分析
- 增长指标 dashboard

---

## 三大能力

### 1. AI 友好的 Excel 知识库（progressive disclosure）

```
references/
├── 01-formulas-cheatsheet.md       # 公式速查（按场景而非字母顺序）
├── 02-pivot-tables.md              # 透视表常用模式
├── 03-power-query.md               # Power Query 入门 + M 语言常用片段
├── 04-charts-and-dashboards.md     # 图表选择 + Dashboard 设计原则
├── 05-openpyxl-python.md           # Python 自动化生成 Excel
├── 06-pandas-excel.md              # pandas 读写 Excel 常见坑
├── 07-vba-quick-reference.md       # VBA 速查（仅必要时）
└── 08-excel-pitfalls.md            # ★ 最重要：AI / 人最容易踩的坑
```

`SKILL.md` 是入口，按用户问题路由到对应的 reference。

### 2. 4 大行业模板库

```
templates/
├── finance/                  # 金融
│   ├── three-statements.xlsx          # 三表联动
│   ├── dcf-valuation.xlsx             # DCF 估值
│   ├── portfolio-tracker.xlsx         # 持仓跟踪
│   └── financial-ratios.xlsx          # 财务比率
├── fmcg/                     # 快消
│   ├── sales-vs-target.xlsx
│   ├── rtm-analysis.xlsx
│   ├── distributor-inventory.xlsx
│   └── promotion-roi.xlsx
├── ecommerce/                # 电商
│   ├── gmv-dashboard.xlsx
│   ├── sales-funnel.xlsx
│   ├── roas-analysis.xlsx
│   └── rfm-segmentation.xlsx
└── internet/                 # 互联网
    ├── dau-mau-cohort.xlsx
    ├── activation-funnel.xlsx
    ├── ltv-cac.xlsx
    └── ab-test-analysis.xlsx
```

每个模板都用 Python `openpyxl` 脚本生成（见 `scripts/generate_*.py`），不是手工 .xlsx 二进制——意味着你可以：
- 看 .py 脚本知道结构怎么搭
- 改参数（行数、字段名）重新生成
- 复用样式逻辑做新模板

### 3. Python 工具脚本

```
scripts/
├── excel_lint.py              # 检查 AI 生成的 openpyxl 代码常见坑
├── recalc.py                  # 用 LibreOffice 重算公式（fork from buildwithclaude）
├── generate_finance/          # 4 个行业的模板生成脚本
├── generate_fmcg/
├── generate_ecommerce/
├── generate_internet/
└── helpers/
    ├── styling.py             # 通用样式（标题色 / 千分位 / 百分比 / 边框）
    ├── data_validation.py     # 数据验证（下拉列表 / 范围限制）
    └── formulas.py            # 常用公式生成器
```

---

## 30 秒上手

### 不用 IDE，命令行直接生成

```bash
git clone https://github.com/gaaiyun/excel-skill.git
cd excel-skill
pip install -r requirements.txt

# 生成一个金融三表模板
python scripts/generate_finance/three_statements.py --output my_company.xlsx

# 生成一个电商 GMV dashboard 模板
python scripts/generate_ecommerce/gmv_dashboard.py --output gmv.xlsx
```

### 在 Cursor / Claude Code 里直接说

```
用 excel-skill 帮我做一个电商 SKU 销售排名 + ABC 分析的模板，输入 100 个 SKU
```

```
我有一份销售数据，按月份和产品类目，帮我用 Power Query 做一个透视分析
```

```
帮我审一下这段 openpyxl 代码（贴代码），看有没有公式被字符串化的问题
```

---

## 学习路径建议

### 完全新手 → 中级
1. 读 [`tutorials/01-excel-basics.md`](./tutorials/01-excel-basics.md) — Excel 必备 80 个快捷键 + 10 个核心公式
2. 看一个行业模板的 .py 脚本，理解 openpyxl 怎么用
3. 跟着 [`tutorials/02-pivot-tables-101.md`](./tutorials/02-pivot-tables-101.md) 做透视表

### 中级 → 进阶
4. 读 [`references/03-power-query.md`](./references/03-power-query.md) — Power Query 替代 80% 的复杂公式
5. 跟着 [`tutorials/03-dashboard-design.md`](./tutorials/03-dashboard-design.md) 做交互式 Dashboard

### 进阶 → 自动化
6. 读 [`references/05-openpyxl-python.md`](./references/05-openpyxl-python.md) — 用 Python 批量生成 Excel
7. 把你的常用工作流封装成 `scripts/my_workflow.py`

---

## 项目结构

```
excel-skill/
├── README.md                  你正在看
├── SKILL.md                   Cursor / Claude Code skill 入口
├── WORKFLOW.md                跨场景工作流（清洗→分析→可视化→交付）
├── INSTALL_CN.md              Windows / 中文用户指南
├── LICENSE / .gitignore / requirements.txt / pyproject.toml
├── references/                8 个核心知识文档
├── templates/                 16 个行业 xlsx 模板（由脚本生成）
├── scripts/                   Python 自动化脚本
├── tutorials/                 4 个入门到进阶教程
└── examples/                  完整案例（如「双 11 GMV 分析全流程」）
```

---

## License

MIT。详见 [LICENSE](./LICENSE)。

---

## Credits

- **Anthropic 官方 xlsx skill** — openpyxl 操作标准
- **claude-office-skills / buildwithclaude** — 通用 Excel skill 设计参考
- **sartrus/modelling-team-skill** — 三 agent 财务建模启发
- **熊猫办公 / 各 Excel 模板站** — 中文行业模板参考
- **LondonMarket Financial-Model-Excel-Template** — DCF 模板参考
