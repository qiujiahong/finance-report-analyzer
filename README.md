# 📊 Finance Report Analyzer

从 Excel 财务数据自动生成专业分析报告，结合大模型深度分析 + 新闻搜索 + SVG 趋势图。

[![ClawHub](https://img.shields.io/badge/ClawHub-finance--report--analyzer-blue)](https://clawhub.com)
[![GitHub](https://img.shields.io/badge/GitHub-qiujiahong%2Ffinance--report--analyzer-black)](https://github.com/qiujiahong/finance-report-analyzer)

## ✨ 功能特性

- 🤖 **大模型深度分析** — 盈利能力、资产负债、现金流等章节由 LLM 生成专业分析
- 🏭 **竞争对手分析** — 自动分析关键竞争对手及潜在威胁
- 📰 **热点新闻整合** — 搜索公司近期新闻，整合进报告
- 📈 **SVG 迷你趋势图** — 每个指标行内嵌 sparkline
- 🔮 **预测值标识** — 盈利预测数据用 ⟡ 符号和黄色背景标注
- 📄 **多格式输出** — HTML + PDF / DOCX / Markdown
- 📱 **PDF 兼容** — 无 emoji/flex/gradient，完美支持 wkhtmltopdf

## 🏗️ 架构

采用**混合管线**设计：

```
Excel 数据 → Python 脚本（提数据/建表格/渲染）
                ↓
           --json 输出结构化数据
                ↓
         LLM 分析 + Web 搜索新闻
                ↓
         --analysis-json + --news-html 注入
                ↓
           生成最终 HTML + PDF 报告
```

| 组件 | 职责 | 成本 |
|------|------|------|
| Python 脚本 | 数据提取、表格、趋势图、排版 | 零 |
| LLM（OpenClaw） | 深度分析文字、竞争对手、新闻总结 | 按模型计费 |
| Web 搜索 | 公司新闻、竞争对手信息 | 免费 |

## 🚀 快速开始

### 安装

```bash
# 作为 OpenClaw Skill 安装（推荐）
clawhub install finance-report-analyzer

# 或直接克隆
git clone https://github.com/qiujiahong/finance-report-analyzer.git
pip install openpyxl
```

### 方式一：飞书群使用（推荐）

1. 上传财务 Excel 文件到群里
2. @机器人 说「帮我生成财务分析报告」
3. 自动完成：下载 → 提数据 → LLM 分析 → 搜索新闻 → 生成 HTML+PDF

### 方式二：LLM 增强模式（命令行）

```bash
# Step 1: 提取结构化数据
python3 scripts/generate_report.py data.xlsx --json \
  --company "百济神州-U" --ticker "688235.SH" --output-dir /tmp/reports

# Step 2: 准备 LLM 分析（analysis.json）和新闻（news.html）
# 详见 SKILL.md 中的分析指南

# Step 3: 生成最终报告
python3 scripts/generate_report.py data.xlsx \
  --company "百济神州-U" --ticker "688235.SH" \
  --analysis-json /tmp/reports/analysis.json \
  --news-html /tmp/reports/news.html \
  --output-dir /tmp/reports -o html,pdf
```

### 方式三：纯脚本模式（无需 LLM）

```bash
python3 scripts/generate_report.py data.xlsx \
  --company "公司名" --ticker "000001.SZ" -o html,pdf
```

自动使用规则引擎生成分析文字，零成本、可离线。

### 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `input` | 输入 Excel 文件路径 | (必填) |
| `-o, --output-format` | 逗号分隔格式：`html,pdf,doc,md` | `html,pdf` |
| `--output-dir` | 输出目录 | `.` |
| `--company` | 公司名称 | (空) |
| `--ticker` | 股票代码 | (空) |
| `--json` | 输出结构化 JSON 数据（供 LLM 分析） | (flag) |
| `--analysis-json` | LLM 分析文本 JSON 文件路径 | (空) |
| `--news-html` | 新闻 HTML 文件路径 | (空) |

### 格式转换依赖

| 输出格式 | 需要安装 |
|----------|----------|
| HTML | 无（内置） |
| PDF | `wkhtmltopdf` |
| DOCX | `pandoc` |
| Markdown | `pandoc` 或 `pip install markdownify` |

## 📋 报告内容（8 章）

| 章节 | 内容 | 数据来源 |
|------|------|----------|
| 核心摘要 | 营收、净利润、毛利率、ROE 及同比 | 脚本 |
| 一、盈利能力分析 | 收入利润趋势 + 利润率指标 | 脚本(表格) + LLM(分析) |
| 二、资产负债分析 | 资产/负债/权益结构 | 脚本(表格) + LLM(分析) |
| 三、现金流分析 | 经营/投资/筹资现金流 + 自由现金流 | 脚本(表格) + LLM(分析) |
| 四、每股指标与效率 | EPS/BPS/周转率/员工 | 脚本(表格) + LLM(分析) |
| 五、行业分析 | 行业定位与对标 | LLM |
| 六、关键竞争对手分析 | 竞品表格 + 战略分析 | LLM |
| 七、风险与机遇分析 | 风险/积极因素双栏 | LLM |
| 八、近期热点新闻 | 公司近期要闻 | Web 搜索 + LLM |

### analysis.json 格式

```json
{
  "profitability": "盈利能力分析 HTML...",
  "balance_sheet": "资产负债分析 HTML...",
  "cash_flow": "现金流分析 HTML...",
  "per_share": "每股指标分析 HTML...",
  "industry": "行业分析 HTML...",
  "competitor": "竞争对手分析 HTML（含表格）...",
  "risk": "风险与机遇分析 HTML（双栏布局）..."
}
```

## 📁 项目结构

```
finance-report-analyzer/
├── SKILL.md                          # OpenClaw Skill 工作流说明
├── README.md                         # 本文件
├── scripts/
│   └── generate_report.py            # 报告生成脚本
└── references/
    └── metrics.md                    # 财务指标定义参考
```

## 🔧 数据格式要求

输入 Excel 文件应包含：
- 第一列为指标名称（如"营业总收入"、"净利润"等）
- 后续列为各报告期数据
- 支持 Wind / Choice / 同花顺 导出格式

支持的指标（自动匹配）：

| 类别 | 关键词 |
|------|--------|
| 收入 | 营业总收入, 营业收入 |
| 利润 | 营业利润, 净利润, 归母净利润 |
| 资产 | 资产总计, 流动资产 |
| 负债 | 负债合计, 资产负债率 |
| 现金流 | 经营活动现金净流量, 资本开支 |
| 比率 | 销售毛利率, ROE, ROA |
| 每股 | EPS, 每股净资产 |
| 其他 | EBITDA, 研发支出, 员工总数 |

## 📝 License

MIT
