---
name: finance-report-analyzer
description: |
  Analyze financial data from uploaded Excel/PDF files and generate interactive HTML reports with charts and insights. Also supports web search for industry benchmarks and competitor data.
  Use when: (1) User uploads an Excel/CSV/PDF with financial data (income statement, balance sheet, cash flow), (2) User asks for financial analysis or company report, (3) User wants visual HTML reports from financial statements, (4) User mentions stock ticker + financial analysis, (5) User shares a Feishu sheet/doc link with financial data.
---

# Finance Report Analyzer

Generate interactive HTML financial analysis reports from uploaded Excel/PDF files, with optional web search for industry benchmarks.

## Workflow

### Step 1: Acquire the Data File

Supported sources (try in order):

1. **Feishu chat file attachment** — Extract file_key from message via Feishu API, download with `im/v1/messages/{msg_id}/resources/{file_key}?type=file`
2. **Feishu Sheet/Doc/Bitable link** — Use appropriate feishu_* tools
3. **Direct file path** — If user provides a local path
4. **User pastes data** — Parse from text

#### Feishu Chat File Download Procedure

When the source is a file attachment in Feishu chat:

```bash
# 1. Get tenant access token
TOKEN=$(curl -s -X POST 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal' \
  -H 'Content-Type: application/json' \
  -d '{"app_id":"APP_ID","app_secret":"APP_SECRET"}' | python3 -c "import json,sys; print(json.load(sys.stdin)['tenant_access_token'])")

# 2. Get message content to find file_key and file_name
curl -s "https://open.feishu.cn/open-apis/im/v1/messages/{message_id}" \
  -H "Authorization: Bearer $TOKEN" | python3 -m json.tool

# 3. Download the file
curl -s "https://open.feishu.cn/open-apis/im/v1/messages/{message_id}/resources/{file_key}?type=file" \
  -H "Authorization: Bearer $TOKEN" -o /tmp/financial_data.xlsx
```

Get app credentials from: `openclaw config get channels.feishu.appId` and `channels.feishu.appSecret` (read from openclaw.json directly if redacted).

### Step 2: Parse the Data

Use `openpyxl` for Excel files. Install if needed: `pip install openpyxl`

```python
import openpyxl
wb = openpyxl.load_workbook('/tmp/file.xlsx', data_only=True)
# Iterate sheets, extract row labels and values
# Identify annual reports vs quarterly by checking report period labels
```

Key parsing rules:
- Identify period columns (年报/季报/中报/盈利预测)
- For annual trend analysis, filter to 年报 columns only (exclude 下半年报)
- Include 盈利预测 columns as forecast data
- Build a dict of {metric_label: [values_per_period]}

For PDF: use `pdfplumber` or `tabula-py` to extract tables first.

### Step 3: Analyze and Generate Report

Run `scripts/generate_report.py` with the parsed data. The script generates a self-contained HTML report with:

1. **Executive Summary** — Key metrics cards (revenue, profit, margins, forecasts)
2. **Profitability Analysis** — Revenue, profits, margins, ROE/ROA trends
3. **Growth Analysis** — Revenue growth, employee productivity, CAGR
4. **Balance Sheet Analysis** — Assets, liabilities, equity, leverage ratios
5. **Cash Flow Analysis** — OCF, ICF, FCF, free cash flow, cash position
6. **Valuation & Efficiency** — EPS, BPS, P/E, P/B, asset turnover
7. **Risk & Outlook** — Risk factors and positive catalysts

### Step 4: Web Search Enhancement (Optional)

When user requests industry comparison or competitor benchmarking:

```
web_search("百济神州 行业对比 BTK抑制剂 市场份额")
web_search("{company} financial results {year}")
```

Add findings to the report's benchmark section.

### Step 5: Deliver the Report

1. Save HTML to workspace
2. Send file via message tool to user's chat
3. Provide a text summary of key findings

## Report Design Guidelines

- Self-contained HTML (no external dependencies)
- Responsive design with CSS Grid
- Color coding: green=positive, red=negative
- Forecast columns highlighted with yellow background
- Sticky first column in tables for scrollability
- Summary boxes after each section with key takeaways
- Mobile-friendly layout

## Metric Definitions

See [references/metrics.md](references/metrics.md) for financial metric calculations and definitions.
