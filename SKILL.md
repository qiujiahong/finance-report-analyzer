---
name: finance-report-analyzer
description: |
  Analyze financial data from uploaded Excel/PDF files and generate interactive reports with sparkline trend charts. Supports output to PDF, DOCX, Markdown, and HTML.
  Use when: (1) User uploads Excel/CSV/PDF with financial data, (2) User asks for financial analysis or company report, (3) User wants visual reports from financial statements, (4) User mentions stock ticker + financial analysis, (5) User shares a Feishu sheet/doc link with financial data.
---

# Finance Report Analyzer

Generate financial analysis reports from uploaded Excel/PDF files with LLM-powered analysis, web search for news, and inline SVG sparkline trend charts.

## Architecture

The tool works in a **hybrid pipeline**:

1. **Python script** → Extract data, build tables/charts, render HTML/PDF (deterministic, zero-cost)
2. **LLM (OpenClaw agent)** → Write in-depth analysis text for each section + news summary (intelligent, contextual)
3. **Web search** → Fetch recent company news and hot topics

## Workflow (OpenClaw Agent)

### Step 1: Acquire Data File

Try in order:

1. **Feishu chat file attachment** — Download via API:
   ```bash
   # Get token
   TOKEN=$(curl -s -X POST 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal' \
     -H 'Content-Type: application/json' \
     -d '{"app_id":"APP_ID","app_secret":"APP_SECRET"}' | python3 -c "import json,sys; print(json.load(sys.stdin)['tenant_access_token'])")
   # Get file_key from message
   curl -s "https://open.feishu.cn/open-apis/im/v1/messages/{message_id}" -H "Authorization: Bearer $TOKEN"
   # Download
   curl -s "https://open.feishu.cn/open-apis/im/v1/messages/{message_id}/resources/{file_key}?type=file" \
     -H "Authorization: Bearer $TOKEN" -o /tmp/data.xlsx
   ```

2. **Feishu Doc/Bitable link** — Use feishu_doc/feishu_bitable tools
3. **Local file** — Use directly

### Step 2: Extract Financial Data (JSON)

```bash
python3 scripts/generate_report.py /tmp/data.xlsx --company "公司名" --ticker "000001.SZ" --output-dir /tmp/reports --json
```

This outputs structured JSON with all financial metrics organized by category.

### Step 3: Web Search for Company News

```
web_search("{company} {year} 最新新闻 业绩 研发 重大事件")
```

Summarize 5-8 key news items as HTML list. Save to `/tmp/reports/news.html`.

### Step 4: LLM Analysis

Based on the JSON data from Step 2, write professional analysis for 6 sections. Save as JSON to `/tmp/reports/analysis.json`:

```json
{
  "profitability": "HTML text - revenue trends, margins, profitability inflection points...",
  "balance_sheet": "HTML text - asset structure, leverage, liquidity...",
  "cash_flow": "HTML text - operating CF trends, capex, FCF, cash reserves...",
  "per_share": "HTML text - EPS/BPS trends, efficiency ratios, workforce...",
  "industry": "HTML text wrapped in <div class=\"analysis-box\">...",
  "risk": "HTML text using <div class=\"two-col\"><div class=\"col\">... layout"
}
```

**Analysis guidelines:**
- Use `<strong>` for key conclusions as the opening line
- Use `<br><br>` for paragraph breaks
- Reference specific numbers from the data
- Compare year-over-year trends
- Highlight inflection points and turning points
- For risk section, use the two-column float layout with `.risk-list`

### Step 5: Generate Final Report

```bash
python3 scripts/generate_report.py /tmp/data.xlsx \
  --company "公司名" --ticker "000001.SZ" \
  --output-dir /tmp/reports \
  --analysis-json /tmp/reports/analysis.json \
  --news-html /tmp/reports/news.html \
  -o html,pdf
```

### Step 6: Deliver Files via Feishu API

```bash
# Upload and send file
FK=$(curl -s -X POST 'https://open.feishu.cn/open-apis/im/v1/files' \
  -H "Authorization: Bearer $TOKEN" \
  -F 'file_type=stream' -F "file_name=report.html" -F "file=@/tmp/reports/report.html" \
  | python3 -c "import json,sys; print(json.load(sys.stdin)['data']['file_key'])")

curl -s -X POST 'https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id' \
  -H "Authorization: Bearer $TOKEN" -H 'Content-Type: application/json' \
  -d "{\"receive_id\":\"CHAT_ID\",\"msg_type\":\"file\",\"content\":\"{\\\"file_key\\\":\\\"$FK\\\"}\"}"
```

## CLI Reference

```bash
# JSON data export (for LLM pipeline)
python3 scripts/generate_report.py input.xlsx --json --company NAME --ticker TICKER --output-dir DIR

# Full report with LLM analysis
python3 scripts/generate_report.py input.xlsx \
  --analysis-json analysis.json \
  --news-html news.html \
  --company NAME --ticker TICKER \
  -o html,pdf --output-dir DIR

# Standalone (rule-based analysis, no LLM needed)
python3 scripts/generate_report.py input.xlsx --company NAME --ticker TICKER -o html,pdf
```

## Output Formats

| Flag | Output | Requires |
|------|--------|----------|
| `-o html` | HTML only | (built-in) |
| `-o pdf` | HTML + PDF | wkhtmltopdf |
| `-o doc` | HTML + DOCX | pandoc |
| `-o md` | HTML + Markdown | pandoc or markdownify |

## Report Features

- **LLM-powered analysis**: Deep, contextual financial commentary (when using --analysis-json)
- **Company news section**: Recent hot topics and events (when using --news-html)
- **Sparkline trend charts**: SVG mini-charts in each data row
- **Forecast markers**: Predicted values with ⟡ symbol and yellow background
- **PDF compatible**: No emoji/flex/gradient/CSS variables (works with wkhtmltopdf)
- **Fallback mode**: Rule-based analysis when no LLM analysis is provided
