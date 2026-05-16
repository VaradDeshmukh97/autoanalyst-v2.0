# AutoAnalyst v3 — Run Guide & Word Templating Strategy

---

## 1. Test run command

First, make sure your `.env` has:

```
OPENAI_API_KEY=sk-...
```

Then run:

```bash
python autoanalyst_v3.py \
    --company_name "ABM Industries Incorporated" \
    --short_name "ABM" \
    --ticker "ABM" \
    --links_xlsx links.xlsx \
    --source_model gpt-4o-mini \
    --report_model gpt-4o-mini \
    --max_sources 5 \
    --auto_accept_highlights \
    --output_dir test_output_ABM
```

**What each flag does here:**

| Flag | Value | Why |
|---|---|---|
| `--source_model gpt-4o-mini` | fast/cheap | Source notes don't need frontier model |
| `--report_model gpt-4o-mini` | fast/cheap | Swap to `gpt-4o` for real reports |
| `--max_sources 5` | capped | Test with 5 URLs, not all of them |
| `--auto_accept_highlights` | set | Skip interactive prompt for the test run |
| `--output_dir test_output_ABM` | named folder | Keeps test outputs separate |

**Production run (no caps, interactive highlights, frontier model):**

```bash
python autoanalyst_v3.py \
    --company_name "ABM Industries Incorporated" \
    --short_name "ABM" \
    --ticker "ABM" \
    --links_xlsx links.xlsx \
    --source_model gpt-4o-mini \
    --report_model gpt-4o \
    --output_dir output_ABM_v1
```

**Output files you'll find in the output directory:**

```
ABM_01_ingested_urls.json
ABM_02_knowledge_base.txt       ← per-source analytical notes
ABM_03_master_memo.txt          ← the synthesis layer
ABM_04a_thesis.json             ← title + thesis + variant perception
ABM_04b_highlights_preedit.json ← raw model highlights
ABM_04b_highlights_approved.json← post-analyst-edit highlights
ABM_05a_company_overview.txt
ABM_05b_highlight_1.txt ... _5.txt
ABM_05c_financials.txt
ABM_05d_risks.txt
ABM_06_report_draft.md
ABM_07_report_final.md          ← the clean final draft
ABM_08_report_payload.json      ← structured input for Word renderer
```

---

## 2. Word templating: docxtpl vs your current approach

### Short answer

Use **`docxtpl`** for everything the sidebar and cover page need (static fields:
company name, ticker, stats, date, analyst). Use **`python-docx`** directly (your
existing `word_render_v2.py` approach) for the body content, because the body is
not a template — it is dynamically generated prose that python-docx inserts
paragraph by paragraph.

The two tools do different jobs and they work together cleanly.

---

### Why docxtpl for the cover/sidebar

`docxtpl` uses Jinja2 `{{ variable }}` placeholders directly inside a `.docx` file.
You open the template in Word, type `{{ company_name }}` wherever you want the
company name, save, and docxtpl fills it at runtime. No XML surgery, no run-splitting
bugs, no placeholder-search-and-replace fragility.

**What it handles perfectly:**
- Company name, ticker, exchange in the header
- At-a-glance table: address, sector, website
- Key statistics table: price, market cap, EV, float
- Valuation table: P/E, EV/EBITDA, EV/Sales
- Report date, analyst name, analyst email
- Firm name in header and footer
- Report type label ("Initiation of Coverage")
- Rating label ("Buy / Hold / Sell" when you add one)

**What it cannot handle:**
- Dynamically generated body prose (the thesis, 5 sections, financials, risks)
- Content that varies in length across reports (each report has different paragraph counts)
- Repeating blocks (the 5 highlight sections are not template loops — they're full prose sections)

The body of your report is not a template. It is generated content. docxtpl's
`{% for %}` loops work for tabular data, not for multi-paragraph analytical prose
with varying structure per section.

---

### Recommended architecture: two-pass rendering

```
ABM_08_report_payload.json
        │
        ▼
  Pass 1: docxtpl
  ─────────────────────────────────────────────────────
  Input:  template_blank.docx   (has {{ }} placeholders
                                  in cover, sidebar, header, footer)
  Output: report_populated.docx (cover page + sidebar filled,
                                  body still has [[REPORT_BODY]] marker)
        │
        ▼
  Pass 2: python-docx (word_render_v3.py)
  ─────────────────────────────────────────────────────
  Input:  report_populated.docx + ABM_08_report_payload.json
  Output: ABM_report_final.docx (complete, print-ready)
```

Or, since your template is already fully populated for ABM (it's a filled example,
not a blank template), the immediate path is: **clean the template into a blank**,
add `{{ }}` placeholders in the sidebar, then run this two-pass flow.

---

### Step-by-step: preparing the blank template

1. Open `ABM_IoC_01_14_v1.docx` in Word
2. Replace all company-specific text in the sidebar and cover with `{{ }}` tokens:

| What to replace | Token |
|---|---|
| "ABM Industries Incorporated" (cover header) | `{{ company_name }}` |
| "(NYSE: ABM)" | `({{ exchange }}: {{ ticker }})` |
| "ABM Industries Incorporated, through its subsidiaries..." (corporate overview cell) | `{{ corporate_overview_short }}` |
| "One Liberty Plaza, New York, U.S." | `{{ address }}` |
| "www.abm.com" | `{{ website }}` |
| "+1.212.297.0200" | `{{ phone }}` |
| "Commercial Services" | `{{ sector }}` |
| "Miscellaneous Commercial Services" | `{{ industry }}` |
| "$44.3" (prev close) | `{{ prev_close }}` |
| "40.0 – 54.9" (52wk) | `{{ week52_range }}` |
| "914,464" (ADTV) | `{{ adtv }}` |
| "2,757" (mkt cap) | `{{ mkt_cap }}` |
| "4,212" (EV) | `{{ ev }}` |
| "60" (shares) | `{{ shares_out }}` |
| "98.5" (float) | `{{ float_pct }}` |
| "19.2x" (P/E) | `{{ pe }}` |
| "0.3x" (P/Sales) | `{{ p_sales }}` |
| "1.5x" (P/BV) | `{{ p_bv }}` |
| "9.7x" (EV/EBITDA) | `{{ ev_ebitda }}` |
| "0.5x" (EV/Sales) | `{{ ev_sales }}` |
| "Intro-act, LLC" (contact block) | `{{ firm_name }}` |
| "research@intro-act.com" | `{{ analyst_email }}` |
| "617-671-5148" | `{{ analyst_phone }}` |
| "01/14" (date label in valuation row) | `{{ valuation_date }}` |

3. Delete all the body content below the first section break (everything from
   "Integrated Facility Services Leader..." onwards)
4. In that empty area, type exactly: `[[REPORT_BODY]]` as a plain Normal paragraph
5. Save as `template_blank.docx`

---

### The word_render_v3.py you need to write (Phase 2)

This is the upgrade to your existing `word_render_v2.py`. The logic stays the same —
find the `[[REPORT_BODY]]` marker and insert paragraphs after it — but with three
important fixes based on what is actually in the template:

**Fix 1: Use the real style names from your template**

Your template has these styles (confirmed from inspection):
```python
STYLE_HEADING    = "Heading 1"       # section titles (red/dark, large)
STYLE_BODY       = "Body_Text"       # main prose paragraphs
STYLE_BULLET1    = "Bullet_1"        # first-level bullet
STYLE_BULLET2    = "Bullet_2"        # second-level bullet  
STYLE_CAPTION    = "Caption"         # figure captions
STYLE_SOURCE     = "Source"          # source attribution lines
STYLE_CHART_HDG  = "Chart_Heading"   # chart/figure headings
STYLE_STRAPLINE  = "Strapline Heading" # the bold thesis opener line
```

Your v2 was referencing `"Section Heading"`, `"Body Text"`, `"Bullet"` — none of
which exactly match the real style names in this template. That's why formatting
was likely wrong.

**Fix 2: The highlights list on the front page uses `Bullet_1` style**

Looking at the actual document, the 5 key highlights on the opening page are
`Bullet_1` paragraphs, not a separate section. The thesis text above them is
`Normal` style. Your renderer needs to match this exactly.

**Fix 3: The body section starts with the title in `Normal` style (bold)**

The `"Integrated Facility Services Leader..."` line is `Normal` style, not a heading.
The section headings like "Company Overview", "Technology-Enabled ESG..." are
`Heading 1`.

---

### docxtpl Pass 1 code (add to word_render_v3.py)

```python
from docxtpl import DocxTemplate

def render_cover_page(
    template_path: str,
    output_path: str,
    context: dict,
) -> None:
    """
    Pass 1: Fill all {{ }} placeholders in the cover page and sidebar.
    The body [[REPORT_BODY]] marker is left intact for Pass 2.
    """
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    tpl.save(output_path)
```

Context dict you pass:
```python
cover_context = {
    "company_name":        payload["company_name"],
    "ticker":              payload["ticker"],
    "exchange":            "NYSE",           # add to payload or hardcode per report
    "corporate_overview_short": "...",       # 2-sentence business description — add to payload
    "address":             "",               # add these fields to payload in Phase 2
    "website":             "",
    "phone":               "",
    "sector":              "",
    "industry":            "",
    "prev_close":          "",
    "week52_range":        "",
    "adtv":                "",
    "mkt_cap":             "",
    "ev":                  "",
    "shares_out":          "",
    "float_pct":           "",
    "pe":                  "",
    "p_sales":             "",
    "p_bv":                "",
    "ev_ebitda":           "",
    "ev_sales":            "",
    "firm_name":           "Intro-act, LLC",
    "analyst_email":       "",
    "analyst_phone":       "",
    "valuation_date":      "",
    "report_date":         "",
}
```

---

### Install docxtpl

```bash
pip install docxtpl
```

It depends on `python-docx` (which you already have) and `Jinja2`.

---

## 3. What to build in Phase 2 (templating sprint)

Priority order:

1. **Clean the template** — create `template_blank.docx` with all `{{ }}` tokens
   in the sidebar (30 min in Word)

2. **word_render_v3.py** — upgrade the renderer with:
   - Pass 1: `docxtpl.render()` for cover context
   - Pass 2: `python-docx` body insertion using correct style names
   - A `--cover_json` flag so market data (price, mkt cap, etc.) can be passed
     separately from the AI-generated payload

3. **Extend the payload** — add a `cover_data` dict to `ABM_08_report_payload.json`
   with the market stats fields. These come from a data source (FactSet, Bloomberg,
   or manual entry for now), not from the AI pipeline.

4. **Full pipeline command** — chain the two scripts:
   ```bash
   python autoanalyst_v3.py --ticker ABM --links_xlsx links.xlsx ...
   python word_render_v3.py \
       --payload output_ABM/ABM_08_report_payload.json \
       --template  template_blank.docx \
       --output    ABM_IoC_final.docx \
       --cover_json cover_data_ABM.json
   ```

---

## 4. What docxtpl cannot do (scope boundary)

- **Charts and figures**: The existing figures (quarterly revenue bar chart, profitability
  margin chart, etc.) are embedded images in the current template. In Phase 2 these will
  need to be generated as PNG/SVG by a charting script (matplotlib or openpyxl charts)
  and then inserted via `python-docx`'s `add_picture`. docxtpl has `InlineImage` support
  but it's limited to images you generate ahead of time.

- **The price performance chart** (Figure in sidebar): This is a live chart pulled from
  FactSet in the current template. In Phase 2 this will be a placeholder image or a
  matplotlib line chart generated from a price series you pass in.

- **Dynamic table population** (the financial estimates table that Goldman / JPM put
  at the back): This is a Phase 2 deliverable — a properly structured XLSX model that
  feeds a Word table via `python-docx`.

These are all solvable but they belong in Phase 2, not Phase 1.
