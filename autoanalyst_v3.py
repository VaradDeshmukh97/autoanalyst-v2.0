"""
autoanalyst_v3.py

Orchestration script for the AutoAnalyst initiation of coverage pipeline — v3.

Phase 1 (current):
  - Ingest an Excel workbook containing secondary research links in a single column.
  - Build source-level analytical notes for each URL.
  - Synthesize a master analytical memo from all source notes.
  - Generate a report title and two-paragraph investment thesis.
  - Generate 5 key investment highlights with full metadata.
  - Prompt the analyst to confirm / edit the highlights interactively.
  - Write the report section-by-section: Company Overview, 5 deep-dives,
    Financials, Risks.
  - Assemble a clean Markdown draft.
  - Run a final editorial pass for tone, redundancy, and precision.
  - Emit a structured JSON payload for Phase 2 (Word templating).

Phase 2 (next): Word template population, financial charts, formatted output.
Phase 3 (next): Team-facing web frontend.

Usage:
    python autoanalyst_v3.py \\
        --company_name "NVIDIA Corporation" \\
        --short_name "NVIDIA" \\
        --ticker "NVDA" \\
        --links_xlsx "links.xlsx"

    python autoanalyst_v3.py \\
        --company_name "NVIDIA Corporation" \\
        --short_name "NVIDIA" \\
        --ticker "NVDA" \\
        --links_xlsx "links.xlsx" \\
        --auto_accept_highlights \\
        --report_model gpt-4o \\
        --source_model gpt-4o-mini
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List

from dotenv import load_dotenv

from autoanalyst_utils_v3 import (
    assemble_report_markdown,
    build_knowledge_base,
    consolidate_notes,
    DEFAULT_MODEL,
    create_openai_client,
    final_edit_report,
    format_highlights_for_prompt,
    generate_company_overview,
    generate_financials_section,
    generate_highlight_detail,
    generate_key_highlights,
    generate_risks_section,
    generate_title_and_thesis,
    read_urls_from_excel,
    save_json,
    save_text,
    slugify,
)


# ---------------------------------------------------------------------------
# ANALYST HIGHLIGHT EDIT LOOP
# ---------------------------------------------------------------------------

def prompt_user_to_edit_highlights(highlights: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """
    Interactive terminal confirmation / edit step for the five highlights.

    The analyst sees the full set at once, then can edit each field.
    Pressing Enter on any field keeps the model's output.
    This is the single gating point before report section writing begins.
    """
    print("\n" + "=" * 90)
    print("ANALYST REVIEW — 5 KEY INVESTMENT HIGHLIGHTS")
    print("=" * 90)
    print("Read all five highlights before editing. Check for:")
    print("  • Overlap or redundancy between highlights")
    print("  • Any highlight that is too generic or not specific enough")
    print("  • Ordering — is highlight #1 really the most important?")
    print("  • Whether the five together cover the full investment case")
    print("=" * 90)
    print()
    print(format_highlights_for_prompt(highlights))
    print()
    print("=" * 90)
    print("EDITING — Press Enter to keep any field as-is.")
    print("=" * 90 + "\n")

    edited: List[Dict[str, str]] = []
    for idx, h in enumerate(highlights, 1):
        print(f"─── Highlight {idx} ───────────────────────────────────────────────")
        print(f"  Title       : {h.get('title', '')}")
        print(f"  Highlight   : {h.get('highlight', '')}")
        print(f"  Implication : {h.get('implication', '')}")
        print(f"  Angle       : {h.get('section_angle', '')}")
        print(f"  Evidence    : {h.get('conviction_basis', '')}")
        print()

        new_title       = input("  New title (ENTER to keep)       : ").strip()
        new_highlight   = input("  New highlight (ENTER to keep)   : ").strip()
        new_implication = input("  New implication (ENTER to keep) : ").strip()
        new_angle       = input("  New angle (ENTER to keep)       : ").strip()
        new_conviction  = input("  New evidence (ENTER to keep)    : ").strip()
        print()

        edited.append({
            "title":           new_title or h.get("title", ""),
            "highlight":       new_highlight or h.get("highlight", ""),
            "implication":     new_implication or h.get("implication", ""),
            "section_angle":   new_angle or h.get("section_angle", ""),
            "conviction_basis": new_conviction or h.get("conviction_basis", ""),
        })

    return edited


# ---------------------------------------------------------------------------
# CLI PARSER
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="AutoAnalyst v3 — Initiation of Coverage Report Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python autoanalyst_v3.py \\
      --company_name "ABM Industries Incorporated" \\
      --short_name "ABM" \\
      --ticker "ABM" \\
      --links_xlsx links.xlsx

  python autoanalyst_v3.py \\
      --company_name "ABM Industries Incorporated" \\
      --short_name "ABM" \\
      --ticker "ABM" \\
      --links_xlsx links.xlsx \\
      --auto_accept_highlights \\
      --source_model gpt-4o-mini \\
      --report_model gpt-4o \\
      --max_sources 20
        """,
    )

    # Identity
    parser.add_argument("--company_name", required=True,
                        help="Full legal company name (used in report prose and filenames)")
    parser.add_argument("--short_name", default=None,
                        help="Short reference name for report prose. Defaults to company_name.")
    parser.add_argument("--ticker", required=True,
                        help="Exchange ticker or internal identifier (used in all output filenames)")

    # Input
    parser.add_argument("--links_xlsx", required=True,
                        help="Excel workbook containing secondary research links in a single column")
    parser.add_argument("--sheet_name", default=None,
                        help="Worksheet name containing the links. Defaults to the first sheet.")
    parser.add_argument("--url_column", default=None,
                        help="Column letter (e.g. A) or header name containing URLs. "
                             "If omitted, auto-detected by URL density.")

    # Output
    parser.add_argument("--output_dir", default="autoanalyst_output",
                        help="Directory where all intermediate and final outputs are written")

    # Models
    parser.add_argument("--source_model", default=None,
                        help="Model for source summarization. "
                             "Defaults to AUTOANALYST_MODEL env var or gpt-4o-mini.")
    parser.add_argument("--report_model", default=None,
                        help="Model for memo, thesis, highlights, and section writing. "
                             "Defaults to AUTOANALYST_MODEL env var or gpt-4o-mini. "
                             "Recommend gpt-4o for highest quality.")

    # Behavior
    parser.add_argument("--timeout", type=int, default=None,
                        help="HTTP timeout in seconds for source fetching (default: 20)")
    parser.add_argument("--max_sources", type=int, default=None,
                        help="Cap the number of source links processed (useful for testing)")
    parser.add_argument("--auto_accept_highlights", action="store_true",
                        help="Skip interactive highlight confirmation and use model output as-is. "
                             "Not recommended for final reports.")

    return parser


# ---------------------------------------------------------------------------
# PIPELINE
# ---------------------------------------------------------------------------

def main() -> None:
    load_dotenv()

    parser = build_parser()
    args = parser.parse_args()

    company_name = args.company_name.strip()
    short_name = (args.short_name or company_name).strip()
    ticker = args.ticker.strip().upper()
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    source_model = args.source_model or DEFAULT_MODEL
    report_model = args.report_model or DEFAULT_MODEL
    timeout = args.timeout or 20

    client = create_openai_client()

    _banner("AUTOANALYST v3 | INITIATION OF COVERAGE PIPELINE")
    print(f"  Company      : {company_name}")
    print(f"  Short name   : {short_name}")
    print(f"  Ticker       : {ticker}")
    print(f"  Links file   : {args.links_xlsx}")
    print(f"  Output dir   : {output_dir}")
    print(f"  Source model : {source_model}")
    print(f"  Report model : {report_model}")
    _divider()

    # -------------------------------------------------------------------------
    # STEP 1 — INGEST LINKS
    # -------------------------------------------------------------------------
    _step(1, "Reading source links from Excel")
    urls = read_urls_from_excel(
        excel_path=args.links_xlsx,
        sheet_name=args.sheet_name,
        url_column=args.url_column,
    )
    if not urls:
        raise RuntimeError("No URLs were found in the Excel workbook.")

    urls = [u for u in urls if u.lower().startswith(("http://", "https://")) or ".pdf" in u.lower()]
    print(f"[INFO] {len(urls)} usable URLs found.\n")

    save_json(
        output_dir / f"{ticker}_01_ingested_urls.json",
        {"company_name": company_name, "short_name": short_name, "ticker": ticker, "urls": urls},
    )

    # -------------------------------------------------------------------------
    # STEP 2 — BUILD KNOWLEDGE BASE
    # -------------------------------------------------------------------------
    _step(2, "Building source-level knowledge base (analytical notes per URL)")
    knowledge_base = build_knowledge_base(
        client=client,
        urls=urls,
        company_short_name=short_name,
        model=source_model,
        timeout=timeout,
        max_sources=args.max_sources,
    )
    kb_path = output_dir / f"{ticker}_02_knowledge_base.txt"
    save_text(kb_path, knowledge_base)
    print(f"[INFO] Knowledge base saved: {kb_path}\n")

    # -------------------------------------------------------------------------
    # STEP 3 — CONSOLIDATE INTO MASTER MEMO
    # -------------------------------------------------------------------------
    _step(3, "Synthesizing master analytical memo")
    memo = consolidate_notes(
        client=client,
        company_short_name=short_name,
        notes=knowledge_base,
        model=report_model,
    )
    memo_path = output_dir / f"{ticker}_03_master_memo.txt"
    save_text(memo_path, memo)
    print(f"[INFO] Master memo saved: {memo_path}\n")

    # -------------------------------------------------------------------------
    # STEP 4A — TITLE + INVESTMENT THESIS
    # -------------------------------------------------------------------------
    _step(4, "Generating report title and investment thesis")
    thesis_payload = generate_title_and_thesis(
        client=client,
        company_short_name=short_name,
        company_name=company_name,
        memo=memo,
        model=report_model,
    )
    save_json(output_dir / f"{ticker}_04a_thesis.json", thesis_payload)

    print(f"\n  Title             : {thesis_payload.get('title', '')}")
    print(f"  Thesis logic      : {thesis_payload.get('thesis_logic', '')}")
    print(f"  Variant perception: {thesis_payload.get('variant_perception', '')}\n")

    # -------------------------------------------------------------------------
    # STEP 4B — 5 KEY HIGHLIGHTS
    # -------------------------------------------------------------------------
    _step("4B", "Generating 5 key investment highlights")
    highlights_payload = generate_key_highlights(
        client=client,
        company_short_name=short_name,
        memo=memo,
        model=report_model,
    )
    highlights = highlights_payload["highlights"]
    save_json(output_dir / f"{ticker}_04b_highlights_preedit.json", highlights_payload)

    # ── ANALYST GATE ──────────────────────────────────────────────────────────
    if not args.auto_accept_highlights:
        highlights = prompt_user_to_edit_highlights(highlights)
    else:
        print("[INFO] --auto_accept_highlights set: skipping interactive review.\n")

    save_json(
        output_dir / f"{ticker}_04b_highlights_approved.json",
        {"highlights": highlights},
    )
    print(f"[INFO] Approved highlights saved: {output_dir / f'{ticker}_04b_highlights_approved.json'}\n")

    # -------------------------------------------------------------------------
    # STEP 5 — SECTION-BY-SECTION WRITING
    # -------------------------------------------------------------------------
    _step(5, "Writing report sections")

    title = thesis_payload.get("title") or f"{short_name} — Initiation of Coverage"
    thesis_paragraphs = thesis_payload.get("investment_thesis", [])
    if isinstance(thesis_paragraphs, str):
        thesis_paragraphs = [thesis_paragraphs]
    thesis_paragraphs = [str(x).strip() for x in thesis_paragraphs if str(x).strip()]

    # Company Overview
    print("[INFO] Writing Company Overview...")
    company_overview = generate_company_overview(
        client=client,
        company_short_name=short_name,
        company_name=company_name,
        memo=memo,
        model=report_model,
    )
    save_text(output_dir / f"{ticker}_05a_company_overview.txt", company_overview)
    print("[INFO] Company Overview complete.\n")

    # Five Highlight Deep-Dives
    highlight_sections = []
    for idx, highlight in enumerate(highlights, 1):
        print(f"[INFO] Writing deep-dive {idx}/5: {highlight.get('title', '')}")
        body = generate_highlight_detail(
            client=client,
            company_short_name=short_name,
            highlight=highlight,
            memo=memo,
            model=report_model,
        )
        highlight_sections.append((highlight, body))
        save_text(output_dir / f"{ticker}_05b_highlight_{idx}.txt", body)

    print()

    # Financials
    print("[INFO] Writing Financials section...")
    financials = generate_financials_section(
        client=client,
        company_short_name=short_name,
        memo=memo,
        model=report_model,
    )
    save_text(output_dir / f"{ticker}_05c_financials.txt", financials)
    print("[INFO] Financials complete.\n")

    # Risks
    print("[INFO] Writing Risks section...")
    risks = generate_risks_section(
        client=client,
        company_short_name=short_name,
        memo=memo,
        model=report_model,
    )
    save_text(output_dir / f"{ticker}_05d_risks.txt", risks)
    print("[INFO] Risks complete.\n")

    # Draft Assembly
    draft_report = assemble_report_markdown(
        title=title,
        investment_thesis=thesis_paragraphs,
        highlights=highlights,
        company_overview=company_overview,
        highlight_sections=highlight_sections,
        financials=financials,
        risks=risks,
    )
    draft_path = output_dir / f"{ticker}_06_report_draft.md"
    save_text(draft_path, draft_report)
    print(f"[INFO] Draft report saved: {draft_path}")

    # -------------------------------------------------------------------------
    # STEP 6 — FINAL EDITORIAL PASS
    # -------------------------------------------------------------------------
    _step(6, "Final editorial pass (tone, redundancy, precision)")
    final_report = final_edit_report(
        client=client,
        company_short_name=short_name,
        report_draft=draft_report,
        model=report_model,
    )
    final_path = output_dir / f"{ticker}_07_report_final.md"
    save_text(final_path, final_report)
    print(f"[INFO] Final report saved: {final_path}")

    # -------------------------------------------------------------------------
    # STEP 7 — STRUCTURED JSON PAYLOAD (for Phase 2 Word templating)
    # -------------------------------------------------------------------------
    _step(7, "Saving structured JSON payload for Word templating")
    payload: Dict[str, Any] = {
        "company_name": company_name,
        "short_name": short_name,
        "ticker": ticker,
        "title": title,
        "thesis_logic": thesis_payload.get("thesis_logic", ""),
        "variant_perception": thesis_payload.get("variant_perception", ""),
        "investment_thesis": thesis_paragraphs,
        "highlights": highlights,
        "company_overview": company_overview,
        "highlight_sections": [
            {
                "highlight": h,
                "section_body": body,
            }
            for h, body in highlight_sections
        ],
        "financials": financials,
        "risks": risks,
        "urls": urls,
    }
    payload_path = output_dir / f"{ticker}_08_report_payload.json"
    save_json(payload_path, payload)
    print(f"[INFO] Structured payload saved: {payload_path}")

    # -------------------------------------------------------------------------
    # DONE
    # -------------------------------------------------------------------------
    _banner("PIPELINE COMPLETE")
    print(f"  Output directory : {output_dir.resolve()}")
    print(f"  Final report     : {final_path.name}")
    print(f"  Word payload     : {payload_path.name}")
    _divider()


# ---------------------------------------------------------------------------
# DISPLAY HELPERS
# ---------------------------------------------------------------------------

def _banner(text: str) -> None:
    width = 90
    print("\n" + "=" * width)
    print(f"  {text}")
    print("=" * width)


def _divider() -> None:
    print("─" * 90 + "\n")


def _step(number: int | str, description: str) -> None:
    print(f"\n[STEP {number}] {description}")
    print("─" * 70)


# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    main()
