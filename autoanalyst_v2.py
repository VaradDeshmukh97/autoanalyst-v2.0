"""
autoanalyst_v2.py

Orchestration script for the v2 AutoAnalyst initiation of coverage pipeline.

Current phase:
- ingest an Excel workbook containing secondary research links in a single column
- build source-level summaries
- consolidate a master memo
- generate investment thesis + 5 key highlights
- ask the user to confirm / edit the 5 key highlights
- write the report section by section
- produce a final cleaned report plus structured artifacts

Usage examples:
python autoanalyst_v2.py \
    --company_name "NVIDIA Corporation" \
    --short_name "NVIDIA" \
    --ticker "NVDA" \
    --links_xlsx "links.xlsx"

CMD example:
python autoanalyst_v2.py --company_name "NVIDIA Corporation" --short_name "NVIDIA" --ticker "NVDA" --links_xlsx "links.xlsx"
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List

from dotenv import load_dotenv

from autoanalyst_utils_v2 import (
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


def prompt_user_to_edit_highlights(highlights: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """
    Interactive terminal confirmation/edit step for the five highlights.
    """
    print("\n" + "=" * 80)
    print("APPROVED 5 KEY HIGHLIGHTS")
    print("=" * 80)
    print(format_highlights_for_prompt(highlights))
    print("=" * 80)
    print("Enter to keep a highlight as-is. Type a replacement to edit it.")
    print("You can revise the title and the highlight sentence independently.")
    print("=" * 80 + "\n")

    edited: List[Dict[str, str]] = []
    for idx, h in enumerate(highlights, 1):
        print(f"Highlight {idx}")
        print(f"Current title     : {h.get('title', '')}")
        print(f"Current highlight : {h.get('highlight', '')}")
        print(f"Current implication: {h.get('implication', '')}")
        print(f"Current angle     : {h.get('section_angle', '')}")

        new_title = input("New title (ENTER to keep): ").strip()
        new_highlight = input("New highlight (ENTER to keep): ").strip()
        new_implication = input("New implication (ENTER to keep): ").strip()
        new_angle = input("New section angle (ENTER to keep): ").strip()

        edited.append(
            {
                "title": new_title or h.get("title", ""),
                "highlight": new_highlight or h.get("highlight", ""),
                "implication": new_implication or h.get("implication", ""),
                "section_angle": new_angle or h.get("section_angle", ""),
            }
        )
        print("")

    return edited


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="AutoAnalyst v2 — initiation of coverage report pipeline"
    )

    parser.add_argument("--company_name", required=True, help="Full legal company name")
    parser.add_argument(
        "--short_name",
        default=None,
        help="Short reference name used in report prose. Defaults to company_name.",
    )
    parser.add_argument(
        "--ticker",
        required=True,
        help="Ticker or internal identifier used for filenames and outputs",
    )
    parser.add_argument(
        "--links_xlsx",
        required=True,
        help="Excel workbook containing secondary research links in a single column",
    )
    parser.add_argument(
        "--sheet_name",
        default=None,
        help="Worksheet name containing the links. Defaults to the first sheet.",
    )
    parser.add_argument(
        "--url_column",
        default=None,
        help="Optional Excel column letter (e.g. A) or header name containing URLs.",
    )
    parser.add_argument(
        "--output_dir",
        default="autoanalyst_output",
        help="Directory where outputs will be written",
    )
    parser.add_argument(
        "--source_model",
        default=None,
        help="Model used for source summarization and memo generation. Defaults to AUTOANALYST_MODEL or gpt-4o-mini.",
    )
    parser.add_argument(
        "--report_model",
        default=None,
        help="Model used for thesis/highlight/report generation. Defaults to AUTOANALYST_MODEL or gpt-4o-mini.",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=None,
        help="HTTP timeout in seconds for source fetching",
    )
    parser.add_argument(
        "--max_sources",
        type=int,
        default=None,
        help="Optional cap on the number of source links processed",
    )
    parser.add_argument(
        "--auto_accept_highlights",
        action="store_true",
        help="Skip interactive highlight confirmation and accept the model output as-is",
    )

    return parser


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
    timeout = args.timeout

    client = create_openai_client()

    print("\n" + "=" * 90)
    print("AUTOANALYST v2 | INITIATION OF COVERAGE PIPELINE")
    print("=" * 90)
    print(f"Company name : {company_name}")
    print(f"Short name   : {short_name}")
    print(f"Ticker       : {ticker}")
    print(f"Links Excel  : {args.links_xlsx}")
    print(f"Output dir   : {output_dir}")
    print("=" * 90 + "\n")

    # -------------------------------------------------------------------------
    # STEP 1 — INGEST LINKS
    # -------------------------------------------------------------------------
    print("[STEP 1] Reading source links from Excel...")
    urls = read_urls_from_excel(
        excel_path=args.links_xlsx,
        sheet_name=args.sheet_name,
        url_column=args.url_column,
    )
    if not urls:
        raise RuntimeError("No URLs were found in the Excel workbook.")

    urls = [u for u in urls if u.lower().startswith(("http://", "https://")) or ".pdf" in u.lower()]
    print(f"[INFO] Found {len(urls)} usable URLs.\n")

    save_json(
        output_dir / f"{ticker}_ingested_urls.json",
        {"company_name": company_name, "short_name": short_name, "ticker": ticker, "urls": urls},
    )

    # -------------------------------------------------------------------------
    # STEP 2 — BUILD KNOWLEDGE BASE
    # -------------------------------------------------------------------------
    print("[STEP 2] Building source-level knowledge base...")
    knowledge_base = build_knowledge_base(
        client=client,
        urls=urls,
        company_short_name=short_name,
        model=source_model,
        timeout=timeout or 20,
        max_sources=args.max_sources,
    )
    save_text(output_dir / f"{ticker}_knowledge_base.txt", knowledge_base)
    print(f"[INFO] Knowledge base saved: {output_dir / f'{ticker}_knowledge_base.txt'}\n")

    # -------------------------------------------------------------------------
    # STEP 3 — CONSOLIDATE NOTES
    # -------------------------------------------------------------------------
    print("[STEP 3] Consolidating notes into a master memo...")
    memo = consolidate_notes(
        client=client,
        company_short_name=short_name,
        notes=knowledge_base,
        model=report_model,
    )
    save_text(output_dir / f"{ticker}_memo.txt", memo)
    print(f"[INFO] Memo saved: {output_dir / f'{ticker}_memo.txt'}\n")

    # -------------------------------------------------------------------------
    # STEP 4 — THESIS + 5 HIGHLIGHTS
    # -------------------------------------------------------------------------
    print("[STEP 4] Generating title and investment thesis...")
    thesis_payload = generate_title_and_thesis(
        client=client,
        company_short_name=short_name,
        company_name=company_name,
        memo=memo,
        model=report_model,
    )
    save_json(output_dir / f"{ticker}_thesis.json", thesis_payload)
    print(f"[INFO] Thesis saved: {output_dir / f'{ticker}_thesis.json'}\n")

    print("[STEP 4B] Generating 5 key highlights...")
    highlights_payload = generate_key_highlights(
        client=client,
        company_short_name=short_name,
        memo=memo,
        model=report_model,
    )
    highlights = highlights_payload["highlights"]
    save_json(output_dir / f"{ticker}_highlights_preedit.json", highlights_payload)

    if not args.auto_accept_highlights:
        highlights = prompt_user_to_edit_highlights(highlights)

    save_json(
        output_dir / f"{ticker}_highlights.json",
        {"highlights": highlights},
    )
    print(f"[INFO] Approved highlights saved: {output_dir / f'{ticker}_highlights.json'}\n")

    # -------------------------------------------------------------------------
    # STEP 5 — SECTION BY SECTION WRITING
    # -------------------------------------------------------------------------
    print("[STEP 5] Writing report sections...\n")

    title = thesis_payload.get("title") or f"{short_name} Initiation of Coverage"
    thesis_paragraphs = thesis_payload.get("investment_thesis", [])
    if isinstance(thesis_paragraphs, str):
        thesis_paragraphs = [thesis_paragraphs]
    thesis_paragraphs = [str(x).strip() for x in thesis_paragraphs if str(x).strip()]

    company_overview = generate_company_overview(
        client=client,
        company_short_name=short_name,
        company_name=company_name,
        memo=memo,
        model=report_model,
    )
    save_text(output_dir / f"{ticker}_company_overview.txt", company_overview)
    print("[INFO] Company Overview complete.")

    highlight_sections = []
    for idx, highlight in enumerate(highlights, 1):
        print(f"[INFO] Writing deep-dive section {idx}/5: {highlight.get('title', '')}")
        body = generate_highlight_detail(
            client=client,
            company_short_name=short_name,
            highlight=highlight,
            memo=memo,
            model=report_model,
        )
        highlight_sections.append((highlight, body))
        save_text(output_dir / f"{ticker}_highlight_{idx}.txt", body)

    financials = generate_financials_section(
        client=client,
        company_short_name=short_name,
        memo=memo,
        model=report_model,
    )
    save_text(output_dir / f"{ticker}_financials.txt", financials)
    print("[INFO] Financials complete.")

    risks = generate_risks_section(
        client=client,
        company_short_name=short_name,
        memo=memo,
        model=report_model,
    )
    save_text(output_dir / f"{ticker}_risks.txt", risks)
    print("[INFO] Risks complete.\n")

    draft_report = assemble_report_markdown(
        title=title,
        investment_thesis=thesis_paragraphs,
        highlights=highlights,
        company_overview=company_overview,
        highlight_sections=highlight_sections,
        financials=financials,
        risks=risks,
    )
    save_text(output_dir / f"{ticker}_report_draft.md", draft_report)
    print(f"[INFO] Draft report saved: {output_dir / f'{ticker}_report_draft.md'}")

    # -------------------------------------------------------------------------
    # STEP 6 — FINAL EDIT
    # -------------------------------------------------------------------------
    print("\n[STEP 6] Final edit pass...")
    final_report = final_edit_report(
        client=client,
        company_short_name=short_name,
        report_draft=draft_report,
        model=report_model,
    )
    save_text(output_dir / f"{ticker}_report_final.md", final_report)
    print(f"[INFO] Final report saved: {output_dir / f'{ticker}_report_final.md'}")

    # -------------------------------------------------------------------------
    # STEP 7 — STRUCTURED PAYLOAD FOR LATER WORD TEMPLATING
    # -------------------------------------------------------------------------
    payload = {
        "company_name": company_name,
        "short_name": short_name,
        "ticker": ticker,
        "title": title,
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
    save_json(output_dir / f"{ticker}_report_payload.json", payload)
    print(f"[INFO] Structured payload saved: {output_dir / f'{ticker}_report_payload.json'}")

    print("\n" + "=" * 90)
    print("PIPELINE COMPLETE")
    print("=" * 90)


if __name__ == "__main__":
    main()
