#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Session 37/38 — bordereau_extractor.py (redesigned)
Detects the "bordereau de prix" (price schedule table) page(s) in an AO PDF
and renders them as a single composite PNG (stacked vertically if multiple
pages match) for embedding into the Tableau decisionnel.

Detection signal (redesigned from the Session 23 keyword-matching approach):
anchors on the document's OWN regulatory label for the price table
("Bordereau de Prix" / "Offre de Prix" / "Bordereau de Soumission") rather
than the table's internal column vocabulary, which was proven across 5+
real municipal AOs (APM, Ville de Montreal, MRC de Lotbiniere, Ville
d'Alma, Ville de Shawinigan) not to generalize - each municipality phrases
its price table's own column headers differently, but every AO reviewed
labels the artifact itself using standardized legal/regulatory terminology.

A heading match alone is not sufficient (definitions clauses, tables of
contents, and cross-references routinely mention "Bordereau de Prix" with
zero numeric content), so every heading match requires currency evidence
(>=2 "$" signs) on the same page. For the one remaining ambiguous case -
a currency-less heading page adjacent to an already-validated page (e.g.
a price table whose item-list page has no $ signs, split across two
pages) - a narrow, targeted Claude call resolves it. Three lexical
heuristics (page-position, line-isolation, unit-of-measure token
frequency) were each tried for this specific join decision and each
failed against real adversarial data before this fallback was adopted.

Verified against 9 real files spanning 5 municipalities, including the
original Lotbiniere bug report and Session 23's own false-positive
reference file (2026-55-100.pdf, correctly isolates page 93 only,
rejecting 31/45/92), before deployment.
"""
import os
import re
from pdfminer.high_level import extract_text
from pdfminer.pdfpage import PDFPage

HEADING_RE = re.compile(
    r'bordereau\s+(des?\s+)?(prix|soumission)|offre\s+de\s+prix',
    re.IGNORECASE
)
MIN_DOLLAR_SIGNS = 2

CONTINUATION_PROMPT = """You are reviewing two adjacent pages from a Quebec municipal procurement (appel d'offres) price bordereau document.

PAGE A (heading suggests it may be part of a "Bordereau de Prix" / price schedule, but contains no dollar amounts):
---
{page_a}
---

PAGE B (contains dollar amounts):
---
{page_b}
---

Question: Are these two pages part of the SAME continuous price table — e.g., Page A lists items/quantities and Page B lists the corresponding prices for those same items — or is Page A unrelated content (a different section, a declaration, a definitions clause) that merely mentions a related term in passing?

Answer with exactly one word: SAME or UNRELATED."""


def _check_continuation(page_a_text, page_b_text):
    """Narrow, targeted Claude call for the one ambiguous case lexical
    rules can't resolve. Only invoked rarely (see find_bordereau_pages).
    Fails closed (returns False) on any API error - a missed continuation
    page is a minor completeness loss, not a correctness risk."""
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=10,
            messages=[{
                "role": "user",
                "content": CONTINUATION_PROMPT.format(
                    page_a=page_a_text, page_b=page_b_text
                )
            }]
        )
        return response.content[0].text.strip().upper() == "SAME"
    except Exception as e:
        print(f"WARNING: continuation check failed, skipping: {e}")
        return False


def find_bordereau_pages(pdf_path):
    """Return a list of 1-indexed page numbers matching the price-table
    signal, anchored on the document's own label for the bordereau rather
    than its internal column vocabulary."""
    with open(pdf_path, "rb") as f:
        n_pages = sum(1 for _ in PDFPage.get_pages(f))

    page_text = {}
    has_heading = {}
    has_currency = {}
    for i in range(1, n_pages + 1):
        text = extract_text(pdf_path, page_numbers=[i - 1]) or ""
        page_text[i] = text
        has_heading[i] = bool(HEADING_RE.search(text))
        has_currency[i] = text.count("$") >= MIN_DOLLAR_SIGNS

    # Tier 1: deterministic core (proven, zero known false positives)
    matches = {i for i in range(1, n_pages + 1)
               if has_heading[i] and has_currency[i]}

    # Tier 2: narrow semantic fallback for currency-less heading pages
    # immediately adjacent to an already-validated page
    for i in range(1, n_pages + 1):
        if has_heading[i] and not has_currency[i] and i not in matches:
            for neighbor in (i - 1, i + 1):
                if neighbor in matches:
                    if _check_continuation(page_text[i], page_text[neighbor]):
                        matches.add(i)
                    break

    return sorted(matches)


def render_bordereau_image(pdf_path, output_png_path, dpi=150, max_width=1000):
    """
    Detect and render the bordereau page(s) as a single composite PNG.
    Returns (output_png_path, [page_numbers]) on success, or (None, []) if
    no bordereau pages were found.
    """
    pages = find_bordereau_pages(pdf_path)
    if not pages:
        return None, []

    from pdf2image import convert_from_path
    from PIL import Image

    rendered = []
    for p in pages:
        imgs = convert_from_path(pdf_path, dpi=dpi, first_page=p, last_page=p)
        if imgs:
            rendered.append(imgs[0])

    if not rendered:
        return None, pages

    # Resize each page to a consistent max_width, preserving aspect ratio
    resized = []
    for im in rendered:
        if im.width > max_width:
            ratio = max_width / im.width
            im = im.resize((max_width, int(im.height * ratio)))
        resized.append(im)

    gap = 15
    total_width = max(im.width for im in resized)
    total_height = sum(im.height for im in resized) + gap * (len(resized) - 1)

    composite = Image.new("RGB", (total_width, total_height), "white")
    y = 0
    for im in resized:
        composite.paste(im, (0, y))
        y += im.height + gap

    composite.save(output_png_path)
    return output_png_path, pages


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python3 bordereau_extractor.py <pdf_path> <output_png_path>")
        sys.exit(1)
    path, out = sys.argv[1], sys.argv[2]
    result, pages = render_bordereau_image(path, out)
    if result:
        print(f"Bordereau found on page(s) {pages} — saved to {result}")
    else:
        print("No bordereau de prix page detected.")
