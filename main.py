"""
Main entry point: links security_check.py, find_contact_links.py, and
fill_contact_form.py into one pipeline.

This does NOT run security_check.py or find_contact_links.py itself — run
those first (they can take a while over ~300 sites). This script reads the
most recently generated report from each and, for every company that is
BOTH:
    - "acceptable": Clean risk in the latest security_report_*.xlsx, and
    - "doable": has a high-confidence Match 1 link in the latest
      contact_link_matches_*.xlsx (Match 1 is already each company's best
      match, since find_contact_links.py sorts matches best-first)
opens that company's contact page and runs the same interactive,
never-auto-submit form fill as fill_contact_form.py — one browser session,
reused across companies, with a review pause between each so you can check
and submit (or skip) before moving on.

Companies that fail either bar (risky site, no contact page found, or only
a lower-confidence match) are skipped and listed in the summary — they're
not silently dropped.

Usage:
    python main.py
"""

import glob
import os

import pandas as pd
from playwright.sync_api import sync_playwright

from fill_contact_form import fill_contact_form_page

SECURITY_REPORT_GLOB = "security_report_*.xlsx"
# The [0-9]* requires a timestamp suffix, which excludes the older, differently
# shaped "contact_link_matches.xlsx" (one row per link, no Match N columns).
CONTACT_LINKS_GLOB = "contact_link_matches_[0-9]*.xlsx"

ACCEPTABLE_RISK_LEVELS = {"Clean"}
ACCEPTABLE_MATCH_CONFIDENCE = {"high confidence"}


def _latest_file(pattern: str) -> str:
    matches = glob.glob(pattern)
    if not matches:
        raise FileNotFoundError(
            f"No file matching '{pattern}' found in the current directory. "
            f"Run security_check.py / find_contact_links.py first."
        )
    return max(matches, key=os.path.getmtime)


def load_qualified_companies(security_report_path: str, contact_links_path: str):
    """Returns (qualified, excluded):
        qualified: list of (company, contact_url, website) tuples
        excluded: dict of {company: reason_str} for everything left out
    """
    security_df = pd.read_excel(security_report_path)
    links_df = pd.read_excel(contact_links_path)

    security_by_company = {
        str(row["Company"]).strip(): row["Risk Level"] for _, row in security_df.iterrows()
    }
    links_by_company = {
        str(row["Company"]).strip(): row for _, row in links_df.iterrows()
    }

    all_companies = sorted(set(security_by_company) | set(links_by_company))
    qualified = []
    excluded = {}

    for company in all_companies:
        risk = security_by_company.get(company)
        link_row = links_by_company.get(company)

        if risk is None:
            excluded[company] = "not present in security report"
            continue
        if risk not in ACCEPTABLE_RISK_LEVELS:
            excluded[company] = f"security risk level is {risk}"
            continue
        if link_row is None:
            excluded[company] = "not present in contact-link report"
            continue

        match_link = link_row.get("Match 1 Link")
        match_confidence = link_row.get("Match 1 Confidence")
        if pd.isna(match_link) or not str(match_link).strip():
            excluded[company] = "no contact-style link found on the site"
            continue
        if match_confidence not in ACCEPTABLE_MATCH_CONFIDENCE:
            excluded[company] = f"best contact link found was only '{match_confidence}'"
            continue

        website = link_row.get("Website", "")
        qualified.append((company, str(match_link).strip(), website))

    return qualified, excluded


def main():
    security_report_path = _latest_file(SECURITY_REPORT_GLOB)
    contact_links_path = _latest_file(CONTACT_LINKS_GLOB)
    print(f"Security report:     {security_report_path}")
    print(f"Contact-link report: {contact_links_path}\n")

    qualified, excluded = load_qualified_companies(security_report_path, contact_links_path)
    total = len(qualified) + len(excluded)

    print(f"{len(qualified)}/{total} companies qualify for outreach (Clean risk + high-confidence contact link).")
    if excluded:
        print(f"{len(excluded)} excluded:")
        for company, reason in sorted(excluded.items()):
            print(f"  - {company}: {reason}")
    print()

    if not qualified:
        print("No qualifying companies — nothing to do.")
        return

    processed, errored, stopped_early = 0, 0, False
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        try:
            for i, (company, contact_url, website) in enumerate(qualified, start=1):
                print(f"\n=== [{i}/{len(qualified)}] {company} ({website}) ===")
                page = browser.new_page()
                try:
                    fill_contact_form_page(page, company, contact_url)
                    print("Form filled. Review it in the browser window.")
                    print("This will NOT auto-submit — submit it yourself once you're happy with it.")
                    response = input(
                        "Press Enter to move to the next company, or type 'q' to stop the run... "
                    ).strip().lower()
                    processed += 1
                    if response in ("q", "quit"):
                        stopped_early = True
                except Exception as exc:
                    errored += 1
                    print(f"  [ERROR] Skipping {company}: {exc}")
                finally:
                    page.close()
                if stopped_early:
                    break
        finally:
            browser.close()

    print(
        f"\nDone. Reviewed {processed}/{len(qualified)} qualifying companies "
        f"({errored} errored){' — stopped early by request' if stopped_early else ''}."
    )


if __name__ == "__main__":
    main()
