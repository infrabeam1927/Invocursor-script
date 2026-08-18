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

Progress is saved to PROGRESS_FILE after every company (reviewed or
errored), scoped to the exact pair of report files used. Re-running against
the same reports offers to resume where you left off instead of starting
the ~216-company review over from the top; running against newer reports
starts fresh automatically.

Every company handled (reviewed or errored) also gets a row in
OUTPUT_TRACKER_FILE — a single Excel file in the "Output tracker" folder
that keeps accumulating across every run (not one file per run), showing
which companies were handled, when, and with what outcome.

Usage:
    python main.py
"""

import glob
import json
import os
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright

from fill_contact_form import fill_contact_form_page

SECURITY_REPORT_GLOB = "security_report_*.xlsx"
# The [0-9]* requires a timestamp suffix, which excludes the older, differently
# shaped "contact_link_matches.xlsx" (one row per link, no Match N columns).
CONTACT_LINKS_GLOB = "contact_link_matches_[0-9]*.xlsx"

ACCEPTABLE_RISK_LEVELS = {"Clean"}
ACCEPTABLE_MATCH_CONFIDENCE = {"high confidence"}

PROGRESS_FILE = "outreach_progress.json"

# Single file that accumulates a row per company across every run — not
# timestamped, always the same path, updated (not replaced) each time.
OUTPUT_TRACKER_DIR = "Output tracker"
OUTPUT_TRACKER_FILE = os.path.join(OUTPUT_TRACKER_DIR, "Output tracker.xlsx")
OUTPUT_TRACKER_COLUMNS = [
    "Company", "Website", "Contact URL", "Outcome",
    "Security Risk", "Match Confidence", "Last Handled",
]


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
        qualified: list of (company, contact_url, website, risk, match_confidence) tuples
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
        qualified.append((company, str(match_link).strip(), website, risk, match_confidence))

    return qualified, excluded


def load_progress(security_report_path: str, contact_links_path: str) -> set:
    """Returns the set of company names already completed in a previous run
    against this exact pair of report files, or an empty set if there's no
    matching progress file."""
    if not os.path.exists(PROGRESS_FILE):
        return set()
    try:
        with open(PROGRESS_FILE, "r") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return set()

    if data.get("security_report") != security_report_path or data.get("contact_links_report") != contact_links_path:
        return set()  # progress is from a different pair of reports — doesn't apply here
    return set(data.get("completed_companies", []))


def save_progress(security_report_path: str, contact_links_path: str, completed_companies: set) -> None:
    with open(PROGRESS_FILE, "w") as f:
        json.dump({
            "security_report": security_report_path,
            "contact_links_report": contact_links_path,
            "completed_companies": sorted(completed_companies),
        }, f, indent=2)


def record_output(company: str, website: str, contact_url: str, outcome: str, risk: str, match_confidence: str) -> None:
    """Upserts one company's row into the single, ever-growing
    OUTPUT_TRACKER_FILE — creates the file/folder on first use, replaces the
    row if this company was already in it (e.g. re-handled after a reset),
    and rewrites the file so it's always current at the end of every run."""
    os.makedirs(OUTPUT_TRACKER_DIR, exist_ok=True)

    if os.path.exists(OUTPUT_TRACKER_FILE):
        try:
            df = pd.read_excel(OUTPUT_TRACKER_FILE)
        except Exception:
            df = pd.DataFrame(columns=OUTPUT_TRACKER_COLUMNS)
    else:
        df = pd.DataFrame(columns=OUTPUT_TRACKER_COLUMNS)

    df = df[df["Company"] != company] if "Company" in df.columns else df
    new_row = pd.DataFrame([{
        "Company": company,
        "Website": website,
        "Contact URL": contact_url,
        "Outcome": outcome,
        "Security Risk": risk,
        "Match Confidence": match_confidence,
        "Last Handled": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }])
    df = pd.concat([df, new_row], ignore_index=True)[OUTPUT_TRACKER_COLUMNS]
    df = df.sort_values("Last Handled", ascending=False, kind="stable")
    df.to_excel(OUTPUT_TRACKER_FILE, index=False)


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

    completed = load_progress(security_report_path, contact_links_path)
    remaining = [c for c in qualified if c[0] not in completed]

    if completed and remaining != qualified:
        already_done = len(qualified) - len(remaining)
        answer = input(
            f"Found previous progress against these reports: {already_done}/{len(qualified)} "
            f"companies already reviewed. Resume and skip those? (Y/n): "
        ).strip().lower()
        if answer in ("", "y", "yes"):
            qualified = remaining
        else:
            completed = set()  # starting over — previous progress will be overwritten as we go

    if not qualified:
        answer = input(
            "All qualifying companies have already been reviewed against these reports. "
            "Reset progress and start over? (y/N): "
        ).strip().lower()
        if answer not in ("y", "yes"):
            print("Nothing to do.")
            return
        completed = set()
        qualified, _ = load_qualified_companies(security_report_path, contact_links_path)

    processed, errored, stopped_early = 0, 0, False
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        try:
            for i, (company, contact_url, website, risk, match_confidence) in enumerate(qualified, start=1):
                print(f"\n=== [{i}/{len(qualified)}] {company} ({website}) ===")
                page = browser.new_page()
                outcome = "Reviewed"
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
                    outcome = "Errored"
                    print(f"  [ERROR] Skipping {company}: {exc}")
                finally:
                    page.close()
                completed.add(company)
                save_progress(security_report_path, contact_links_path, completed)
                record_output(company, website, contact_url, outcome, risk, match_confidence)
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
