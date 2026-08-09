"""
Scans prospect websites for links to a "contact us"-style page (contact,
get-in-touch, support, book-a-demo, etc.) and writes matches to an Excel file.

Input:
    An Excel workbook (default: "Outreach tracker.xlsx") with a company name
    column and a website column. Column F is expected to hold the website
    URL; the script also looks for a header containing "website"/"url" and
    a header containing "company"/"name" and prefers those if present.

Output:
    An Excel workbook named "contact_link_matches_<YYYYMMDD_HHMMSS>.xlsx"
    with one row per company: Company, Website, Status, then a Match N Link /
    Match N Keyword / Match N Confidence triplet for each matched link,
    ordered high confidence first. Rows are sorted by each company's best
    confidence (high, then medium, then no match/error).

Can also be used as a library for a single URL or a list of URLs:

    from find_contact_links import find_contact_links
    matches = find_contact_links("https://example.com")

Usage:
    python find_contact_links.py
"""

import re
import time
from datetime import datetime
from urllib.parse import urljoin, urlparse, unquote

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
INPUT_FILE = "Outreach tracker.xlsx"
OUTPUT_FILE_PREFIX = "contact_link_matches"
FALLBACK_WEBSITE_COLUMN_INDEX = 5  # column F, used if no "website"/"url" header is found

REQUEST_TIMEOUT_SECONDS = 10
DELAY_BETWEEN_REQUESTS_SECONDS = 1.5

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

# Keywords matched against the link's URL path and visible text. Only terms
# that point at an actual contact-style page are considered — no generic
# fallback pages like "about". High-confidence keywords name a contact/
# outreach page specifically; medium-confidence keywords are broader terms
# that often lead to one.
HIGH_CONFIDENCE_KEYWORDS = [
    "contact", "contact us", "contact-us", "contactus",
    "get in touch", "get-in-touch", "getintouch",
    "reach us", "reach-us", "reach out", "reach-out", "reachout",
    "book a call", "book-a-call",
    "book a demo", "book-a-demo",
    "request a demo", "request-a-demo",
    "talk to us", "talk-to-us",
    "say hello", "say-hello",
]
MEDIUM_CONFIDENCE_KEYWORDS = ["support", "help", "connect"]

# Link schemes that don't point at a fetchable page and should be skipped.
SKIPPED_HREF_PREFIXES = ("#", "javascript:", "mailto:", "tel:")
# ---------------------------------------------------------------------------


def _normalize(text: str) -> str:
    """Lowercase, unify hyphens/underscores/spaces so 'Contact-Us',
    'contact_us' and 'Contact Us' all compare equal."""
    text = text.lower()
    text = re.sub(r"[-_]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _compile_patterns(keywords: list) -> list:
    seen = set()
    patterns = []
    for kw in keywords:
        norm = _normalize(kw)
        if norm in seen:
            continue
        seen.add(norm)
        patterns.append((kw, re.compile(rf"\b{re.escape(norm)}\b")))
    return patterns


_HIGH_PATTERNS = _compile_patterns(HIGH_CONFIDENCE_KEYWORDS)
_MEDIUM_PATTERNS = _compile_patterns(MEDIUM_CONFIDENCE_KEYWORDS)

CONFIDENCE_RANK = {"high confidence": 0, "medium confidence": 1}


def match_keyword(path: str, link_text: str):
    """Return (matched_keyword, confidence) for a link's path/text, checking
    high-confidence keywords first and falling back to medium-confidence
    ones. Returns (None, None) if nothing matches."""
    fields = [_normalize(unquote(path)), _normalize(link_text)]
    for kw, pattern in _HIGH_PATTERNS:
        if any(pattern.search(field) for field in fields):
            return kw, "high confidence"
    for kw, pattern in _MEDIUM_PATTERNS:
        if any(pattern.search(field) for field in fields):
            return kw, "medium confidence"
    return None, None


def extract_links(html: str, base_url: str) -> list:
    """Return a list of (absolute_url, visible_text) for every real <a href>
    on the page, resolving relative links against base_url."""
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if not href or href.lower().startswith(SKIPPED_HREF_PREFIXES):
            continue
        links.append((urljoin(base_url, href), a.get_text(" ", strip=True)))
    return links


def find_contact_links(url: str, session: requests.Session = None, timeout: int = REQUEST_TIMEOUT_SECONDS) -> dict:
    """Fetch a single URL and look for contact-style links on it.

    Returns a dict:
        {"status": "ok" | "error" | "no_links" | "no_match",
         "detail": str,              # error message, when status == "error"
         "matches": [{"url": ..., "keyword": ..., "confidence": ...}, ...]}
    """
    session = session or requests
    try:
        response = session.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
        response.raise_for_status()
    except requests.exceptions.RequestException as exc:
        return {"status": "error", "detail": str(exc), "matches": []}

    links = extract_links(response.text, response.url)
    if not links:
        return {"status": "no_links", "detail": "Page has no usable <a> tags.", "matches": []}

    best_by_url = {}
    for link_url, link_text in links:
        keyword, confidence = match_keyword(urlparse(link_url).path, link_text)
        if keyword is None:
            continue
        rank = CONFIDENCE_RANK[confidence]
        if link_url not in best_by_url or rank < best_by_url[link_url][2]:
            best_by_url[link_url] = (keyword, confidence, rank)

    if not best_by_url:
        return {"status": "no_match", "detail": "No contact-style link found.", "matches": []}

    matches = sorted(
        (
            {"url": link_url, "keyword": keyword, "confidence": confidence}
            for link_url, (keyword, confidence, _rank) in best_by_url.items()
        ),
        key=lambda m: CONFIDENCE_RANK[m["confidence"]],
    )
    return {"status": "ok", "detail": "", "matches": matches}


def _find_column(df: pd.DataFrame, name_fragments: list):
    for col in df.columns:
        col_norm = str(col).strip().lower()
        if any(fragment in col_norm for fragment in name_fragments):
            return col
    return None


def load_records(path: str) -> list:
    """Read the tracker workbook and return a list of (company, website)
    tuples, skipping rows with no usable website value."""
    df = pd.read_excel(path)

    website_col = _find_column(df, ["website", "url", "site"])
    if website_col is None:
        if df.shape[1] <= FALLBACK_WEBSITE_COLUMN_INDEX:
            raise ValueError(
                f"Could not find a website/url column by header, and the sheet "
                f"has fewer than {FALLBACK_WEBSITE_COLUMN_INDEX + 1} columns "
                f"to fall back to column F."
            )
        website_col = df.columns[FALLBACK_WEBSITE_COLUMN_INDEX]

    company_col = _find_column(df, ["company", "name"]) or df.columns[0]

    records = []
    for _, row in df.iterrows():
        website = str(row.get(website_col, "")).strip()
        if not website or website.lower() == "nan":
            continue
        if not website.lower().startswith(("http://", "https://")):
            website = f"https://{website}"
        company = str(row.get(company_col, "")).strip()
        records.append((company, website))
    return records


def check_websites(records: list, delay: float = DELAY_BETWEEN_REQUESTS_SECONDS) -> list:
    """Run find_contact_links across a batch of (company, website) records,
    one entry per company. Each entry's matches are already sorted best
    confidence first."""
    session = requests.Session()
    results = []
    for i, (company, website) in enumerate(records, start=1):
        print(f"[{i}/{len(records)}] Checking {company or website} -> {website}")
        result = find_contact_links(website, session=session)
        results.append({
            "company": company,
            "website": website,
            "status": result["status"],
            "matches": result["matches"],
        })
        if i < len(records):
            time.sleep(delay)
    return results


def _best_confidence_rank(matches: list) -> int:
    if not matches:
        return len(CONFIDENCE_RANK)  # sorts after every real match
    return min(CONFIDENCE_RANK[m["confidence"]] for m in matches)


def build_output_rows(results: list) -> pd.DataFrame:
    """Consolidate per-company results into one row per company, with a
    Match N Link/Keyword/Confidence triplet per matched link (best
    confidence first), sorted so companies with the strongest match come
    first."""
    ordered = sorted(results, key=lambda r: _best_confidence_rank(r["matches"]))
    max_matches = max((len(r["matches"]) for r in ordered), default=0)

    columns = ["Company", "Website", "Status"]
    for i in range(1, max_matches + 1):
        columns += [f"Match {i} Link", f"Match {i} Keyword", f"Match {i} Confidence"]

    rows = []
    for r in ordered:
        row = {"Company": r["company"], "Website": r["website"], "Status": r["status"]}
        for i, match in enumerate(r["matches"], start=1):
            row[f"Match {i} Link"] = match["url"]
            row[f"Match {i} Keyword"] = match["keyword"]
            row[f"Match {i} Confidence"] = match["confidence"]
        rows.append(row)

    return pd.DataFrame(rows, columns=columns)


def main():
    records = load_records(INPUT_FILE)
    if not records:
        raise ValueError(f"No usable website URLs found in '{INPUT_FILE}'.")

    results = check_websites(records)
    out_df = build_output_rows(results)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"{OUTPUT_FILE_PREFIX}_{timestamp}.xlsx"
    out_df.to_excel(output_file, index=False)
    print(f"\nWrote {len(out_df)} rows to {output_file}")


if __name__ == "__main__":
    main()
