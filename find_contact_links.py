"""
Scans prospect websites for links to a "contact us"-style page (contact,
get-in-touch, support, book-a-demo, etc.) and writes matches to an Excel file.

Input:
    An Excel workbook (default: "Outreach tracker.xlsx") with a company name
    column and a website column. Column F is expected to hold the website
    URL; the script also looks for a header containing "website"/"url" and
    a header containing "company"/"name" and prefers those if present.

Output:
    An Excel workbook (default: "contact_link_matches.xlsx") with one row
    per matched link: Company, Website, Matched Link, Matched Keyword,
    Confidence, Status.

Can also be used as a library for a single URL or a list of URLs:

    from find_contact_links import find_contact_links
    matches = find_contact_links("https://example.com")

Usage:
    python find_contact_links.py
"""

import re
import time
from urllib.parse import urljoin, urlparse, unquote

import pandas as pd
import requests
from bs4 import BeautifulSoup

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
INPUT_FILE = "Outreach tracker.xlsx"
OUTPUT_FILE = "contact_link_matches.xlsx"
FALLBACK_WEBSITE_COLUMN_INDEX = 5  # column F, used if no "website"/"url" header is found

REQUEST_TIMEOUT_SECONDS = 10
DELAY_BETWEEN_REQUESTS_SECONDS = 1.5

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
HEADERS = {"User-Agent": USER_AGENT}

# Keywords matched against the link's URL path and visible text.
# "Exact" keywords are page types that are unambiguously contact-related.
# "Fallback" keywords (e.g. "about") sometimes hold contact info but are a
# much weaker signal, hence the lower confidence tier.
EXACT_KEYWORDS = [
    "contact", "contact us", "contact-us", "contactus",
    "get in touch", "get-in-touch", "getintouch",
    "reach us", "reach-us", "reach out", "reach-out", "reachout",
    "support", "help",
    "connect",
    "book a call", "book-a-call",
    "book a demo", "book-a-demo",
    "request a demo", "request-a-demo",
    "talk to us", "talk-to-us",
    "say hello", "say-hello",
]
FALLBACK_KEYWORDS = ["about", "about us", "about-us", "aboutus"]

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


_EXACT_PATTERNS = _compile_patterns(EXACT_KEYWORDS)
_FALLBACK_PATTERNS = _compile_patterns(FALLBACK_KEYWORDS)


def match_keyword(path: str, link_text: str):
    """Return (matched_keyword, confidence) for a link's path/text, checking
    exact keywords first and falling back to lower-confidence ones. Returns
    (None, None) if nothing matches."""
    fields = [_normalize(unquote(path)), _normalize(link_text)]
    for kw, pattern in _EXACT_PATTERNS:
        if any(pattern.search(field) for field in fields):
            return kw, "exact match"
    for kw, pattern in _FALLBACK_PATTERNS:
        if any(pattern.search(field) for field in fields):
            return kw, "fallback match"
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
        rank = 0 if confidence == "exact match" else 1
        if link_url not in best_by_url or rank < best_by_url[link_url][2]:
            best_by_url[link_url] = (keyword, confidence, rank)

    if not best_by_url:
        return {"status": "no_match", "detail": "No contact-style link found.", "matches": []}

    matches = [
        {"url": link_url, "keyword": keyword, "confidence": confidence}
        for link_url, (keyword, confidence, _rank) in best_by_url.items()
    ]
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
    """Run find_contact_links across a batch of (company, website) records
    and flatten the results into one row per matched link (or one status
    row for sites with no match/error)."""
    session = requests.Session()
    rows = []
    for i, (company, website) in enumerate(records, start=1):
        print(f"[{i}/{len(records)}] Checking {company or website} -> {website}")
        result = find_contact_links(website, session=session)

        if not result["matches"]:
            rows.append({
                "Company": company,
                "Website": website,
                "Matched Link": "",
                "Matched Keyword": "",
                "Confidence": "",
                "Status": result["status"] if result["status"] != "ok" else "no_match",
            })
        else:
            for match in result["matches"]:
                rows.append({
                    "Company": company,
                    "Website": website,
                    "Matched Link": match["url"],
                    "Matched Keyword": match["keyword"],
                    "Confidence": match["confidence"],
                    "Status": "ok",
                })

        if i < len(records):
            time.sleep(delay)
    return rows


def main():
    records = load_records(INPUT_FILE)
    if not records:
        raise ValueError(f"No usable website URLs found in '{INPUT_FILE}'.")

    rows = check_websites(records)

    out_df = pd.DataFrame(rows, columns=[
        "Company", "Website", "Matched Link", "Matched Keyword", "Confidence", "Status",
    ])
    out_df.to_excel(OUTPUT_FILE, index=False)
    print(f"\nWrote {len(out_df)} rows to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
