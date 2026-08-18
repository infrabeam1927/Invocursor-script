# Invocursor Outreach Scripts

Tools for running Invocursor's cold outreach to prospect companies: vet each
website, find its contact page, and fill out the contact form for a human to
review and submit — never fully automated, always a manual submit.

## Pipeline

```
Outreach tracker.xlsx
        │
        ▼
security_check.py  ───────►  security_report_<timestamp>.xlsx
        │                           (Risk Level per company)
        ▼
find_contact_links.py  ───►  contact_link_matches_<timestamp>.xlsx
        │                           (best contact-page link per company)
        ▼
main.py  ──────────────────►  opens each qualifying company's contact
                               page and fills it in for you to review
                               and submit manually
```

`security_check.py` and `find_contact_links.py` can be run in either order
(or independently) — `main.py` just needs both of their latest output files
to exist before it runs.

## Setup

```powershell
pip install -r requirements.txt
playwright install chromium
```

Before running anything, open `fill_contact_form.py` and set `PHONE` (and
`EMAIL`, if you want something other than the default) under "FIXED
DETAILS" — these are the same for every company and don't need to change
per run.

## Input file

All scripts read from `Outreach tracker.xlsx` by default (change
`INPUT_FILE` at the top of a script to point elsewhere). Expected layout:

| Column | Contents |
|---|---|
| A | Company name |
| F | Website URL |

A header row is optional — scripts auto-detect "Company"/"Name" and
"Website"/"URL"/"Site" columns by header text, falling back to column F for
the website if no header matches.

## Scripts

### `security_check.py`
Screens every website in the tracker for red flags before you point any
automation at it: reachability, SSL certificate validity/expiry, domain age
(WHOIS), parked/thin-content pages, and redirects to an unrelated domain.
Optionally also checks Google's Safe Browsing blocklist if you set a
`GOOGLE_SAFE_BROWSING_API_KEY` environment variable (free key, never
hardcoded — skipped entirely if unset).

Any single flag marks a site at least **Medium** risk; unreachable, invalid
certificate, blocklisted, or parked-domain flags mark it **High** risk.

```powershell
python security_check.py
```

Writes `security_report_<timestamp>.xlsx` — one row per company, color-coded
red (High) / amber (Medium) / green (Clean), worst first.

### `find_contact_links.py`
Scans each website's homepage for links to a contact-style page (`contact`,
`get-in-touch`, `book-a-demo`, `support`, etc.), matching both the link's URL
and its visible text. High-confidence keywords name a contact/outreach page
specifically; medium-confidence keywords are broader terms that often lead
to one.

```powershell
python find_contact_links.py
```

Writes `contact_link_matches_<timestamp>.xlsx` — one row per company, every
matched link spread across `Match 1`, `Match 2`, ... column triplets
(Link/Keyword/Confidence), best confidence first; companies sorted by their
best match.

### `main.py`
The main pipeline entry point. Reads the **most recently generated** report
from each of the two scripts above (does not re-run them) and filters to
companies that are both:

- **Acceptable**: `Clean` risk level in the security report
- **Doable**: a high-confidence `Match 1` contact link in the link report

For every qualifying company, it opens that contact page in a shared browser
session and fills in Invocursor's outreach message (personalized with the
company name) — then pauses for you to review and submit manually before
moving to the next company. Type `q` at the pause to stop the run early.
Companies that don't qualify are listed with the specific reason (security
risk, no link found, or only a low-confidence link) — nothing is silently
dropped.

Progress is saved to `outreach_progress.json` after every company (reviewed
or errored), scoped to the exact pair of report files used. Run `main.py`
again against the same reports and it offers to resume where you left off
instead of starting over; run it against newer reports and it starts fresh
automatically.

Every company handled (reviewed or errored) also gets a row written to
`Output tracker/Output tracker.xlsx` — Company, Website, Contact URL,
Outcome, Security Risk, Match Confidence, and when it was last handled.
This is a single file that keeps accumulating across every run (not a new
timestamped file each time, unlike the security/contact-link reports):
re-handling a company updates its existing row instead of adding a
duplicate, so it always reflects the latest state of every company that's
ever gone through the pipeline.

```powershell
python main.py
```

### `fill_contact_form.py`
Fills a single company's contact form. Can be run standalone for a one-off
company, or imported by `main.py` to do the same thing in bulk.

```powershell
python fill_contact_form.py
```

Set `TARGET_COMPANY` and `CONTACT_URL` at the top of the file first (the
script raises an error if either is left blank). Opens a visible (non-headless)
browser, fills the form, and **never clicks Submit** — that's on you, once
you've checked it over.

## Notes

- All three checking/filling scripts wait `DELAY_BETWEEN_REQUESTS_SECONDS`
  between requests to be polite to the sites being checked — don't lower
  this by much.
- WHOIS domain-age lookups are frequently rate-limited or blocked
  regardless of a domain's legitimacy — a blank "Domain Age" cell in the
  security report just means WHOIS was unavailable, not a signal on its own.
- Nothing in this pipeline submits a form automatically. Every company still
  gets a manual review and a manual click before anything is sent.
