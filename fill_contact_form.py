"""
Fills a prospect's contact form with Invocursor's outreach details using Playwright.

Usage:
    1. Set TARGET_COMPANY and CONTACT_URL below for the company you're reaching out to.
    2. Fill in PHONE (and EMAIL, if different) in the fixed details section.
    3. Run: python fill_contact_form.py
    4. The browser opens visibly and the form is filled in for you to review.
       The script does NOT submit the form — click Submit yourself in the
       browser once you've checked everything looks right.

Note: contact forms vary widely across sites. This script tries a set of
common field patterns (by label, placeholder, name, and id) for each field,
but you may need to fill in or adjust a field manually if a site's form
doesn't match any of the patterns tried.

Can also be imported by another script to fill many companies' forms in one
browser session — see fill_contact_form_page() (fills an existing Playwright
page, no browser lifecycle or review pause) and fill_one_company() (full
standalone single-company flow, used by main() below).
"""

import re

from playwright.sync_api import sync_playwright, Page

# ---------------------------------------------------------------------------
# CHANGE THESE FOR EACH NEW COMPANY (only used when running this file directly)
# ---------------------------------------------------------------------------
TARGET_COMPANY = ""   # e.g. "Acme Property Group"
CONTACT_URL = ""      # e.g. "https://www.acme.com/contact"
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# FIXED DETAILS (same every run)
# ---------------------------------------------------------------------------
NAME = "Nihith Mekala"
COMPANY = "Invocursor"
EMAIL = "Sales@invocursor.com"
PHONE = "(647) 735-1001"

MESSAGE_TEMPLATE = (
    "Invocursor is an AI layer that lives inside vertical SaaS products, sees "
    "the user's live screen, and completes workflows with them — so support "
    "tickets and onboarding friction never happen in the first place. We're "
    "live in production with a childcare-management SaaS (Attendify) since "
    "April 2026, seeing early reductions in repetitive support questions and "
    "onboarding hand-holding. Given {target_company}'s focus on real estate "
    "tech, I'd love to talk about how Invocursor applies to your portfolio's "
    "operator-facing software — property management platforms, leasing "
    "tools, maintenance systems — where non-technical staff hit the same "
    "'how do I…' walls every day. Happy to share a live demo."
)


def build_message(target_company: str) -> str:
    return MESSAGE_TEMPLATE.format(target_company=target_company)
# ---------------------------------------------------------------------------


def _word_pattern(keywords: list) -> str:
    """Build a regex matching any keyword as a whole word/phrase (not as a
    substring of a longer word), so e.g. 'name' won't match inside
    'companyName' or 'firstName'."""
    return "|".join(rf"(?<![a-zA-Z]){re.escape(kw)}(?![a-zA-Z])" for kw in keywords)


def find_field(page: Page, include_keywords: list, used: set, exclude_keywords: list = None, tag: str = "input, textarea"):
    """Scan visible input/textarea elements once and return the first one whose
    name/id/placeholder/aria-label/associated <label> text matches an include
    keyword as a whole word and no exclude keyword, skipping fields already
    claimed by another logical field."""
    include_pattern = _word_pattern(include_keywords)
    exclude_pattern = _word_pattern(exclude_keywords) if exclude_keywords else None

    elements = page.locator(tag)
    for i in range(elements.count()):
        el = elements.nth(i)
        try:
            if not el.is_visible():
                continue

            name_attr = el.get_attribute("name") or ""
            id_attr = el.get_attribute("id") or ""
            placeholder = el.get_attribute("placeholder") or ""
            aria_label = el.get_attribute("aria-label") or ""
            label_text = ""
            if id_attr:
                label_loc = page.locator(f"label[for='{id_attr}']")
                if label_loc.count() > 0:
                    label_text = label_loc.first.inner_text()

            identity = (name_attr, id_attr, placeholder)
            if identity in used:
                continue

            combined = " ".join([name_attr, id_attr, placeholder, aria_label, label_text])
            if not re.search(include_pattern, combined, re.I):
                continue
            if exclude_pattern and re.search(exclude_pattern, combined, re.I):
                continue

            used.add(identity)
            return el
        except Exception:
            continue
    return None


def fill_field(page: Page, include_keywords: list, value: str, field_label: str, used: set, exclude_keywords: list = None, tag: str = "input, textarea") -> bool:
    field = find_field(page, include_keywords, used, exclude_keywords=exclude_keywords, tag=tag)
    if field is None:
        print(f"  [SKIP] Could not find a field for '{field_label}' — fill it in manually.")
        return False
    try:
        field.fill(value)
        print(f"  [OK] Filled '{field_label}'")
        return True
    except Exception as exc:
        print(f"  [SKIP] Found a field for '{field_label}' but couldn't fill it ({exc}) — fill it in manually.")
        return False


def fill_contact_form_page(page: Page, target_company: str, contact_url: str) -> None:
    """Navigate `page` to contact_url and fill in Invocursor's outreach
    details, with the message personalized for target_company. Does not
    manage the browser's lifecycle or wait for a review pause — see
    fill_one_company() for the standalone single-company flow that does."""
    if not target_company or not target_company.strip():
        raise ValueError("target_company must not be blank.")
    if not contact_url or not contact_url.strip():
        raise ValueError("contact_url must not be blank.")

    page.goto(contact_url)
    print(f"Filling contact form for {target_company} at {contact_url}")

    used_fields = set()
    fill_field(page, ["company", "organization", "organisation", "business"], COMPANY, "Company", used_fields)
    fill_field(page, ["email"], EMAIL, "Email", used_fields)
    fill_field(page, ["phone", "mobile", "tel"], PHONE, "Phone", used_fields)
    fill_field(
        page,
        ["name", "full name", "your name"],
        NAME,
        "Name",
        used_fields,
        exclude_keywords=["company", "organization", "organisation", "business", "email", "phone"],
    )
    fill_field(
        page,
        ["message", "comments", "how can we help", "inquiry", "enquiry"],
        build_message(target_company),
        "Message",
        used_fields,
        tag="textarea, input",
    )


def fill_one_company(target_company: str, contact_url: str) -> None:
    """Standalone single-company flow: launches its own browser, fills the
    form, and waits for the user to review/submit before closing."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        try:
            page = browser.new_page()
            fill_contact_form_page(page, target_company, contact_url)

            print("\nForm filled. Review it in the browser window.")
            print("This script will NOT click Submit — submit it yourself once you're happy with it.")
            input("Press Enter here once you're done (this will close the browser)... ")
        finally:
            browser.close()


def main():
    if not TARGET_COMPANY.strip():
        raise ValueError("TARGET_COMPANY must be set before running this script.")
    if not CONTACT_URL.strip():
        raise ValueError("CONTACT_URL must be set before running this script.")
    fill_one_company(TARGET_COMPANY, CONTACT_URL)


if __name__ == "__main__":
    main()
