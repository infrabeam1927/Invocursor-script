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
"""

from playwright.sync_api import sync_playwright, Page, TimeoutError as PlaywrightTimeoutError

# ---------------------------------------------------------------------------
# CHANGE THESE FOR EACH NEW COMPANY
# ---------------------------------------------------------------------------
TARGET_COMPANY = ""   # e.g. "Acme Property Group"
CONTACT_URL = ""      # e.g. "https://www.acme.com/contact"
# ---------------------------------------------------------------------------

if not TARGET_COMPANY.strip():
    raise ValueError("TARGET_COMPANY must be set before running this script.")
if not CONTACT_URL.strip():
    raise ValueError("CONTACT_URL must be set before running this script.")

# ---------------------------------------------------------------------------
# FIXED DETAILS (same every run)
# ---------------------------------------------------------------------------
NAME = "Nihith Mekala"
COMPANY = "Invocursor"
EMAIL = "Sales@invocursor.com"
PHONE = "REPLACE_ME"  # <-- put your phone number here before running

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

MESSAGE = MESSAGE_TEMPLATE.format(target_company=TARGET_COMPANY)
# ---------------------------------------------------------------------------


def fill_first_match(page: Page, locators: list, value: str, field_label: str) -> bool:
    """Try a list of Playwright locators in order and fill the first visible match."""
    for locator in locators:
        try:
            if locator.count() > 0 and locator.first.is_visible():
                locator.first.fill(value)
                print(f"  [OK] Filled '{field_label}'")
                return True
        except PlaywrightTimeoutError:
            continue
        except Exception:
            continue
    print(f"  [SKIP] Could not find a field for '{field_label}' — fill it in manually.")
    return False


def build_locators(page: Page, keywords: list, tag: str = "input, textarea"):
    """Build a list of candidate locators for a field based on common keywords."""
    locators = []
    for kw in keywords:
        locators.append(page.get_by_label(kw, exact=False))
        locators.append(page.get_by_placeholder(kw, exact=False))
        locators.append(page.locator(f"{tag}[name*='{kw}' i]"))
        locators.append(page.locator(f"{tag}[id*='{kw}' i]"))
    return locators


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto(CONTACT_URL)

        print(f"Filling contact form for {TARGET_COMPANY} at {CONTACT_URL}")

        fill_first_match(page, build_locators(page, ["name", "full name", "your name"]), NAME, "Name")
        fill_first_match(page, build_locators(page, ["company", "organization", "business"]), COMPANY, "Company")
        fill_first_match(page, build_locators(page, ["email"]), EMAIL, "Email")
        fill_first_match(page, build_locators(page, ["phone", "mobile", "tel"]), PHONE, "Phone")
        fill_first_match(
            page,
            build_locators(page, ["message", "comments", "how can we help", "inquiry"], tag="textarea, input"),
            MESSAGE,
            "Message",
        )

        print("\nForm filled. Review it in the browser window.")
        print("This script will NOT click Submit — submit it yourself once you're happy with it.")
        input("Press Enter here once you're done (this will close the browser)... ")

        browser.close()


if __name__ == "__main__":
    main()
