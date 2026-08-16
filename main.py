"""
Interactive entry point for filling a single company's contact form with
Invocursor's outreach details.

Prompts you step by step for the company name and its contact-page URL,
opens the form in a visible browser and fills it in, then pauses so you can
review it and submit it yourself in the browser — this script never clicks
Submit for you. Offers to do another company afterward, reusing the same
browser session.

For checking many companies in bulk against the tracker (security check,
finding contact pages, filling many forms in one run), see
run_outreach_pipeline.py instead. This script is for one-off / ad hoc runs.

Usage:
    python main.py
"""

from playwright.sync_api import sync_playwright

from fill_contact_form import fill_contact_form_page


def prompt_nonblank(prompt: str) -> str:
    while True:
        value = input(prompt).strip()
        if value:
            return value
        print("This can't be blank — try again.")


def prompt_url(prompt: str) -> str:
    url = prompt_nonblank(prompt)
    if not url.lower().startswith(("http://", "https://")):
        url = f"https://{url}"
    return url


def run_one_company(browser) -> None:
    target_company = prompt_nonblank("Company name: ")
    contact_url = prompt_url("Contact page URL: ")

    page = browser.new_page()
    try:
        fill_contact_form_page(page, target_company, contact_url)
        print("\nForm filled. Review it in the browser window.")
        print("This will NOT auto-submit — submit it yourself once you're happy with it.")
        input("Press Enter here once you've submitted (or decided not to)... ")
    except Exception as exc:
        print(f"[ERROR] {exc}")
    finally:
        page.close()


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        try:
            while True:
                run_one_company(browser)
                again = input("\nFill another company's form? (y/N): ").strip().lower()
                if again not in ("y", "yes"):
                    break
        finally:
            browser.close()

    print("Done.")


if __name__ == "__main__":
    main()
