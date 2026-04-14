import os
import openpyxl
import requests
from playwright.sync_api import Playwright, sync_playwright
from time import sleep

SLOW_MO = 1000
DOWNLOAD_DIR = "./downloads"
STATE_FILE = "state.json"


def get_ids_from_excel(filepath: str) -> list:
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    ids = [row[0] for row in ws.values]
    return ids[1:]


def ensure_download_dir():
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)


def run(playwright: Playwright) -> None:
    ensure_download_dir()

    browser = playwright.chromium.launch(headless=False, slow_mo=SLOW_MO, args=["--start-maximized"])

    context = browser.new_context(
        viewport=None, accept_downloads=True, storage_state=STATE_FILE if os.path.exists(STATE_FILE) else None
    )

    context.set_default_timeout(30000)
    context.set_default_navigation_timeout(30000)

    page = context.new_page()

    page.goto("https://cms.ric.edu.pk/login")

    # ✅ Only login if session not saved
    if not os.path.exists(STATE_FILE):
        page.get_by_role("textbox", name="Email").fill("muhammad.nadeem@ric.edu.pk")
        page.get_by_role("textbox", name="Password").fill("Nadeem555")

        # captcha manual pause (best approach)
        # input("Solve captcha manually, then press Enter...")
        frame = page.frame_locator('iframe[title="reCAPTCHA"]')

        frame.locator(".recaptcha-checkbox").wait_for(state="visible")
        frame.locator(".recaptcha-checkbox").click()
        sleep(10)

        page.get_by_role("button", name="Log In").click()
        page.wait_for_load_state("load")

        # ✅ Save session
        context.storage_state(path=STATE_FILE)
        print("Session saved!")

    ids = get_ids_from_excel("practice print.xlsx")

    print(f"Found {len(ids)} IDs to process")

    for student_id in ids:
        print(f"Processing: {student_id}")
        url = f"https://cms.ric.edu.pk/exam/print_transcript/pdf/{student_id}"

        cookies = context.cookies()
        session = requests.Session()
        for cookie in cookies:
            session.cookies.set(cookie["name"], cookie["value"], domain=cookie.get("domain", ".cms.ric.edu.pk"))

        try:
            response = session.get(url, timeout=30)
            response.raise_for_status()

            file_path = os.path.join(DOWNLOAD_DIR, f"{student_id}.pdf")
            with open(file_path, "wb") as f:
                f.write(response.content)
            print(f"Saved: {file_path}")
        except requests.RequestException as e:
            print(f"Failed to download for {student_id}: {e}")

    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
