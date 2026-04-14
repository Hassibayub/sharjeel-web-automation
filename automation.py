import os
import threading
import openpyxl
import requests
from playwright.sync_api import Playwright, sync_playwright
from time import sleep
from PySide6.QtCore import Signal, QObject

STATE_FILE = "state.json"


class AutomationWorker(QThread):
    log_signal = Signal(str)
    progress_signal = Signal(int, int, str)
    finished_signal = Signal()
    error_signal = Signal(str)
    captcha_pending_signal = Signal()
    captcha_resolved_signal = Signal()

    def __init__(self, excel_path, column_name, output_dir, email, password):
        super().__init__()
        self.excel_path = excel_path
        self.column_name = column_name
        self.output_dir = output_dir
        self.email = email
        self.password = password
        self.captcha_event = threading.Event()

    def run(self):
        try:
            run_automation(
                excel_path=self.excel_path,
                column_name=self.column_name,
                output_dir=self.output_dir,
                email=self.email,
                password=self.password,
                progress_callback=self._on_progress,
                log_callback=self._on_log,
                captcha_pending_callback=self._on_captcha_pending,
                captcha_resolved_callback=self._on_captcha_resolved,
                captcha_event=self.captcha_event,
            )
        except Exception as e:
            self.error_signal.emit(str(e))
        finally:
            self.finished_signal.emit()

    def _on_log(self, msg):
        self.log_signal.emit(msg)

    def _on_progress(self, current, total, status=""):
        self.progress_signal.emit(current, total, status)

    def _on_captcha_pending(self):
        self.captcha_pending_signal.emit()

    def _on_captcha_resolved(self):
        self.captcha_resolved_signal.emit()

    def resolve_captcha(self):
        self.captcha_event.set()


def get_ids_from_excel(filepath: str, column_name: str) -> list:
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    if column_name not in headers:
        raise ValueError(f"Column '{column_name}' not found. Available columns: {headers}")

    col_index = headers.index(column_name)
    ids = [row[col_index] for row in ws.values]
    return ids[1:]


def ensure_download_dir(download_dir: str):
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)


def reset_login():
    if os.path.exists(STATE_FILE):
        os.remove(STATE_FILE)
        return True
    return False


def run_automation(
    excel_path: str,
    column_name: str,
    output_dir: str,
    email: str,
    password: str,
    progress_callback=None,
    log_callback=None,
    captcha_pending_callback=None,
    captcha_resolved_callback=None,
    captcha_event=None,
) -> None:
    if captcha_event is None:
        captcha_event = threading.Event()

    def on_captcha_pending():
        captcha_event.clear()
        if captcha_pending_callback:
            captcha_pending_callback()
        captcha_event.wait()
        if captcha_resolved_callback:
            captcha_resolved_callback()

    ensure_download_dir(output_dir)

    def log(msg):
        if log_callback:
            log_callback(msg)
        print(msg)

    def report_progress(current, total, status=""):
        if progress_callback:
            progress_callback(current, total, status)

    with sync_playwright() as playwright:
        log("Starting browser...")
        browser = playwright.chromium.launch(headless=False, slow_mo=1000, args=["--start-maximized"])

        context = browser.new_context(
            viewport=None, accept_downloads=True, storage_state=STATE_FILE if os.path.exists(STATE_FILE) else None
        )

        context.set_default_timeout(30000)
        context.set_default_navigation_timeout(30000)

        page = context.new_page()
        page.goto("https://cms.ric.edu.pk/login")

        if not os.path.exists(STATE_FILE):
            log("No saved session found. Please login manually...")
            page.get_by_role("textbox", name="Email").fill(email)
            page.get_by_role("textbox", name="Password").fill(password)

            frame = page.frame_locator('iframe[title="reCAPTCHA"]')
            frame.locator(".recaptcha-checkbox").wait_for(state="visible")
            frame.locator(".recaptcha-checkbox").click()
            log("Please solve the CAPTCHA manually...")
            on_captcha_pending()

            page.get_by_role("button", name="Log In").click()
            page.wait_for_load_state("load")

            context.storage_state(path=STATE_FILE)
            log("Session saved!")

        log(f"Reading IDs from '{excel_path}', column '{column_name}'...")
        ids = get_ids_from_excel(excel_path, column_name)
        log(f"Found {len(ids)} IDs to process")

        total = len(ids)
        for idx, student_id in enumerate(ids, 1):
            report_progress(idx, total, f"Processing: {student_id}")
            log(f"[{idx}/{total}] Processing: {student_id}")

            url = f"https://cms.ric.edu.pk/exam/print_transcript/pdf/{student_id}"

            cookies = context.cookies()
            session = requests.Session()
            for cookie in cookies:
                session.cookies.set(cookie["name"], cookie["value"], domain=cookie.get("domain", ".cms.ric.edu.pk"))

            try:
                response = session.get(url, timeout=30)
                response.raise_for_status()

                file_path = os.path.join(output_dir, f"{student_id}.pdf")
                with open(file_path, "wb") as f:
                    f.write(response.content)
                log(f"  Saved: {file_path}")
            except requests.RequestException as e:
                log(f"  Failed to download for {student_id}: {e}")

        report_progress(total, total, "Complete!")
        log("Automation complete!")

        context.close()
        browser.close()
