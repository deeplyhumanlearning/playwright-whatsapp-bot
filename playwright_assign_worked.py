import json
import random
import re
import time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

BASE_DIR = Path(__file__).resolve().parent
CONTACTS_FILE = BASE_DIR / "contacts.xlsx"
OUTPUT_DIR = BASE_DIR / "output"
SCREENSHOT_DIR = OUTPUT_DIR / "screenshots"
SESSION_DIR = BASE_DIR / "whatsapp_session"

OUTPUT_DIR.mkdir(exist_ok=True)
SCREENSHOT_DIR.mkdir(exist_ok=True)


def pause(page, minimum=0.5, maximum=1.0):
    page.wait_for_timeout(random.randint(int(minimum * 1000), int(maximum * 1000)))


def clean_phone(phone):
    return re.sub(r"\D", "", str(phone))


def normalize_phone(phone):
    digits = clean_phone(phone)

    if digits.startswith("00"):
        digits = digits[2:]

    if digits.startswith("0"):
        digits = digits.lstrip("0")

    if not digits.startswith("91"):
        digits = "91" + digits

    return digits


def read_contacts():
    if not CONTACTS_FILE.exists():
        raise FileNotFoundError(f"Missing {CONTACTS_FILE.name}")

    workbook = load_workbook(CONTACTS_FILE, read_only=True, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()

    if not rows:
        return []

    headers = [str(value).strip() if value is not None else "" for value in rows[0]]
    required = {"Name", "Phone", "Message"}

    if not required.issubset(headers):
        raise ValueError("contacts.xlsx must contain Name, Phone and Message columns")

    indexes = {header: headers.index(header) for header in required}
    contacts = []

    for row in rows[1:]:
        name = row[indexes["Name"]] if len(row) > indexes["Name"] else None
        phone = row[indexes["Phone"]] if len(row) > indexes["Phone"] else None
        message = row[indexes["Message"]] if len(row) > indexes["Message"] else None

        if not name or not phone:
            continue

        contacts.append(
            {
                "name": str(name).strip(),
                "phone": str(phone).strip(),
                "message": str(message).strip() if message else "",
            }
        )

    return contacts


def find_search_box(page, timeout=10000):
    selectors = [
        'input[placeholder="Search or start a new chat"]',
        'input[title="Search or start a new chat"]',
        'div[role="textbox"][aria-label*="Search"]',
        'div[contenteditable="true"][aria-label*="Search"]',
        'div[contenteditable="true"][data-tab="3"]',
    ]

    for selector in selectors:
        locator = page.locator(selector).first
        try:
            locator.wait_for(state="visible", timeout=timeout)
            return locator
        except PlaywrightTimeoutError:
            continue

    raise PlaywrightTimeoutError("WhatsApp search box was not found")


def wait_for_whatsapp(page):
    find_search_box(page, timeout=120000)


def open_contact(page, name, phone):
    phone_digits = normalize_phone(phone)

    page.goto(
        f"https://web.whatsapp.com/send?phone={phone_digits}",
        wait_until="domcontentloaded",
    )

    pause(page, 0.8, 1.2)

    try:
        message_box = find_message_box(page, timeout=10000)
        message_box.scroll_into_view_if_needed()
        message_box.click()
        return True
    except PlaywrightTimeoutError:
        pass

    error_selectors = [
        'div[role="dialog"]',
        'div[aria-label*="phone number"]',
        'div[aria-label*="not on WhatsApp"]',
    ]

    for selector in error_selectors:
        locator = page.locator(selector)
        if locator.count():
            try:
                text = locator.last.inner_text(timeout=2000).strip()
                if text:
                    raise ValueError(f"Unable to open chat for {phone}: {text}")
            except PlaywrightTimeoutError:
                continue

    raise ValueError(f"Unable to open chat for phone number {phone}")


def find_message_box(page, timeout=10000):
    selectors = [
        'div[contenteditable="true"][aria-label="Type a message"]',
        'div[contenteditable="true"][data-tab="10"]',
        'div[contenteditable="true"][role="textbox"]',
    ]

    for selector in selectors:
        locator = page.locator(selector).last
        try:
            locator.wait_for(state="visible", timeout=timeout)
            return locator
        except PlaywrightTimeoutError:
            continue

    raise PlaywrightTimeoutError("Message box was not found")


def find_sent_message(page, message, timeout=15000):
    """
    Find the message we just sent.

    WhatsApp Web's DOM changes frequently, so do not rely on a single
    selector such as div.message-out. Try several known outbound-message
    patterns and verify the actual message text.
    """
    selectors = [
        'div.message-out',
        'div[data-testid="msg-container"]',
        'div[data-id^="true_"]',
        'div[data-id*="-true"]',
    ]

    deadline = time.monotonic() + (timeout / 1000)

    while time.monotonic() < deadline:
        for selector in selectors:
            locator = page.locator(selector)

            try:
                count = locator.count()
            except Exception:
                continue

            for index in range(count - 1, max(-1, count - 10), -1):
                try:
                    candidate = locator.nth(index)
                    if not candidate.is_visible():
                        continue

                    text = candidate.inner_text(timeout=1000).strip()

                    # Compare normalized text because WhatsApp may add
                    # whitespace or line breaks around the message.
                    normalized_actual = re.sub(r"\\s+", " ", text).strip()
                    normalized_expected = re.sub(r"\\s+", " ", message).strip()

                    if normalized_expected and normalized_expected in normalized_actual:
                        return candidate

                except Exception:
                    continue

        page.wait_for_timeout(250)

    return None


def send_message(page, message):
    message_box = find_message_box(page)

    message_box.click()
    message_box.fill(message)
    pause(page)

    # Capture the message count before pressing Enter. This helps us
    # distinguish a newly sent message from an older identical message.
    before_count = page.locator("div.message-out").count()

    message_box.press("Enter")

    # WhatsApp can take a moment to update the DOM after Enter.
    page.wait_for_timeout(400)

    sent = find_sent_message(page, message, timeout=8000)

    if sent:
        return sent

    # Fallback verification: after sending, the composer should normally
    # be empty. This is useful when WhatsApp changes the outbound bubble DOM.
    try:
        current_text = message_box.inner_text().strip()
    except Exception:
        current_text = ""

    if not current_text:
        # Take a temporary screenshot so a failed verification can be
        # diagnosed without falsely reporting the message as sent.
        debug_path = SCREENSHOT_DIR / (
            f"send_verification_failed_{datetime.now().strftime('%H%M%S')}.png"
        )
        page.screenshot(path=str(debug_path))
        raise ValueError(
            "Message was entered and the composer is empty, but the sent-message "
            "confirmation could not be identified. Debug screenshot: "
            f"{debug_path.relative_to(BASE_DIR)}"
        )

    raise ValueError(
        "Message was not confirmed as sent and remains in the message box"
    )


def extract_messages(page):
    messages = page.locator("div.message-in")
    count = messages.count()
    extracted = []

    for index in range(max(0, count - 3), count):
        text = messages.nth(index).inner_text().strip()
        if text:
            extracted.append({"sender": "contact", "message": text})

    return extracted


def save_report(results, run_date):
    json_file = OUTPUT_DIR / f"whatsapp_report_{run_date}.json"
    xlsx_file = OUTPUT_DIR / f"whatsapp_report_{run_date}.xlsx"

    with json_file.open("w", encoding="utf-8") as file:
        json.dump(
            {
                "run_date": run_date,
                "contacts": results,
            },
            file,
            indent=2,
            ensure_ascii=False,
        )

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report"
    sheet.append(["Name", "Phone", "Status", "Message", "Last 3 Messages", "Error"])

    for result in results:
        last_messages = " | ".join(
            f"{item['sender']}: {item['message']}"
            for item in result["last_3_messages"]
        )

        sheet.append(
            [
                result["name"],
                result["phone"],
                result["status"],
                result["message"],
                last_messages,
                result["error"],
            ]
        )

    workbook.save(xlsx_file)
    return json_file, xlsx_file


def main():
    contacts = read_contacts()

    if not contacts:
        print("No contacts found in contacts.xlsx")
        return

    run_date = datetime.now().strftime("%Y-%m-%d")
    results = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch_persistent_context(
            str(SESSION_DIR),
            headless=False,
            viewport={"width": 1280, "height": 900},
        )

        page = browser.pages[0] if browser.pages else browser.new_page()
        page.goto("https://web.whatsapp.com", wait_until="domcontentloaded")

        print("Waiting for WhatsApp Web. Scan the QR code on the first run.")

        try:
            wait_for_whatsapp(page)
        except PlaywrightTimeoutError:
            browser.close()
            raise RuntimeError("WhatsApp Web login timed out")

        for contact in contacts:
            result = {
                "name": contact["name"],
                "phone": contact["phone"],
                "message": "",
                "status": "failed",
                "last_3_messages": [],
                "screenshot": "",
                "error": "",
            }

            try:
                print(f"Processing: {contact['name']} ({contact['phone']})")

                open_contact(page, contact["name"], contact["phone"])
                print("  Contact opened")

                message = contact["message"].replace("{name}", contact["name"]).strip()

                if not message:
                    raise ValueError("Message is empty")

                result["message"] = message

                sent_message = send_message(page, message)
                print("  Message sent")

                # Keep the latest sent message visible and centered in the
                # viewport so the screenshot clearly shows the active chat.
                try:
                    sent_message.scroll_into_view_if_needed()
                    page.wait_for_timeout(300)
                except Exception:
                    pass

                screenshot_name = (
                    f"{datetime.now().strftime('%H%M%S')}_"
                    f"{clean_phone(contact['phone'])}.png"
                )
                screenshot_path = SCREENSHOT_DIR / screenshot_name
                sent_message.screenshot(path=str(screenshot_path))

                result["screenshot"] = str(screenshot_path.relative_to(BASE_DIR))
                result["status"] = "sent"
                result["last_3_messages"] = extract_messages(page)

            except Exception as error:
                result["error"] = str(error)
                print(f"  Failed: {error}")

            results.append(result)
            pause(page, 0.5, 0.8)

        json_file, xlsx_file = save_report(results, run_date)

        print(f"JSON report: {json_file}")
        print(f"Excel report: {xlsx_file}")

        browser.close()


if __name__ == "__main__":
    main()
