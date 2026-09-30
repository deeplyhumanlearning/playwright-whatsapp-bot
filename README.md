# WhatsApp Web Automation Bot

Playwright-based WhatsApp Web automation for personalized messaging and chat data extraction.

The bot reads contact information from an Excel-compatible spreadsheet, sends personalized messages through WhatsApp Web, captures screenshots of sent messages, extracts recent incoming messages, and generates JSON and Excel reports.

## Features

- WhatsApp Web automation using Playwright
- Manual QR-code authentication on the first run
- Contact and message management through `contacts.xlsx`
- Personalized messages using `{name}` placeholders
- Contact search by name or phone number
- Randomized delays between actions
- Screenshot capture after successful message sending
- Extraction of the last three incoming messages
- Error handling for unavailable contacts and failed actions
- JSON and Excel report generation
- Persistent browser session between runs

## Requirements

- Python 3.9+
- Google Chrome/Chromium-compatible environment
- LibreOffice Calc or Microsoft Excel for editing `.xlsx` files
- A WhatsApp account
- Internet connection

Microsoft Excel is not required. The project uses the standard `.xlsx` format, which can be edited using LibreOffice Calc.

## Project Structure

```text
playwright-whatsapp-bot/
├── playwright_assign.py
├── contacts.xlsx
├── requirements.txt
├── README.md
├── .gitignore
└── output/
    └── screenshots/
```

The browser session is stored locally and is excluded from Git.

## Setup

### 1. Create a virtual environment

Windows PowerShell:

```powershell
python -m venv .venv
```

Activate it:

```powershell
.venv\Scripts\Activate.ps1
```

### 2. Install dependencies

```powershell
pip install -r requirements.txt
```

Install the Playwright browser:

```powershell
playwright install
```

### 3. Configure contacts

Open `contacts.xlsx` using LibreOffice Calc.

The spreadsheet should contain these columns:

| Name | Phone | Message |
|---|---|---|
| TestUser | +919876543210 | Hi {name}, this is a test message. |

`Phone` should include the country code.

The `{name}` placeholder is automatically replaced with the contact's name.

For example:

```text
Hi {name}, this is a test message.
```

becomes:

```text
Hi TestUser, this is a test message.
```

Save the file as:

```text
contacts.xlsx
```

## Run

Activate the virtual environment and run:

```powershell
python playwright_assign.py
```

On the first run, WhatsApp Web will display a QR code.

Open WhatsApp on your phone and go to:

```text
WhatsApp → Linked Devices → Link a Device
```

Scan the QR code.

After authentication, the browser session is saved locally so subsequent runs can reuse the session.

## Output

After a successful run, the `output` folder contains dated reports and screenshots:

```text
output/
├── whatsapp_report_YYYY-MM-DD.json
├── whatsapp_report_YYYY-MM-DD.xlsx
└── screenshots/
    ├── ...
    └── ...
```

### JSON Report

The JSON report contains detailed information for each contact, including:

- Contact details
- Message
- Sending status
- Error information, when applicable
- Screenshot path
- Extracted recent messages

### Excel Report

The Excel report provides a summary of the run and can be opened using LibreOffice Calc.

### Screenshots

A screenshot is captured after a message is successfully sent.

## Error Handling

The bot is designed to continue processing when an individual contact cannot be found or an action fails.

The result for each contact is recorded in the report instead of terminating the entire run.

## Responsible Use

Use this automation only with contacts who have agreed to receive messages.

Avoid sending unsolicited or excessive messages. Randomized delays are included between actions to make the automation less aggressive.

## Development

This project was developed as a practical exercise in:

- Browser automation
- Playwright
- Python
- Dynamic web-page interaction
- Data extraction
- Spreadsheet processing
- JSON reporting
- Error handling

## License

This project is intended for personal learning and development purposes.
