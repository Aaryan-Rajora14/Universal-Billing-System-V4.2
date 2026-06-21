# Universal Billing System v5.0

<img width="1920" height="1020" alt="Login · Billing System and 1 more page - Personal - Microsoft​ Edge 21-06-2026 08_54_56" src="https://github.com/user-attachments/assets/5824055c-0fd4-4209-825b-2cf9ef01b3a8" />

A secure, offline billing and invoice system with fixed bugs and glitches built with Flask and Excel. This project provides a desktop-friendly billing dashboard with PDF bill generation, QR codes, item returns, audit logging, admin controls, and configurable store settings.

## Key Features

- Admin and staff login with role-based access
- Create bills with live preview, item-level discount, and automatic totals
- Generate printable PDF invoices with QR code receipts
- Store bills in `store_bills.xlsx` with separate sheets for Bills, Returns, Profit/Loss, and Settings
- Process item returns and track return history
- Export Excel data for offline reporting
- Admin-only controls for deleting bills, adjusting profit margin, and timeout settings
- Secure session handling, CSRF protection, and audit logging

## Default Users

- Admin: `admin` / `Admin@1234`
- Staff: `staff` / `Staff@1234`

> Change the credentials after the first login for a production deployment.

## Prerequisites

- Python 3.11+ (or compatible Python 3.x)
- pip

## Setup

1. Open a terminal in the project root.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Start the app:

- Windows: run `run_windows.bat`
- Cross-platform: run `python run_app.py`

4. Open the browser if it does not open automatically:

```text
http://localhost:5004
```

## Project Structure

- `app.py` — Flask application and route handlers
- `database.py` — Excel-backed storage manager for bills, returns, settings, and reports
- `run_app.py` — Cross-platform launcher that opens the browser automatically
- `run_windows.bat` — Windows launcher that installs dependencies and starts the app
- `requirements.txt` — Python dependencies
- `templates/` — HTML templates for login and dashboard
- `static/` — CSS, JavaScript, generated bills, and QR code assets

## Usage

- Login as Admin or Staff
- Add customer details and bill items
- Generate a bill to create a PDF invoice and record the sale
- Use the Bill History page to view past invoices
- Admin users can export Excel, adjust settings, and erase data
- Use the Settings page to update store details, theme, and password

## Configuration

The app stores configuration values in the Excel file at `store_bills.xlsx`.

Settings include:

- `store_name`
- `store_tagline`
- `store_address`
- `store_phone`
- `store_email`
- `store_website`
- `currency_symbol`
- `profit_margin`
- `session_timeout_minutes`

## Notes

- `store_bills.xlsx` is created automatically when the app first starts.
- Generated PDF invoices are saved under `static/bills/`.
- QR codes are saved under `static/qr_codes/`.
- Audit events are logged to `audit.log`.

## Security Recommendations

- Replace the default user passwords before using in a real environment
- Set the environment variable `SECRET_KEY` for the Flask app to persist a secure secret key
- Do not expose this app to the public internet without additional hardening

## Troubleshooting

- If Python is not found on Windows, ensure it is installed and added to `PATH`
- If dependencies fail, run `pip install -r requirements.txt` again
- If the browser does not open automatically, visit `http://localhost:5004`

## License

This project does not include a license file. Add one if you plan to share it publicly.
