<img src="https://www.upload.ee/image/19807469/2026-10-02_031006.png" border="0" alt="2026-10-02_031006.png" />

# Mosque Management Professional v4.0

A PyQt6 desktop application for managing mosque operations.

## Features

- User accounts with role permissions
- Secure PBKDF2-SHA256 password storage
- SQLite database management
- Dashboard and reports
- Income and expense tracking
- Recurring transactions
- Activity reminders
- Asset maintenance schedules
- Import/export support
- Backup and restore tools
- Dark interface

## Installation

```bash
git clone <your-repository-url>
cd Mosque_Management_v4.0_GitHub_Project
python -m venv .venv
```

Activate environment:

Windows:
```bash
.venv\Scripts\activate
```

Linux/macOS:
```bash
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Run:

```bash
python Mosque_Management_v4.0_en_US.py
```

## Optional

Build executable:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed Mosque_Management_v4.0_en_US.py
```

## License

Choose a license before publishing this repository.
