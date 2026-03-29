# CRM Auto-Entry

Automated lead entry into the Fasttrack CRM via Selenium for my documentary outreach. Searches for existing contacts and organisations, creates new ones where needed, and adds notes and call records for each lead.

## How It Works

For each lead the script:

1. **Searches** the CRM by last name
2. **Matches** the contact by first name + institution
3. If found: opens their record, creates a **New Note** with the outreach text, then creates a **New Call** (Intro Email, Low priority)
4. If the contact isn't found but the company exists: **adds the contact** to the existing organisation, then does the note + call
5. If neither exists: creates a **New Company** (with switchboard placeholder), adds the contact, then does the note + call

## Usage

### Quick Start — Paste & Run

1. Open `leads.txt` and paste your leads (one per line):
   ```
   Institution | Position | Contact Name | Email | Paste in CRM text
   ```
   Use `\n` for newlines within the paste text.

2. Run:
   ```bash
   python3 crm_entry.py --file leads.txt --dry-run   # preview
   python3 crm_entry.py --file leads.txt              # execute
   ```

### From Excel

```bash
python3 crm_entry.py --rows 132-146      # process a range of rows
python3 crm_entry.py --row 139           # single row
python3 crm_entry.py --dry-run           # preview all default rows
```

## Requirements

- Python 3.8+
- `selenium` — `pip install selenium`
- `openpyxl` — `pip install openpyxl`
- ChromeDriver (matching your Chrome version)
