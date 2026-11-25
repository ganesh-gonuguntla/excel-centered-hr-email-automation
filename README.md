# 📧 Internship Mailer

A production-quality CLI tool that reads HR/recruiter contacts from an Excel file and sends personalized internship-request emails via SMTP (Gmail).

## Features

- **Excel/CSV input** with configurable column mapping
- **Email validation**: syntax check + DNS MX record lookup
- **Smart personalization**: first-name extraction, honorific handling, template selection by role
- **Resume attachment**: automatically attached for high-scored contacts (score ≥ 3)
- **Duplicate detection**: never double-sends (tracks via `sent_log.csv`)
- **Resume from interruption**: re-running picks up where you left off
- **Throttling & daily cap**: randomized delays + configurable daily limit
- **Retry logic**: transient SMTP errors retried with exponential backoff
- **Rich CLI output**: progress bar, color-coded stats table
- **Dry-run & test-send**: validate everything before sending for real
- **Follow-up emails**: send polite follow-ups to contacts after N days

---

## 🚀 Quick Start

### 1. Install Python 3.10+

Download from [python.org](https://www.python.org/downloads/) if not installed.

### 2. Clone & Create Virtual Environment

```bash
cd "excel driven hr email automation"
python -m venv venv

# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Enable Gmail App Password

1. Go to [Google Account Security](https://myaccount.google.com/security)
2. Enable **2-Step Verification** (required)
3. Go to [App Passwords](https://myaccount.google.com/apppasswords)
4. Select "Mail" and your device, then click **Generate**
5. Copy the **16-character password** (e.g., `abcd efgh ijkl mnop`)

> **⚠️ Important**: This is NOT your regular Gmail password. It's a special app-specific password.

### 5. Configure `.env`

```bash
copy .env.example .env
```

Edit `.env`:
```
SMTP_USER=your.email@gmail.com
SMTP_APP_PASSWORD=abcd efgh ijkl mnop
```

### 6. Configure `profile.yaml`

```bash
copy profile.yaml.example profile.yaml
```

Open `profile.yaml` and replace every `REPLACE_ME` with your actual information. (Both `profile.yaml` and `.env` are ignored by git in `.gitignore` so your personal details and credentials will never be committed to GitHub.)

### 7. Preview Emails

```bash
python main.py preview --file contacts.xlsx
```

This renders 3 sample emails to the terminal so you can proofread the content.

### 8. Test Send

```bash
python main.py test-send --file contacts.xlsx
```

Sends ONE email **to yourself** using the first valid contact's data. Check your inbox for formatting.

### 9. Dry Run

```bash
python main.py run --file contacts.xlsx --dry-run
```

Runs the full pipeline (validation, personalization, logging) without actually sending emails.

### 10. Real Run

```bash
python main.py run --file contacts.xlsx
```

---

## 📋 CLI Reference

### `run` – Send emails

```bash
python main.py run --file contacts.xlsx [OPTIONS]
```

| Option | Description |
|--------|-------------|
| `--file`, `-f` | Path to contacts Excel/CSV file (required) |
| `--sheet`, `-s` | Sheet name (default: first sheet) |
| `--limit`, `-l` | Max number of rows to process |
| `--min-score` | Minimum Intern Interest Score (1-5) |
| `--start-row` | Start from this row (1-based, default: 1) |
| `--dry-run` | Validate and log, but don't send |
| `--retry-failed` | Only retry previously FAILED rows |
| `--mask` | Mask email addresses in console output |
| `--resume`, `-r` | Path to resume PDF to attach |

### `preview` – Preview emails

```bash
python main.py preview --file contacts.xlsx [--count 3]
```

### `test-send` – Send one test email to yourself

```bash
python main.py test-send --file contacts.xlsx
```

### `stats` – View sent_log.csv statistics

```bash
python main.py stats
```

### `followup` – Send follow-up emails

```bash
python main.py followup --file contacts.xlsx [--days 7] [--dry-run]
```

---

## 📁 Project Structure

```
internship-mailer/
├── main.py                  # CLI entry point
├── mailer/
│   ├── __init__.py
│   ├── reader.py            # Load + normalize Excel rows
│   ├── validator.py         # Email validation (syntax + MX)
│   ├── personalizer.py      # Name/company cleaning, Jinja2 templates
│   ├── sender.py            # SMTP connection, send, retry, throttling
│   └── reporter.py          # Results Excel + Rich stats table
├── templates/
│   ├── subjects.txt         # Subject line variants
│   ├── body_ta.html.j2      # HTML body – Talent Acquisition roles
│   ├── body_ta.txt.j2       # Plain text body – TA roles
│   ├── body_hr.html.j2      # HTML body – general HR roles
│   ├── body_hr.txt.j2       # Plain text body – HR roles
│   ├── followup.html.j2     # Follow-up email HTML
│   └── followup.txt.j2      # Follow-up email plain text
├── config.yaml              # Settings and column mapping
├── profile.yaml             # YOUR personal details
├── .env.example             # Template for SMTP credentials
├── .gitignore
├── requirements.txt
├── README.md
└── tests/
    ├── __init__.py
    └── test_mailer.py       # Unit tests
```

---

## 📊 Input File Format

Your Excel file should have these columns (case-insensitive, spaces/underscores ignored):

| Column | Required | Description |
|--------|----------|-------------|
| Name | ✅ | Contact's full name |
| Email | ✅ | Contact's email address |
| Designation | ✅ | Job title / role |
| Company | ✅ | Company name |
| Intern Interest Score (1-5) | ❌ | Used with `--min-score` and resume attachment |

Extra columns (Rank, Interest Level, etc.) are preserved in the output report.

---

## 📎 Resume Attachment Logic

- Resume is attached **only** for contacts with an Intern Interest Score of **3, 4, or 5**
- If no score column exists, resume is not attached (unless via `--resume` flag in test-send)
- Configure the resume path in `profile.yaml` → `resume_path`

---

## 🔧 Troubleshooting

### `535 Authentication failed`
- You're using your regular Gmail password instead of an App Password
- 2-Step Verification is not enabled
- The App Password has been revoked

### `421 / 450 Too many connections / rate limit`
- You're sending too fast. Increase `delay_range` in `config.yaml`
- Reduce `daily_cap` to 20-30 for warm-up
- Wait 24 hours before retrying

### Emails landing in Spam
- **Warm up**: Start with 20-30 emails/day for the first week
- Use a **professional email address** (not throwaway)
- Keep `daily_cap` low (50-100)
- Don't use ALL CAPS, exclamation marks, or spammy words in templates
- Make sure your email has a proper display name

### `Connection refused` or timeout
- Check your internet connection
- Your ISP may block port 587. Try port 465 with `use_ssl: true` in config.yaml
- Corporate/campus networks often block SMTP ports

### Google Workspace vs Personal Gmail
- **Personal Gmail**: 500 emails/day limit
- **Google Workspace**: up to 2,000 emails/day
- Both require App Passwords for this tool

---

## ⚡ Assumptions & Design Decisions

1. **Resume attachment**: Only for contacts with Intern Interest Score ≥ 3 (the "3-5 rated tuples" requirement)
2. **Subject rotation**: Deterministic (hash-based) so the same contact always gets the same subject across runs
3. **Template selection**: Based on keyword matching in the Designation field — keywords like "talent acquisition", "recruit", "campus", "hiring" trigger the TA template; everything else gets the HR template
4. **Daily cap**: Counts from `sent_log.csv` entries timestamped today, so the cap is respected across multiple runs in a day
5. **Graceful shutdown**: Ctrl+C finishes the current email, saves progress, and exits cleanly
6. **DNS timeout**: 5 seconds per domain, with LRU cache to avoid repeat lookups

---

## 📜 License

This project is for personal/educational use. Use responsibly and in compliance with applicable anti-spam laws.
