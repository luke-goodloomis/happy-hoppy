# Happy Hoppy 🐰

A local web application for searching and exploring email data extracted from Outlook PST archive files.
Built with Python (Flask) and SQLite, it provides full-text search, filtering, and browsing across
emails, contacts, and calendar events — all running locally with no cloud dependencies.

---

## Purpose

Outlook PST files are opaque binary archives that are difficult to search or analyze outside of
Outlook itself. Happy Hoppy solves this by:

1. **Extracting** all items from one or more PST files into a structured SQLite database
   (`pst_to_sqlite.py`)
2. **Serving** a web UI for fast full-text search, date/sender/folder filtering, and detailed
   viewing of emails, contacts, and calendar events (`app.py`)

Primary use case: finding information about subjects, topics, keywords, products, manufacturers,
people, and projects buried across large email archives.

---

## Project Structure

```
happy-hoppy/
│
├── pst_to_sqlite.py        # Step 1 – PST extraction script (run once per archive)
├── app.py                  # Step 2 – Flask web server (run to use the UI)
├── fts.py                  # FTS5 full-text search index builder (called by app.py on first run)
│
├── templates/              # Jinja2 HTML templates (Bootstrap 5)
│   ├── base.html           #   Navigation shell, CDN links
│   ├── index.html          #   Home dashboard (stats, recent emails, top senders)
│   ├── search.html         #   Search results with Emails / Contacts / Calendar tabs
│   ├── email.html          #   Single email detail view (HTML body, recipients, attachments)
│   ├── contacts.html       #   Contacts list with search
│   ├── contact.html        #   Single contact detail view
│   └── calendar.html       #   Calendar events list with search and date filter
│
├── static/
│   ├── css/main.css        # Custom styles (navy navbar, stat cards, result highlighting)
│   └── js/main.js          # UI helpers (iframe resize, keyboard shortcut, title update)
│
├── requirements.txt        # Python dependencies (flask, pywin32)
├── .gitignore              # Excludes *.db, *.pst, *.ost — never commit private data
└── README.md               # This file
```

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Windows 10/11 | COM automation requires Windows |
| Microsoft Outlook (desktop) | Installed and activated; used to open PST files |
| Python 3.10+ | Tested on 3.11 |
| pip packages | `flask`, `pywin32` (see `requirements.txt`) |

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Extract a PST file into a SQLite database

```bash
python pst_to_sqlite.py "C:\path\to\archive.pst" "C:\path\to\output.db"
```

**What this does:**
- Opens the PST via Outlook COM automation (Outlook must be installed but does **not** need to be
  running — the script starts it in the background)
- Walks every folder in the archive recursively
- Extracts emails, contacts, calendar items, tasks, and unknown item types
- Writes everything to a SQLite database
- Removes the PST from Outlook's store list when done (cleanup)

**Estimated time:**
| PST size | Approximate time |
|---|---|
| 28 MB (test) | ~60 seconds |
| 500 MB | ~15–20 minutes |
| 4 GB | ~2–4 hours |

**Progress** is printed to the console every 500 items.

### 3. Start the web UI

```bash
# Default: uses the test database path hardcoded as default arg
python app.py

# Or specify a database explicitly
python app.py --db "C:\path\to\output.db"

# LAN access (allow connections from other machines)
python app.py --db "C:\path\to\output.db" --host 0.0.0.0 --port 5000
```

Open **http://127.0.0.1:5000** in any browser.

> **First run:** FTS5 full-text search indexes are built automatically on first launch.
> For the 28 MB test database this takes under 2 seconds. For multi-GB databases allow 30–60 seconds.
> Subsequent runs reuse the existing indexes instantly.

---

## Database Schema

All data lives in a single SQLite file created by `pst_to_sqlite.py`.

| Table | Description |
|---|---|
| `emails` | All mail items (subject, body_text, body_html, sender, dates, folder, importance) |
| `email_recipients` | To / CC / BCC recipients for each email (FK → emails.id) |
| `email_attachments` | Attachment metadata — filename and size only, not file contents |
| `contacts` | Contact items (name, company, title, emails, phones, addresses, notes) |
| `calendar_items` | Appointment/meeting items (subject, start/end, location, attendees, recurrence) |
| `tasks` | Task items (subject, due date, status, priority, percent complete) |
| `unknown_items` | Any item type not explicitly handled (logged by message class) |
| `extraction_log` | Timestamped INFO/WARN/ERROR messages from the extraction run |
| `emails_fts` | FTS5 virtual table — built by fts.py on first app launch |
| `contacts_fts` | FTS5 virtual table — built by fts.py on first app launch |
| `calendar_fts` | FTS5 virtual table — built by fts.py on first app launch |

**Key columns in `emails`:**

| Column | Notes |
|---|---|
| `body_text` | Plain-text body |
| `body_html` | HTML body (sanitized before display) |
| `sender_email_smtp` | Best-effort SMTP address (may be blank for Exchange-internal senders) |
| `sender_email_raw` | Raw Outlook address (may be X.500 / Exchange DN for internal senders) |
| `date_sent` / `date_received` | ISO-8601 UTC strings (`2024-03-15T09:30:00Z`) |
| `folder_path` | Slash-delimited path from PST root (e.g. `/archive.pst/Inbox/Projects`) |

---

## Key Technical Decisions

### PST extraction via Outlook COM
`pst_to_sqlite.py` uses `win32com.client` (pywin32) to drive Outlook as a COM server.
This is the most reliable approach on Windows because it handles Unicode PSTs, Exchange-format
addresses, and recurrence data without needing third-party PST parsing libraries.

**Caveats:**
- Outlook must be installed (but does not need to be open)
- Password-protected PSTs will fail with a COM error
- Some items may have unreadable properties — these are caught per-property and logged

### FTS5 indexes are written into the source DB
`fts.py` adds `emails_fts`, `contacts_fts`, and `calendar_fts` virtual tables to the same SQLite
file as the extracted data. This is a deliberate simplicity trade-off:

- **Pro:** Single file to manage; no ATTACH complexity
- **Con:** Slightly mutates the "source" DB

If you need an immutable source DB (e.g., for forensic use), run the extraction to a temp copy
and point the app at that copy, or refactor `fts.py` to write indexes to a separate `_cache.db`.

### HTML email sanitization
`sanitize_email_html()` in `app.py` strips `<script>` tags, `on*` event handlers, and replaces
external `src=` on `<img>` tags with `data-blocked-src`. The sanitized HTML is rendered in a
sandboxed `<iframe srcdoc="...">` with `sandbox="allow-same-origin"` (no script execution,
no navigation, no form submission).

Users can click "Show anyway" to restore blocked images for a specific email.

### Sender email normalization
Outlook stores internal senders as Exchange/X.500 DN strings (e.g. `/O=ExchangeLabs/CN=...`).
The app uses a SQL `coalesce` expression throughout to prefer `sender_email_smtp` over the raw
value, falling back to `sender_name` for display. The extraction script attempts to resolve
Exchange users to SMTP via `AddressEntry.GetExchangeUser().PrimarySmtpAddress`.

---

## Web UI Routes

| Route | Description |
|---|---|
| `GET /` | Home dashboard: stats, recent emails, top senders, top folders |
| `GET /search?q=&type=&sender=&folder=&date_from=&date_to=&page=` | Search results (tabs: emails / contacts / calendar) |
| `GET /email/<id>` | Full email detail: HTML body, recipients, attachments, related searches |
| `GET /contacts?q=&page=` | Contacts list with FTS search |
| `GET /contact/<id>` | Single contact detail |
| `GET /calendar?q=&date_from=&date_to=&page=` | Calendar events with search and date filter |
| `GET /api/stats` | JSON: total counts, top senders, email volume by month |

**Search tips (FTS5 syntax):**
- `crown hearing` → any email containing both words (anywhere in subject or body)
- `"crown hearing"` → exact phrase
- `Cisco OR Apple` → either term
- `Bosch NOT email` → excludes matches containing "email"
- Searches are case-insensitive and diacritic-insensitive

---

## Adding a New PST to an Existing Database

`pst_to_sqlite.py` **appends** to an existing database — it does not wipe existing data.
The `pst_source` column in every table records which PST file each item came from.

After appending new data, the FTS indexes need to be rebuilt. Delete the FTS tables so they
are recreated on next launch:

```sql
-- Run this in any SQLite client (e.g. DB Browser for SQLite)
DROP TABLE IF EXISTS emails_fts;
DROP TABLE IF EXISTS contacts_fts;
DROP TABLE IF EXISTS calendar_fts;
```

Then restart `app.py` — it will rebuild the indexes automatically.

---

## Common Tasks for Future Development

### Add a new item type to extraction
1. Add a handler function `extract_xyz(item, ...)` in `pst_to_sqlite.py` following the pattern
   of `extract_email` / `extract_contact`
2. Add a `CREATE TABLE` statement to the `SCHEMA` constant
3. Add the item class constant (e.g. `OL_NOTE_ITEM = 44`) and dispatch in `walk_folder()`
4. Add a corresponding FTS virtual table in `fts.py` if needed

### Add a new search filter
1. Add the filter parameter to the relevant route in `app.py`
2. Add a `WHERE` clause in the appropriate `_search_*` function
3. Add the filter input to the corresponding template

### Export search results to CSV
Add a `/search/export` route that runs the same query without `LIMIT`/`OFFSET` and returns
a `text/csv` response using Python's `csv` module.

### Support multiple databases
The app currently reads `app.config["DB_PATH"]` which is set at startup.
To support switching databases at runtime, store the path in the session or accept it as a
query parameter, and update `get_db()` accordingly.

### Process large PST files (multi-GB)
- Run `pst_to_sqlite.py` overnight; output is written incrementally (committed every 500 items)
- If extraction is interrupted it can be re-run — items will be duplicated; deduplicate using
  `entry_id` if needed
- After extraction, rebuild FTS indexes as described above

---

## Security Notes

- The app binds to `127.0.0.1` by default — only accessible from localhost
- Use `--host 0.0.0.0` only on a trusted private network
- **Never commit `.db` or `.pst` files to git** — they contain private email data
- The `.gitignore` already excludes these file types

---

## Dependencies

| Package | Version | Purpose |
|---|---|---|
| `flask` | ≥3.0 | Web framework |
| `pywin32` | ≥311 | Outlook COM automation (extraction only) |

No database ORM, no JavaScript build step, no external search engine.
All search is handled by SQLite's built-in FTS5 extension.

---

## Tested Configuration

- Windows 11, Microsoft 365 Outlook 16.x
- Python 3.11.9
- Flask 3.1.x
- SQLite 3.45+ (bundled with Python 3.11)
- Tested against PST files ranging from 28 MB to 4.6 GB
