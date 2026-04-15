"""
fts.py  –  Build FTS5 full-text search indexes and B-tree indexes.

Indexes are written into the source DB once on first run.
Subsequent runs detect the existing tables and return immediately.
"""

import sqlite3
import logging

logger = logging.getLogger(__name__)

_indexed: set[str] = set()


def ensure_fts_indexes(db_path: str) -> None:
    """Create FTS5 virtual tables and B-tree indexes if they do not yet exist."""
    if db_path in _indexed:
        return

    conn = sqlite3.connect(db_path)
    try:
        exists = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='emails_fts'"
        ).fetchone()

        if not exists:
            logger.info("Building FTS5 and B-tree indexes (first run)…")
            _build(conn)
            logger.info("Indexes built.")
        else:
            logger.debug("FTS indexes already present.")
    finally:
        conn.close()

    _indexed.add(db_path)


def _build(conn: sqlite3.Connection) -> None:
    conn.execute("""
        CREATE VIRTUAL TABLE emails_fts USING fts5(
            subject,
            body_text,
            sender_name,
            sender_email,
            folder_path,
            categories,
            tokenize = 'unicode61 remove_diacritics 1'
        )
    """)
    conn.execute("""
        INSERT INTO emails_fts(rowid, subject, body_text, sender_name, sender_email, folder_path, categories)
        SELECT
            id,
            coalesce(subject,  ''),
            coalesce(body_text, ''),
            coalesce(sender_name, ''),
            coalesce(nullif(sender_email_smtp, ''), sender_email_raw, sender_name, ''),
            coalesce(folder_path, ''),
            coalesce(categories, '')
        FROM emails
    """)

    conn.execute("""
        CREATE VIRTUAL TABLE contacts_fts USING fts5(
            full_name,
            company,
            title,
            department,
            email1,
            email2,
            email3,
            notes,
            tokenize = 'unicode61 remove_diacritics 1'
        )
    """)
    conn.execute("""
        INSERT INTO contacts_fts(rowid, full_name, company, title, department, email1, email2, email3, notes)
        SELECT
            id,
            coalesce(full_name,   ''),
            coalesce(company,     ''),
            coalesce(title,       ''),
            coalesce(department,  ''),
            coalesce(email1,      ''),
            coalesce(email2,      ''),
            coalesce(email3,      ''),
            coalesce(notes,       '')
        FROM contacts
    """)

    conn.execute("""
        CREATE VIRTUAL TABLE calendar_fts USING fts5(
            subject,
            body_text,
            location,
            organizer,
            required_attendees,
            optional_attendees,
            tokenize = 'unicode61 remove_diacritics 1'
        )
    """)
    conn.execute("""
        INSERT INTO calendar_fts(
            rowid, subject, body_text, location, organizer,
            required_attendees, optional_attendees)
        SELECT
            id,
            coalesce(subject,             ''),
            coalesce(body_text,           ''),
            coalesce(location,            ''),
            coalesce(organizer,           ''),
            coalesce(required_attendees,  ''),
            coalesce(optional_attendees,  '')
        FROM calendar_items
    """)

    # B-tree indexes for filter queries
    conn.execute("CREATE INDEX IF NOT EXISTS idx_emails_date_recv   ON emails(date_received)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_emails_date_sent   ON emails(date_sent)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_emails_folder      ON emails(folder_path)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_emails_sender      ON emails(sender_email_smtp)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_recip_email_id     ON email_recipients(email_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_attach_email_id    ON email_attachments(email_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_cal_start          ON calendar_items(start_dt)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_contacts_name      ON contacts(full_name)")

    conn.commit()
