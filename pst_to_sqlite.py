"""
pst_to_sqlite.py  –  Extract all items from an Outlook PST file into SQLite.

Usage:
    python pst_to_sqlite.py <pst_path> <db_path>

Extracts: emails, contacts, calendar items, tasks, and logs unknown item types.
Requires: Python 3, pywin32, Outlook installed on Windows.
"""

import sys
import os
import sqlite3
import traceback
import datetime
from pathlib import Path

try:
    import win32com.client
    import pywintypes
except ImportError:
    print("ERROR: pywin32 not installed. Run: pip install pywin32")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Outlook item class constants
# ---------------------------------------------------------------------------
OL_MAIL_ITEM        = 43
OL_CONTACT_ITEM     = 40
OL_APPOINTMENT_ITEM = 26
OL_TASK_ITEM        = 48
OL_POST_ITEM        = 45
OL_NOTE_ITEM        = 44
OL_JOURNAL_ITEM     = 42
OL_DIST_LIST_ITEM   = 69
OL_MEETING_REQUEST  = 53

# AddStoreEx type flag for existing Unicode PST
OL_STORE_UNICODE = 2

# Recipient types
OL_TO  = 1
OL_CC  = 2
OL_BCC = 3
RECIPIENT_TYPES = {OL_TO: "To", OL_CC: "CC", OL_BCC: "BCC"}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def safe_get(obj, prop, default=None):
    """Get a COM property, returning default on any error."""
    try:
        val = getattr(obj, prop)
        return val
    except Exception:
        return default


def to_iso(dt):
    """Convert a pywintypes.datetime or Python datetime to ISO-8601 string (UTC)."""
    if dt is None:
        return None
    try:
        if hasattr(dt, 'utctimetuple'):
            # pywintypes datetime – treat as local, convert to UTC string
            import time
            ts = time.mktime(dt.timetuple())
            utc = datetime.datetime.utcfromtimestamp(ts)
            return utc.strftime("%Y-%m-%dT%H:%M:%SZ")
        return str(dt)
    except Exception:
        return None


def folder_path(folder):
    """Build a slash-delimited folder path string."""
    parts = []
    try:
        f = folder
        while f:
            parts.append(f.Name)
            try:
                f = f.Parent
                if not hasattr(f, 'Name'):
                    break
            except Exception:
                break
    except Exception:
        pass
    return "/" + "/".join(reversed(parts))


# ---------------------------------------------------------------------------
# Database setup
# ---------------------------------------------------------------------------

SCHEMA = """
CREATE TABLE IF NOT EXISTS emails (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_id          TEXT,
    store_id          TEXT,
    pst_source        TEXT,
    folder_path       TEXT,
    message_class     TEXT,
    subject           TEXT,
    body_text         TEXT,
    body_html         TEXT,
    sender_name       TEXT,
    sender_email_raw  TEXT,
    sender_email_smtp TEXT,
    date_sent         TEXT,
    date_received     TEXT,
    has_attachments   INTEGER,
    importance        INTEGER,
    categories        TEXT,
    last_modified     TEXT
);

CREATE TABLE IF NOT EXISTS email_recipients (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    email_id    INTEGER REFERENCES emails(id),
    name        TEXT,
    email_raw   TEXT,
    email_smtp  TEXT,
    type        TEXT
);

CREATE TABLE IF NOT EXISTS email_attachments (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    email_id        INTEGER REFERENCES emails(id),
    filename        TEXT,
    size_bytes      INTEGER,
    attachment_type INTEGER
);

CREATE TABLE IF NOT EXISTS contacts (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_id           TEXT,
    store_id           TEXT,
    pst_source         TEXT,
    folder_path        TEXT,
    full_name          TEXT,
    email1             TEXT,
    email1_display     TEXT,
    email2             TEXT,
    email2_display     TEXT,
    email3             TEXT,
    email3_display     TEXT,
    phone_business     TEXT,
    phone_mobile       TEXT,
    phone_home         TEXT,
    company            TEXT,
    title              TEXT,
    department         TEXT,
    address_business   TEXT,
    address_home       TEXT,
    birthday           TEXT,
    notes              TEXT,
    categories         TEXT,
    last_modified      TEXT
);

CREATE TABLE IF NOT EXISTS calendar_items (
    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_id             TEXT,
    store_id             TEXT,
    pst_source           TEXT,
    folder_path          TEXT,
    message_class        TEXT,
    subject              TEXT,
    body_text            TEXT,
    body_html            TEXT,
    start_dt             TEXT,
    end_dt               TEXT,
    location             TEXT,
    organizer            TEXT,
    required_attendees   TEXT,
    optional_attendees   TEXT,
    is_all_day           INTEGER,
    is_recurring         INTEGER,
    recurrence_pattern   TEXT,
    categories           TEXT,
    last_modified        TEXT
);

CREATE TABLE IF NOT EXISTS tasks (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_id         TEXT,
    store_id         TEXT,
    pst_source       TEXT,
    folder_path      TEXT,
    subject          TEXT,
    body_text        TEXT,
    start_date       TEXT,
    due_date         TEXT,
    completion_date  TEXT,
    status           INTEGER,
    priority         INTEGER,
    percent_complete REAL,
    categories       TEXT,
    last_modified    TEXT
);

CREATE TABLE IF NOT EXISTS unknown_items (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_id      TEXT,
    store_id      TEXT,
    pst_source    TEXT,
    folder_path   TEXT,
    message_class TEXT,
    item_class    INTEGER,
    subject       TEXT,
    last_modified TEXT
);

CREATE TABLE IF NOT EXISTS extraction_log (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp  TEXT,
    level      TEXT,
    message    TEXT
);
"""


def init_db(db_path):
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.executescript(SCHEMA)
    conn.commit()
    return conn


def log_db(conn, level, msg):
    conn.execute(
        "INSERT INTO extraction_log (timestamp, level, message) VALUES (?,?,?)",
        (datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"), level, msg)
    )


# ---------------------------------------------------------------------------
# Item extractors
# ---------------------------------------------------------------------------

def extract_email(item, folder_fpath, pst_source, store_id, conn):
    entry_id  = safe_get(item, 'EntryID')
    subject   = safe_get(item, 'Subject', '')
    body_text = safe_get(item, 'Body', '')
    body_html = safe_get(item, 'HTMLBody', '')
    sndr_name = safe_get(item, 'SenderName', '')
    sndr_raw  = safe_get(item, 'SenderEmailAddress', '')
    sndr_smtp = safe_get(item, 'SenderEmailAddress', '')
    # Try to get SMTP address via sender object when it's an EX address
    try:
        sender_obj = item.Sender
        if sender_obj and sndr_raw and sndr_raw.upper().startswith('/O='):
            sndr_smtp = safe_get(sender_obj, 'GetExchangeUser', None)
            if sndr_smtp:
                sndr_smtp = safe_get(sndr_smtp, 'PrimarySmtpAddress', sndr_raw)
            else:
                sndr_smtp = sndr_raw
    except Exception:
        pass

    date_sent     = to_iso(safe_get(item, 'SentOn'))
    date_received = to_iso(safe_get(item, 'ReceivedTime'))
    has_att       = int(bool(safe_get(item, 'Attachments') and safe_get(item.Attachments, 'Count', 0) > 0))
    importance    = safe_get(item, 'Importance', 1)
    categories    = safe_get(item, 'Categories', '')
    msg_class     = safe_get(item, 'MessageClass', '')
    last_mod      = to_iso(safe_get(item, 'LastModificationTime'))

    cur = conn.execute(
        """INSERT INTO emails
           (entry_id,store_id,pst_source,folder_path,message_class,subject,
            body_text,body_html,sender_name,sender_email_raw,sender_email_smtp,
            date_sent,date_received,has_attachments,importance,categories,last_modified)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (entry_id, store_id, pst_source, folder_fpath, msg_class, subject,
         body_text, body_html, sndr_name, sndr_raw, sndr_smtp,
         date_sent, date_received, has_att, importance, categories, last_mod)
    )
    email_id = cur.lastrowid

    # Recipients
    try:
        recipients = item.Recipients
        for i in range(1, recipients.Count + 1):
            try:
                r = recipients.Item(i)
                r_name  = safe_get(r, 'Name', '')
                r_raw   = safe_get(r, 'Address', '')
                r_smtp  = r_raw
                try:
                    eu = r.AddressEntry.GetExchangeUser()
                    if eu:
                        r_smtp = safe_get(eu, 'PrimarySmtpAddress', r_raw)
                except Exception:
                    pass
                r_type  = RECIPIENT_TYPES.get(safe_get(r, 'Type', OL_TO), 'To')
                conn.execute(
                    "INSERT INTO email_recipients (email_id,name,email_raw,email_smtp,type) VALUES (?,?,?,?,?)",
                    (email_id, r_name, r_raw, r_smtp, r_type)
                )
            except Exception as e:
                log_db(conn, 'WARN', f"Recipient error on email {entry_id}: {e}")
    except Exception as e:
        log_db(conn, 'WARN', f"Recipients collection error on email {entry_id}: {e}")

    # Attachments
    try:
        attachments = item.Attachments
        for i in range(1, attachments.Count + 1):
            try:
                a = attachments.Item(i)
                a_name = safe_get(a, 'FileName', '') or safe_get(a, 'DisplayName', '')
                a_size = safe_get(a, 'Size', 0)
                a_type = safe_get(a, 'Type', 0)
                conn.execute(
                    "INSERT INTO email_attachments (email_id,filename,size_bytes,attachment_type) VALUES (?,?,?,?)",
                    (email_id, a_name, a_size, a_type)
                )
            except Exception as e:
                log_db(conn, 'WARN', f"Attachment error on email {entry_id}: {e}")
    except Exception as e:
        log_db(conn, 'WARN', f"Attachments collection error on email {entry_id}: {e}")


def extract_contact(item, folder_fpath, pst_source, store_id, conn):
    entry_id = safe_get(item, 'EntryID')
    conn.execute(
        """INSERT INTO contacts
           (entry_id,store_id,pst_source,folder_path,full_name,
            email1,email1_display,email2,email2_display,email3,email3_display,
            phone_business,phone_mobile,phone_home,company,title,department,
            address_business,address_home,birthday,notes,categories,last_modified)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (entry_id, store_id, pst_source, folder_fpath,
         safe_get(item, 'FullName', ''),
         safe_get(item, 'Email1Address', ''),
         safe_get(item, 'Email1DisplayName', ''),
         safe_get(item, 'Email2Address', ''),
         safe_get(item, 'Email2DisplayName', ''),
         safe_get(item, 'Email3Address', ''),
         safe_get(item, 'Email3DisplayName', ''),
         safe_get(item, 'BusinessTelephoneNumber', ''),
         safe_get(item, 'MobileTelephoneNumber', ''),
         safe_get(item, 'HomeTelephoneNumber', ''),
         safe_get(item, 'CompanyName', ''),
         safe_get(item, 'JobTitle', ''),
         safe_get(item, 'Department', ''),
         safe_get(item, 'BusinessAddress', ''),
         safe_get(item, 'HomeAddress', ''),
         to_iso(safe_get(item, 'Birthday')),
         safe_get(item, 'Body', ''),
         safe_get(item, 'Categories', ''),
         to_iso(safe_get(item, 'LastModificationTime')))
    )


def extract_calendar(item, folder_fpath, pst_source, store_id, conn):
    entry_id   = safe_get(item, 'EntryID')
    is_recurr  = int(bool(safe_get(item, 'IsRecurring', False)))
    recur_pat  = None
    if is_recurr:
        try:
            rp = item.GetRecurrencePattern()
            recur_pat = safe_get(rp, 'RecurrenceType', None)
        except Exception:
            pass
    conn.execute(
        """INSERT INTO calendar_items
           (entry_id,store_id,pst_source,folder_path,message_class,subject,
            body_text,body_html,start_dt,end_dt,location,organizer,
            required_attendees,optional_attendees,is_all_day,is_recurring,
            recurrence_pattern,categories,last_modified)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (entry_id, store_id, pst_source, folder_fpath,
         safe_get(item, 'MessageClass', ''),
         safe_get(item, 'Subject', ''),
         safe_get(item, 'Body', ''),
         safe_get(item, 'HTMLBody', ''),
         to_iso(safe_get(item, 'Start')),
         to_iso(safe_get(item, 'End')),
         safe_get(item, 'Location', ''),
         safe_get(item, 'Organizer', ''),
         safe_get(item, 'RequiredAttendees', ''),
         safe_get(item, 'OptionalAttendees', ''),
         int(bool(safe_get(item, 'AllDayEvent', False))),
         is_recurr,
         str(recur_pat) if recur_pat is not None else None,
         safe_get(item, 'Categories', ''),
         to_iso(safe_get(item, 'LastModificationTime')))
    )


def extract_task(item, folder_fpath, pst_source, store_id, conn):
    entry_id = safe_get(item, 'EntryID')
    conn.execute(
        """INSERT INTO tasks
           (entry_id,store_id,pst_source,folder_path,subject,body_text,
            start_date,due_date,completion_date,status,priority,
            percent_complete,categories,last_modified)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (entry_id, store_id, pst_source, folder_fpath,
         safe_get(item, 'Subject', ''),
         safe_get(item, 'Body', ''),
         to_iso(safe_get(item, 'StartDate')),
         to_iso(safe_get(item, 'DueDate')),
         to_iso(safe_get(item, 'DateCompleted')),
         safe_get(item, 'Status', 0),
         safe_get(item, 'Importance', 1),
         safe_get(item, 'PercentComplete', 0.0),
         safe_get(item, 'Categories', ''),
         to_iso(safe_get(item, 'LastModificationTime')))
    )


def extract_unknown(item, item_class, folder_fpath, pst_source, store_id, conn):
    conn.execute(
        """INSERT INTO unknown_items
           (entry_id,store_id,pst_source,folder_path,message_class,item_class,subject,last_modified)
           VALUES (?,?,?,?,?,?,?,?)""",
        (safe_get(item, 'EntryID'),
         store_id, pst_source, folder_fpath,
         safe_get(item, 'MessageClass', ''),
         item_class,
         safe_get(item, 'Subject', ''),
         to_iso(safe_get(item, 'LastModificationTime')))
    )


# ---------------------------------------------------------------------------
# Folder walker
# ---------------------------------------------------------------------------

def walk_folder(folder, pst_source, store_id, conn, counts):
    fpath = folder_path(folder)
    items_collection = None

    try:
        items_collection = folder.Items
        n = items_collection.Count
    except Exception as e:
        log_db(conn, 'WARN', f"Cannot access items in folder '{fpath}': {e}")
        n = 0

    for i in range(1, n + 1):
        try:
            item = items_collection.Item(i)
            item_class = safe_get(item, 'Class', -1)

            if item_class == OL_MAIL_ITEM or item_class == OL_MEETING_REQUEST:
                extract_email(item, fpath, pst_source, store_id, conn)
                counts['emails'] += 1
            elif item_class == OL_CONTACT_ITEM:
                extract_contact(item, fpath, pst_source, store_id, conn)
                counts['contacts'] += 1
            elif item_class == OL_APPOINTMENT_ITEM:
                extract_calendar(item, fpath, pst_source, store_id, conn)
                counts['calendar'] += 1
            elif item_class == OL_TASK_ITEM:
                extract_task(item, fpath, pst_source, store_id, conn)
                counts['tasks'] += 1
            else:
                extract_unknown(item, item_class, fpath, pst_source, store_id, conn)
                counts['unknown'] += 1

            # Commit in batches of 500
            if sum(counts.values()) % 500 == 0:
                conn.commit()
                total = sum(counts.values())
                print(f"  ... {total} items processed "
                      f"(emails={counts['emails']}, contacts={counts['contacts']}, "
                      f"calendar={counts['calendar']}, tasks={counts['tasks']}, "
                      f"other={counts['unknown']})")

        except Exception as e:
            msg = f"Error processing item {i} in '{fpath}': {e}"
            log_db(conn, 'ERROR', msg)
            print(f"  WARN: {msg}")

    # Recurse into subfolders
    try:
        subfolders = folder.Folders
        for j in range(1, subfolders.Count + 1):
            try:
                walk_folder(subfolders.Item(j), pst_source, store_id, conn, counts)
            except Exception as e:
                log_db(conn, 'WARN', f"Error accessing subfolder {j} of '{fpath}': {e}")
    except Exception as e:
        log_db(conn, 'WARN', f"Cannot enumerate subfolders of '{fpath}': {e}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) != 3:
        print("Usage: python pst_to_sqlite.py <pst_path> <db_path>")
        sys.exit(1)

    pst_path = os.path.abspath(sys.argv[1])
    db_path  = os.path.abspath(sys.argv[2])

    if not os.path.exists(pst_path):
        print(f"ERROR: PST file not found: {pst_path}")
        sys.exit(1)

    pst_source = os.path.basename(pst_path)
    print(f"PST source : {pst_path}")
    print(f"Database   : {db_path}")

    # Connect to Outlook
    print("\nConnecting to Outlook...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        session = outlook.Session
    except Exception as e:
        print(f"ERROR: Cannot connect to Outlook: {e}")
        sys.exit(1)

    # Check if PST is already mounted
    opened_by_script = False
    target_store = None
    pst_lower = pst_path.lower()

    for store in session.Stores:
        try:
            if safe_get(store, 'FilePath', '').lower() == pst_lower:
                target_store = store
                print(f"PST already mounted in Outlook (will not remove on exit).")
                break
        except Exception:
            pass

    if target_store is None:
        print("Opening PST file in Outlook...")
        try:
            session.AddStoreEx(pst_path, OL_STORE_UNICODE)
            opened_by_script = True
        except Exception as e:
            # May already be open or different error
            log_msg = f"AddStoreEx warning: {e}"
            print(f"  Note: {log_msg}")

        # Find the store we just added
        for store in session.Stores:
            try:
                if safe_get(store, 'FilePath', '').lower() == pst_lower:
                    target_store = store
                    break
            except Exception:
                pass

    if target_store is None:
        print("ERROR: Could not locate the PST store in Outlook after opening.")
        sys.exit(1)

    store_id = safe_get(target_store, 'StoreID', '')
    print(f"Store ID   : {store_id[:40]}..." if len(store_id) > 40 else f"Store ID   : {store_id}")

    # Init database
    print(f"\nInitializing database...")
    conn = init_db(db_path)
    log_db(conn, 'INFO', f"Extraction started: pst={pst_path}, db={db_path}")
    conn.commit()

    # Walk the PST
    print("Walking PST folders...\n")
    counts = {'emails': 0, 'contacts': 0, 'calendar': 0, 'tasks': 0, 'unknown': 0}
    start_time = datetime.datetime.utcnow()

    try:
        root = target_store.GetRootFolder()
        walk_folder(root, pst_source, store_id, conn, counts)
    except Exception as e:
        print(f"ERROR during walk: {e}")
        traceback.print_exc()
        log_db(conn, 'ERROR', f"Walk failed: {e}")

    conn.commit()

    elapsed = (datetime.datetime.utcnow() - start_time).total_seconds()
    total = sum(counts.values())
    summary = (f"Done in {elapsed:.1f}s. Total={total} "
               f"(emails={counts['emails']}, contacts={counts['contacts']}, "
               f"calendar={counts['calendar']}, tasks={counts['tasks']}, "
               f"other={counts['unknown']})")
    print(f"\n{summary}")
    log_db(conn, 'INFO', summary)
    conn.commit()
    conn.close()

    # Remove store only if we opened it
    if opened_by_script:
        print("Removing PST from Outlook (cleaning up)...")
        try:
            session.RemoveStore(target_store.GetRootFolder())
        except Exception as e:
            print(f"  Note: Could not remove store (may need manual removal): {e}")

    print(f"\nDatabase saved to: {db_path}")


if __name__ == "__main__":
    main()
