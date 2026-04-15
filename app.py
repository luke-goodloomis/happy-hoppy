"""
app.py  –  Happy Hoppy Flask web UI for Outlook PST email search.

Usage:
    python app.py --db "path/to/outlook_data.db"
"""

import argparse
import re
import sqlite3
from pathlib import Path

from flask import Flask, abort, g, jsonify, redirect, render_template, request, url_for

from fts import ensure_fts_indexes

app = Flask(__name__)
PAGE_SIZE = 25

# ── DB helpers ────────────────────────────────────────────────────────────────

def get_db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = sqlite3.connect(app.config["DB_PATH"])
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA journal_mode=WAL")
    return g.db


@app.teardown_appcontext
def close_db(exc=None):
    db = g.pop("db", None)
    if db:
        db.close()


# ── Search helpers ────────────────────────────────────────────────────────────

SENDER_EXPR = "coalesce(nullif(sender_email_smtp,''), sender_email_raw, sender_name, '')"


def prepare_fts_query(raw: str) -> str | None:
    q = raw.strip()
    if not q:
        return None
    # Balance unmatched double-quotes so FTS5 doesn't error
    if q.count('"') % 2 != 0:
        q = q.replace('"', "")
    return q


def paginate(page: int, total: int) -> dict:
    total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
    return {
        "page": page,
        "total": total,
        "total_pages": total_pages,
        "has_prev": page > 1,
        "has_next": page < total_pages,
        "prev": page - 1,
        "next": page + 1,
        "showing_from": (page - 1) * PAGE_SIZE + 1,
        "showing_to": min(page * PAGE_SIZE, total),
    }


def sanitize_email_html(body: str) -> str:
    """Strip scripts, event handlers, and external image sources from HTML."""
    if not body:
        return ""
    body = re.sub(r"<script[\s\S]*?</script>", "", body, flags=re.IGNORECASE)
    body = re.sub(r"\s+on\w+\s*=\s*(?:\"[^\"]*\"|'[^']*')", "", body, flags=re.IGNORECASE)
    # Block remote images (replace src with data-blocked-src)
    body = re.sub(
        r"(<img[^>]+\s)src(\s*=\s*[\"']https?://)",
        r"\1data-blocked-src\2",
        body,
        flags=re.IGNORECASE,
    )
    body = re.sub(r"<link\b[^>]*>", "", body, flags=re.IGNORECASE)
    return body


# ── Email search ──────────────────────────────────────────────────────────────

def _search_emails(db, fts_q, sender, folder, date_from, date_to, page, offset):
    where: list[str] = []
    params: list = []

    if fts_q:
        select = f"""
            SELECT e.id, e.subject, e.sender_name,
                   {SENDER_EXPR.replace('sender_', 'e.sender_')} AS sender_email,
                   e.date_received, e.has_attachments, e.folder_path, e.categories,
                   snippet(emails_fts, 1, '<mark>', '</mark>', '…', 40) AS snippet
            FROM emails_fts
            JOIN emails e ON emails_fts.rowid = e.id
            WHERE emails_fts MATCH ?"""
        params.append(fts_q)
        alias = "e"
    else:
        select = f"""
            SELECT e.id, e.subject, e.sender_name,
                   {SENDER_EXPR.replace('sender_', 'e.sender_')} AS sender_email,
                   e.date_received, e.has_attachments, e.folder_path, e.categories,
                   substr(e.body_text, 1, 200) AS snippet
            FROM emails e
            WHERE 1=1"""
        alias = "e"

    if sender:
        where.append(f"lower({SENDER_EXPR.replace('sender_', alias+'.sender_')}) LIKE lower(?)")
        params.append(f"%{sender}%")
    if folder:
        where.append(f"{alias}.folder_path LIKE ?")
        params.append(f"%{folder}%")
    if date_from:
        where.append(f"{alias}.date_received >= ?")
        params.append(date_from + "T00:00:00Z")
    if date_to:
        where.append(f"{alias}.date_received <= ?")
        params.append(date_to + "T23:59:59Z")

    where_sql = (" AND " + " AND ".join(where)) if where else ""
    order = "ORDER BY rank" if fts_q else "ORDER BY e.date_received DESC"

    try:
        total = db.execute(f"SELECT COUNT(*) FROM ({select}{where_sql})", params).fetchone()[0]
        results = db.execute(
            f"{select}{where_sql} {order} LIMIT {PAGE_SIZE} OFFSET {offset}", params
        ).fetchall()
    except Exception:
        # Fall back to plain LIKE search if FTS5 query is malformed
        results, total = _email_like_fallback(db, fts_q, page, offset)

    return results, total


def _email_like_fallback(db, q, page, offset):
    pattern = f"%{q}%"
    total = db.execute(
        "SELECT COUNT(*) FROM emails WHERE subject LIKE ? OR body_text LIKE ? OR sender_name LIKE ?",
        (pattern, pattern, pattern),
    ).fetchone()[0]
    results = db.execute(
        f"""SELECT id, subject, sender_name,
               {SENDER_EXPR} AS sender_email,
               date_received, has_attachments, folder_path, categories,
               substr(body_text,1,200) AS snippet
            FROM emails
            WHERE subject LIKE ? OR body_text LIKE ? OR sender_name LIKE ?
            ORDER BY date_received DESC
            LIMIT {PAGE_SIZE} OFFSET {offset}""",
        (pattern, pattern, pattern),
    ).fetchall()
    return results, total


def _search_contacts(db, fts_q, page, offset):
    if fts_q:
        try:
            total = db.execute(
                "SELECT COUNT(*) FROM contacts_fts WHERE contacts_fts MATCH ?", (fts_q,)
            ).fetchone()[0]
            results = db.execute(
                f"""SELECT c.id, c.full_name, c.company, c.title, c.email1,
                           c.phone_business, c.phone_mobile
                    FROM contacts_fts
                    JOIN contacts c ON contacts_fts.rowid = c.id
                    WHERE contacts_fts MATCH ?
                    ORDER BY rank
                    LIMIT {PAGE_SIZE} OFFSET {offset}""",
                (fts_q,),
            ).fetchall()
        except Exception:
            total, results = 0, []
    else:
        total = db.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]
        results = db.execute(
            f"""SELECT id, full_name, company, title, email1, phone_business, phone_mobile
                FROM contacts
                ORDER BY full_name
                LIMIT {PAGE_SIZE} OFFSET {offset}"""
        ).fetchall()
    return results, total


def _search_calendar(db, fts_q, date_from, date_to, page, offset):
    where: list[str] = []
    params: list = []

    if fts_q:
        select = """
            SELECT c.id, c.subject, c.start_dt, c.end_dt, c.location,
                   c.organizer, c.is_all_day, c.is_recurring
            FROM calendar_fts
            JOIN calendar_items c ON calendar_fts.rowid = c.id
            WHERE calendar_fts MATCH ?"""
        params.append(fts_q)
        alias = "c"
        order = "ORDER BY rank"
    else:
        select = """
            SELECT id, subject, start_dt, end_dt, location,
                   organizer, is_all_day, is_recurring
            FROM calendar_items
            WHERE 1=1"""
        alias = ""
        order = "ORDER BY start_dt DESC"

    def col(name):
        return f"{alias}.{name}" if alias else name

    if date_from:
        where.append(f"{col('start_dt')} >= ?")
        params.append(date_from + "T00:00:00Z")
    if date_to:
        where.append(f"{col('start_dt')} <= ?")
        params.append(date_to + "T23:59:59Z")

    where_sql = (" AND " + " AND ".join(where)) if where else ""

    try:
        total = db.execute(f"SELECT COUNT(*) FROM ({select}{where_sql})", params).fetchone()[0]
        results = db.execute(
            f"{select}{where_sql} {order} LIMIT {PAGE_SIZE} OFFSET {offset}", params
        ).fetchall()
    except Exception:
        total, results = 0, []

    return results, total


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    db = get_db()
    stats = {
        "emails":      db.execute("SELECT COUNT(*) FROM emails").fetchone()[0],
        "contacts":    db.execute("SELECT COUNT(*) FROM contacts").fetchone()[0],
        "calendar":    db.execute("SELECT COUNT(*) FROM calendar_items").fetchone()[0],
        "attachments": db.execute("SELECT COUNT(*) FROM email_attachments").fetchone()[0],
    }
    date_range = db.execute(
        "SELECT MIN(date_received), MAX(date_received) FROM emails WHERE date_received IS NOT NULL"
    ).fetchone()
    top_senders = db.execute(
        f"""SELECT {SENDER_EXPR} AS sender, sender_name, COUNT(*) AS cnt
            FROM emails
            GROUP BY lower({SENDER_EXPR})
            ORDER BY cnt DESC LIMIT 10"""
    ).fetchall()
    top_folders = db.execute(
        """SELECT folder_path, COUNT(*) AS cnt
           FROM emails GROUP BY folder_path ORDER BY cnt DESC LIMIT 10"""
    ).fetchall()
    recent = db.execute(
        f"""SELECT id, subject, sender_name,
                   {SENDER_EXPR} AS sender_email,
                   date_received, has_attachments
            FROM emails
            ORDER BY date_received DESC LIMIT 10"""
    ).fetchall()
    return render_template(
        "index.html",
        stats=stats,
        date_range=date_range,
        top_senders=top_senders,
        top_folders=top_folders,
        recent=recent,
    )


@app.route("/search")
def search():
    q           = request.args.get("q", "").strip()
    result_type = request.args.get("type", "emails")
    sender      = request.args.get("sender", "").strip()
    folder      = request.args.get("folder", "").strip()
    date_from   = request.args.get("date_from", "").strip()
    date_to     = request.args.get("date_to", "").strip()
    page        = max(1, int(request.args.get("page", 1)))
    offset      = (page - 1) * PAGE_SIZE

    db    = get_db()
    fts_q = prepare_fts_query(q)

    if result_type == "contacts":
        results, total = _search_contacts(db, fts_q, page, offset)
    elif result_type == "calendar":
        results, total = _search_calendar(db, fts_q, date_from, date_to, page, offset)
    else:
        result_type = "emails"
        results, total = _search_emails(db, fts_q, sender, folder, date_from, date_to, page, offset)

    pag     = paginate(page, total)
    folders = db.execute(
        "SELECT DISTINCT folder_path FROM emails WHERE folder_path IS NOT NULL ORDER BY folder_path"
    ).fetchall()

    return render_template(
        "search.html",
        q=q, result_type=result_type, results=results, pag=pag,
        sender=sender, folder=folder, date_from=date_from, date_to=date_to,
        folders=folders,
    )


@app.route("/email/<int:email_id>")
def view_email(email_id):
    db   = get_db()
    email = db.execute("SELECT * FROM emails WHERE id = ?", (email_id,)).fetchone()
    if not email:
        abort(404)
    recipients  = db.execute(
        "SELECT * FROM email_recipients WHERE email_id = ? ORDER BY type", (email_id,)
    ).fetchall()
    attachments = db.execute(
        "SELECT * FROM email_attachments WHERE email_id = ?", (email_id,)
    ).fetchall()
    html_body = sanitize_email_html(email["body_html"]) if email["body_html"] else None
    return render_template(
        "email.html",
        email=email, recipients=recipients, attachments=attachments, html_body=html_body,
    )


@app.route("/contacts")
def contacts():
    q      = request.args.get("q", "").strip()
    page   = max(1, int(request.args.get("page", 1)))
    offset = (page - 1) * PAGE_SIZE
    results, total = _search_contacts(get_db(), prepare_fts_query(q), page, offset)
    return render_template("contacts.html", q=q, results=results, pag=paginate(page, total))


@app.route("/contact/<int:contact_id>")
def view_contact(contact_id):
    contact = get_db().execute("SELECT * FROM contacts WHERE id = ?", (contact_id,)).fetchone()
    if not contact:
        abort(404)
    return render_template("contact.html", contact=contact)


@app.route("/calendar")
def calendar():
    q         = request.args.get("q", "").strip()
    date_from = request.args.get("date_from", "").strip()
    date_to   = request.args.get("date_to", "").strip()
    page      = max(1, int(request.args.get("page", 1)))
    offset    = (page - 1) * PAGE_SIZE
    results, total = _search_calendar(get_db(), prepare_fts_query(q), date_from, date_to, page, offset)
    return render_template(
        "calendar.html", q=q, results=results, pag=paginate(page, total),
        date_from=date_from, date_to=date_to,
    )


@app.route("/api/stats")
def api_stats():
    db = get_db()
    return jsonify({
        "totals": {
            "emails":      db.execute("SELECT COUNT(*) FROM emails").fetchone()[0],
            "contacts":    db.execute("SELECT COUNT(*) FROM contacts").fetchone()[0],
            "calendar":    db.execute("SELECT COUNT(*) FROM calendar_items").fetchone()[0],
            "attachments": db.execute("SELECT COUNT(*) FROM email_attachments").fetchone()[0],
        },
        "top_senders": [dict(r) for r in db.execute(
            f"""SELECT {SENDER_EXPR} AS sender, sender_name, COUNT(*) AS count
                FROM emails GROUP BY lower({SENDER_EXPR}) ORDER BY count DESC LIMIT 20"""
        ).fetchall()],
        "by_month": [dict(r) for r in db.execute(
            """SELECT substr(date_received,1,7) AS month, COUNT(*) AS count
               FROM emails WHERE date_received IS NOT NULL
               GROUP BY month ORDER BY month"""
        ).fetchall()],
    })


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Happy Hoppy – PST Email Search UI")
    parser.add_argument(
        "--db",
        default=r"C:\Users\lukeg\OneDrive\Documents\Processed Outlook databases\Test 1 crown hearing\outlook_data.db",
        help="Path to the SQLite database created by pst_to_sqlite.py",
    )
    parser.add_argument("--host",  default="127.0.0.1")
    parser.add_argument("--port",  type=int, default=5000)
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()

    db_path = str(Path(args.db).resolve())
    if not Path(db_path).exists():
        print(f"ERROR: Database not found: {db_path}")
        raise SystemExit(1)

    print("Happy Hoppy 🐰")
    print(f"  Database : {db_path}")
    print("  Building search indexes (first run may take a moment)…")
    ensure_fts_indexes(db_path)
    app.config["DB_PATH"] = db_path
    print(f"  Ready!   Open http://{args.host}:{args.port} in your browser.")
    app.run(host=args.host, port=args.port, debug=args.debug)
