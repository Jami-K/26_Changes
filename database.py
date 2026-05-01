import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'changes.db')


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_conn() as conn:
        conn.executescript('''
            CREATE TABLE IF NOT EXISTS changes (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                notice_date      TEXT,
                product_code     TEXT,
                product_name     TEXT,
                changed_material TEXT,
                mixable          TEXT,
                change_content   TEXT,
                apply_date       TEXT,
                notes            TEXT,
                design_file        TEXT DEFAULT '',
                pdf_path           TEXT DEFAULT '',
                actual_apply_date  TEXT DEFAULT '',
                row_hash           TEXT UNIQUE
            );
            CREATE TABLE IF NOT EXISTS settings (
                key   TEXT PRIMARY KEY,
                value TEXT
            );
            CREATE TABLE IF NOT EXISTS recipients (
                id    INTEGER PRIMARY KEY AUTOINCREMENT,
                name  TEXT NOT NULL,
                email TEXT NOT NULL
            );
        ''')
        try:
            conn.execute("ALTER TABLE changes ADD COLUMN pdf_path TEXT DEFAULT ''")
        except Exception:
            pass
        try:
            conn.execute("ALTER TABLE changes ADD COLUMN actual_apply_date TEXT DEFAULT ''")
        except Exception:
            pass


def get_all_changes(search=''):
    with get_conn() as conn:
        if search:
            rows = conn.execute('''
                SELECT * FROM changes
                WHERE product_code     LIKE ?
                   OR product_name     LIKE ?
                   OR changed_material LIKE ?
                ORDER BY notice_date DESC, id DESC
            ''', (f'%{search}%',) * 3).fetchall()
        else:
            rows = conn.execute(
                'SELECT * FROM changes ORDER BY notice_date DESC, id DESC'
            ).fetchall()
    return [dict(r) for r in rows]


def get_change_by_id(cid):
    with get_conn() as conn:
        row = conn.execute('SELECT * FROM changes WHERE id = ?', (cid,)).fetchone()
    return dict(row) if row else None


def upsert_change(data):
    with get_conn() as conn:
        cur = conn.execute('''
            INSERT OR IGNORE INTO changes
            (notice_date, product_code, product_name, changed_material,
             mixable, change_content, apply_date, notes, design_file, row_hash)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            data.get('notice_date', ''),     data.get('product_code', ''),
            data.get('product_name', ''),    data.get('changed_material', ''),
            data.get('mixable', ''),         data.get('change_content', ''),
            data.get('apply_date', ''),      data.get('notes', ''),
            data.get('design_file', ''),     data['row_hash'],
        ))
        return cur.rowcount


def clear_changes():
    with get_conn() as conn:
        conn.execute('DELETE FROM changes')


def get_setting(key, default=''):
    with get_conn() as conn:
        row = conn.execute('SELECT value FROM settings WHERE key = ?', (key,)).fetchone()
    return row[0] if row else default


def set_setting(key, value):
    with get_conn() as conn:
        conn.execute(
            'INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)', (key, value)
        )


def get_recipients():
    with get_conn() as conn:
        rows = conn.execute('SELECT * FROM recipients ORDER BY id').fetchall()
    return [dict(r) for r in rows]


def add_recipient(name, email):
    with get_conn() as conn:
        conn.execute('INSERT INTO recipients (name, email) VALUES (?, ?)', (name, email))


def delete_recipient(rid):
    with get_conn() as conn:
        conn.execute('DELETE FROM recipients WHERE id = ?', (rid,))


def set_pdf_path(cid, path):
    with get_conn() as conn:
        conn.execute('UPDATE changes SET pdf_path = ? WHERE id = ?', (path, cid))


def set_actual_apply_date(cid, date):
    with get_conn() as conn:
        conn.execute('UPDATE changes SET actual_apply_date = ? WHERE id = ?', (date, cid))
