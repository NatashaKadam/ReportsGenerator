import os
import json
import time
import sqlite3
import datetime
import atexit
from .utilities import TEMP_FILES

SESSION_FILE_PATH = os.path.join(os.getenv('APPDATA'), 'ReportsGenerator', "session_data.json")
APP_DATA_DIR = os.path.join(os.getenv('APPDATA'), 'ReportsGenerator')
DB_PATH = os.path.join(APP_DATA_DIR, "user_data.db")

def load_session_file():
    default_data = {
        "name": "",
        "items": [],
        "total_amount": "₹0.00",
        "jr_asst_engineer": "",
        "deputy_engineer": "",
        "executive_engineer": ""
    }
    try:
        if not os.path.exists(SESSION_FILE_PATH):
            with open(SESSION_FILE_PATH, 'w') as f:
                json.dump(default_data, f)
            return default_data
        with open(SESSION_FILE_PATH, 'r') as f:
            content = f.read()
            if not content.strip():
                raise ValueError("Empty file")
            return json.loads(content)
    except (json.JSONDecodeError, ValueError, IOError) as e:
        try:
            os.remove(SESSION_FILE_PATH)
            with open(SESSION_FILE_PATH, 'w') as f:
                json.dump(default_data, f)
        except IOError:
            pass
        return default_data

def save_session_file(data):
    try:
        with open(SESSION_FILE_PATH, 'w') as f:
            json.dump(data, f)
        return True
    except (IOError, TypeError):
        return False

def db_setup():
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, data TEXT,
                timestamp TEXT
            )
        """)
        try:
            cur.execute("SELECT form_type FROM sessions LIMIT 1")
            cur.execute("CREATE TABLE sessions_new (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, data TEXT, timestamp TEXT)")
            cur.execute("INSERT INTO sessions_new (id, name, data, timestamp) SELECT id, name, data, timestamp FROM sessions")
            cur.execute("DROP TABLE sessions")
            cur.execute("ALTER TABLE sessions_new RENAME TO sessions")
        except sqlite3.OperationalError:
            pass
        cur.execute("DROP TABLE IF EXISTS form_b_entries")
        cur.execute("DROP TABLE IF EXISTS construction_items")
        conn.commit()
    except Exception as e:
        print(f"Database setup failed: {e}")
    finally:
        if conn:
            conn.close()

def save_session(name, data):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT id FROM sessions WHERE name = ?", (name,))
    existing_session = cur.fetchone()
    if existing_session:
        cur.execute("UPDATE sessions SET data = ?, timestamp = ? WHERE id = ?",
                   (json.dumps(data), datetime.datetime.now().isoformat(), existing_session[0]))
    else:
        cur.execute("INSERT INTO sessions (name, data, timestamp) VALUES (?, ?, ?)",
                   (name, json.dumps(data), datetime.datetime.now().isoformat()))
    conn.commit()
    conn.close()

def load_sessions():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT id, name, data, timestamp FROM sessions ORDER BY timestamp DESC")
    sessions = cur.fetchall()
    conn.close()
    return sessions

def delete_session_from_db(session_id):
    try:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute("DELETE FROM sessions WHERE id = ?", (session_id,))
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"Error deleting session: {e}")
        return False