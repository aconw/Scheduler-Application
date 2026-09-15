from __future__ import annotations
import sqlite3
from pathlib import Path
from datetime import datetime
import pandas as pd

DB_PATH = Path(__file__).parent / 'data' / 'scheduler_config.db'

SCHEMA = '''
CREATE TABLE IF NOT EXISTS training_rules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_code TEXT,
    cost_center TEXT,
    supervisory_org TEXT,
    training_title TEXT NOT NULL,
    prerequisite TEXT,
    topic TEXT,
    scheduling_policy TEXT,
    timing_modifier TEXT,
    timing_min REAL,
    timing_max REAL,
    active INTEGER NOT NULL DEFAULT 1,
    created_by TEXT,
    created_at TEXT,
    modified_by TEXT,
    modified_at TEXT,
    version INTEGER NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_rules_audience ON training_rules(job_code,cost_center,supervisory_org,active);
CREATE INDEX IF NOT EXISTS idx_rules_title ON training_rules(training_title,active);

CREATE TABLE IF NOT EXISTS equivalencies (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    required_training_title TEXT NOT NULL,
    equivalent_training_title TEXT NOT NULL,
    active INTEGER NOT NULL DEFAULT 1,
    created_by TEXT,
    created_at TEXT
);

CREATE TABLE IF NOT EXISTS locations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    location TEXT UNIQUE NOT NULL,
    reference_id TEXT,
    address_line_1 TEXT,
    address_line_2 TEXT,
    city TEXT,
    region TEXT,
    postal_code TEXT,
    country TEXT,
    latitude REAL,
    longitude REAL,
    coordinate_precision TEXT,
    coordinate_source TEXT,
    inactive TEXT
);

CREATE TABLE IF NOT EXISTS manual_routing (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    training_title TEXT UNIQUE NOT NULL,
    recipient_email TEXT,
    active INTEGER NOT NULL DEFAULT 1,
    modified_by TEXT,
    modified_at TEXT
);

CREATE TABLE IF NOT EXISTS approvals (
    event_key TEXT PRIMARY KEY,
    decision TEXT,
    reason TEXT,
    decided_by TEXT,
    decided_at TEXT
);

CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_time TEXT NOT NULL,
    user_name TEXT,
    action TEXT NOT NULL,
    entity_type TEXT,
    entity_key TEXT,
    detail TEXT
);
'''

def connect(db_path: Path | str = DB_PATH):
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    conn.commit()
    return conn

def audit(conn, action, entity_type='', entity_key='', detail='', user_name='user'):
    conn.execute(
        'INSERT INTO audit_log(event_time,user_name,action,entity_type,entity_key,detail) VALUES(?,?,?,?,?,?)',
        (datetime.now().isoformat(timespec='seconds'), user_name, action, entity_type, entity_key, detail)
    )
    conn.commit()

def load_rules(conn, active_only=True):
    q='SELECT * FROM training_rules'
    if active_only: q += ' WHERE active=1'
    return pd.read_sql_query(q, conn)

def load_equivalencies(conn, active_only=True):
    q='SELECT * FROM equivalencies'
    if active_only: q += ' WHERE active=1'
    return pd.read_sql_query(q, conn)

def load_locations(conn, active_only=False):
    q='SELECT * FROM locations'
    if active_only: q += " WHERE COALESCE(LOWER(inactive),'') NOT IN ('yes','y','true','1')"
    return pd.read_sql_query(q, conn)

def load_manual_routing(conn, active_only=True):
    q='SELECT * FROM manual_routing'
    if active_only: q += ' WHERE active=1'
    return pd.read_sql_query(q, conn)
