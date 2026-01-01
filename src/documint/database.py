import sqlite3
import os
from datetime import datetime

class DatabaseManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS jobs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                total_items INTEGER,
                success_count INTEGER,
                fail_count INTEGER,
                mode TEXT
            )
        ''')
        conn.commit()
        conn.close()

    def log_job(self, total, success, fail, mode="Standard"):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute('''
            INSERT INTO jobs (timestamp, total_items, success_count, fail_count, mode)
            VALUES (?, ?, ?, ?, ?)
        ''', (datetime.now().isoformat(), total, success, fail, mode))
        conn.commit()
        conn.close()

    def get_stats(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        
        # Aggregate stats
        c.execute("SELECT COUNT(*) as total_jobs, SUM(total_items) as total_emails FROM jobs")
        summary = dict(c.fetchone())
        
        # Recent jobs
        c.execute("SELECT * FROM jobs ORDER BY id DESC LIMIT 5")
        recent = [dict(row) for row in c.fetchall()]
        
        conn.close()
        return {
            "summary": summary,
            "recent": recent
        }
