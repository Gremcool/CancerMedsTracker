import os
import sqlite3
import pandas as pd

def get_connection():
    railway_volume_path = "/data/medicines.db"
    # Use persistent volume if available, else local file
    if os.environ.get("RAILWAY_ENVIRONMENT_ID") or os.path.exists("/data"):
        os.makedirs("/data", exist_ok=True)
        return sqlite3.connect(railway_volume_path, check_same_thread=False)
    return sqlite3.connect("medicines.db", check_same_thread=False)

def update_schema():
    """Safely adds Stock and Transit columns to your existing database."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(medicines)")
    columns = [info[1] for info in cur.fetchall()]
    
    if "stock_on_hand" not in columns:
        cur.execute("ALTER TABLE medicines ADD COLUMN stock_on_hand REAL DEFAULT 0")
    if "in_transit" not in columns:
        cur.execute("ALTER TABLE medicines ADD COLUMN in_transit REAL DEFAULT 0")
    conn.commit()
    conn.close()

def create_tables():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS medicines (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        base_drug_name TEXT,
        medicine_name TEXT UNIQUE,
        stock_on_hand REAL DEFAULT 0,
        in_transit REAL DEFAULT 0
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS medicine_updates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        medicine_name TEXT,
        update_date TEXT,
        status TEXT,
        owner TEXT,
        comment TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    conn.commit()
    conn.close()

def update_stock_levels(medicine_name, soh, transit):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("UPDATE medicines SET stock_on_hand = ?, in_transit = ? WHERE medicine_name = ?", (soh, transit, medicine_name))
    conn.commit()
    conn.close()

def get_medicines_grid():
    conn = get_connection()
    query = """
    SELECT m.base_drug_name, m.medicine_name, 
           COALESCE(m.stock_on_hand, 0) as stock_on_hand, 
           COALESCE(m.in_transit, 0) as in_transit,
           COALESCE(mu.status, 'Open') as status,
           COALESCE(mu.owner, '') as owner,
           COALESCE(mu.comment, '-') as last_comment,
           COALESCE(mu.update_date, '-') as last_updated
    FROM medicines m
    LEFT JOIN (
        SELECT medicine_name, status, owner, comment, update_date,
               ROW_NUMBER() OVER (PARTITION BY medicine_name ORDER BY created_at DESC) as rn
        FROM medicine_updates
    ) mu ON m.medicine_name = mu.medicine_name AND mu.rn = 1
    ORDER BY m.base_drug_name, m.medicine_name
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

# ... (Keep all your existing functions like get_latest_statuses, get_medicine_history, save_update, get_dashboard_stats)