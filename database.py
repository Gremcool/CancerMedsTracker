import os
import sqlite3
import pandas as pd

def get_connection():
    """Attempts to connect to the persistent cloud volume path. 
    If misconfigured or locked, it gracefully falls back to a local 
    project file so the app never crashes on launch."""
    railway_volume_path = "/data/medicines.db"
    
    try:
        # Check if we are on Railway and if the directory is accessible
        if os.environ.get("RAILWAY_ENVIRONMENT_ID") or os.path.exists("/data"):
            # Test if we can open/write to the volume database file
            conn = sqlite3.connect(railway_volume_path, check_same_thread=False)
            conn.execute("SELECT 1") 
            return conn
    except sqlite3.OperationalError:
        # If the volume isn't mounted yet, ignore the error and pass through
        pass
    
    # Safe fallback: Creates 'medicines.db' right inside your project directory
    return sqlite3.connect("medicines.db", check_same_thread=False)

def is_storage_permanent():
    """Helper function to verify if the app is successfully writing to permanent storage."""
    if os.path.exists("/data"):
        try:
            conn = sqlite3.connect("/data/medicines.db")
            conn.close()
            return True
        except sqlite3.OperationalError:
            return False
    return False

def create_tables():
    conn = get_connection()
    cur = conn.cursor()

    try:
        cur.execute("SELECT base_drug_name FROM medicines LIMIT 1")
    except sqlite3.OperationalError:
        cur.execute("DROP TABLE IF EXISTS medicines")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS medicines (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        base_drug_name TEXT,
        medicine_name TEXT UNIQUE
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

def get_unique_base_drugs():
    conn = get_connection()
    query = "SELECT DISTINCT base_drug_name FROM medicines ORDER BY base_drug_name"
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df["base_drug_name"].tolist()

def get_medicines_grid():
    conn = get_connection()
    query = """
    SELECT m.base_drug_name, m.medicine_name,
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

def get_latest_statuses():
    conn = get_connection()
    query = """
    SELECT m.base_drug_name as "Drug Group", 
           m.medicine_name as "Medicine Description", 
           COALESCE(mu.status, 'Open') as "Current Status", 
           COALESCE(mu.owner, '-') as "Action Owner", 
           COALESCE(mu.comment, '-') as "Latest Meeting Note",
           COALESCE(mu.update_date, '-') as "Last Updated"
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

def get_medicine_history(medicine_name):
    conn = get_connection()
    query = """
    SELECT update_date, status, owner, comment 
    FROM medicine_updates 
    WHERE medicine_name = ? 
    ORDER BY created_at DESC
    """
    df = pd.read_sql_query(query, conn, params=(medicine_name,))
    conn.close()
    return df

def save_update(medicine_name, update_date, status, owner, comment):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO medicine_updates (medicine_name, update_date, status, owner, comment)
    VALUES (?, ?, ?, ?, ?)
    """, (medicine_name, update_date, status, owner, comment))
    conn.commit()
    conn.close()

def get_dashboard_stats():
    conn = get_connection()
    cur = conn.cursor()
    
    cur.execute("SELECT COUNT(*) FROM medicines")
    total = cur.fetchone()[0] or 0
    
    query = """
    SELECT COALESCE(mu.status, 'Open') as status, COUNT(*) as count
    FROM medicines m
    LEFT JOIN (
        SELECT medicine_name, status,
               ROW_NUMBER() OVER (PARTITION BY medicine_name ORDER BY created_at DESC) as rn
        FROM medicine_updates
    ) mu ON m.medicine_name = mu.medicine_name AND mu.rn = 1
    GROUP BY COALESCE(mu.status, 'Open')
    """
    cur.execute(query)
    rows = cur.fetchall()
    conn.close()

    stats = {"total": total, "open": 0, "progress": 0, "supplier": 0, "escalated": 0, "completed": 0}
    mapping = {
        "Open": "open",
        "In Progress": "progress",
        "Waiting Supplier": "supplier",
        "Escalated": "escalated",
        "Completed": "completed"
    }

    for row in rows:
        status_val = row[0]
        count_val = row[1]
        if status_val in mapping:
            stats[mapping[status_val]] = count_val
            
    return stats