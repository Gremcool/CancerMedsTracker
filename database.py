import os
import sqlite3
import pandas as pd

def get_connection():
    """Establishes connection to a persistent database folder on Railway, 
    or a fallback local file on your machine."""
    # Checks if running inside Railway's cloud environment environment
    if os.environ.get("RAILWAY_ENVIRONMENT_ID") or os.path.exists("/data"):
        return sqlite3.connect("/data/medicines.db", check_same_thread=False)
    
    # Local fallback for your testing offline
    return sqlite3.connect("medicines.db", check_same_thread=False)

def create_tables():
    conn = get_connection()
    cur = conn.cursor()

    # SCHEMA MIGRATION CHECK: Safe verification step to clear old column layouts
    try:
        cur.execute("SELECT base_drug_name FROM medicines LIMIT 1")
    except sqlite3.OperationalError:
        cur.execute("DROP TABLE IF EXISTS medicines")

    # Re-create clean master table with base drug matching rules
    cur.execute("""
    CREATE TABLE IF NOT EXISTS medicines (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        base_drug_name TEXT,
        medicine_name TEXT UNIQUE
    )
    """)

    # Keeps meeting actions intact over time chronologically
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
    """Retrieves generic parent names to build clean meeting rows."""
    conn = get_connection()
    query = "SELECT DISTINCT base_drug_name FROM medicines ORDER BY base_drug_name"
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df["base_drug_name"].tolist()

def get_medicines_grid():
    """Fetches full data mapping directly into the inline spreadsheet view."""
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
    """Generates the main overview table showing current status details for all items on Dashboard."""
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
    """Retrieves the full tracking timeline history for an item."""
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
    """Appends a new chronological update log row into the table database."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO medicine_updates (medicine_name, update_date, status, owner, comment)
    VALUES (?, ?, ?, ?, ?)
    """, (medicine_name, update_date, status, owner, comment))
    conn.commit()
    conn.close()

def get_dashboard_stats():
    """Aggregates metrics directly from the absolute latest status entries."""
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