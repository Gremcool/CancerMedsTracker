import os
import sqlite3
import pandas as pd
import json

def get_connection():
    railway_volume_path = "/data/medicines.db"
    if os.environ.get("RAILWAY_ENVIRONMENT_ID") or os.path.exists("/data"):
        os.makedirs("/data", exist_ok=True)
        return sqlite3.connect(railway_volume_path, check_same_thread=False)
    return sqlite3.connect("medicines.db", check_same_thread=False)

def is_storage_permanent():
    path = "/data/medicines.db"
    if os.path.exists("/data"):
        try:
            conn = sqlite3.connect(path)
            conn.close()
            return True
        except:
            return False
    return False

def update_schema():
    """Safely upgrades database schema without losing existing data or records."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(medicines)")
    columns = [info[1] for info in cur.fetchall()]
    
    if "stock_on_hand" not in columns:
        cur.execute("ALTER TABLE medicines ADD COLUMN stock_on_hand REAL DEFAULT 0")
    if "in_transit" not in columns:
        cur.execute("ALTER TABLE medicines ADD COLUMN in_transit REAL DEFAULT 0")
    if "amc" not in columns:
        cur.execute("ALTER TABLE medicines ADD COLUMN amc REAL DEFAULT 0")
    if "request_qty" not in columns:
        cur.execute("ALTER TABLE medicines ADD COLUMN request_qty REAL DEFAULT 0")
    if "category" not in columns:
        cur.execute("ALTER TABLE medicines ADD COLUMN category TEXT DEFAULT 'Medicines'")
    if "custom_fields" not in columns:
        cur.execute("ALTER TABLE medicines ADD COLUMN custom_fields TEXT DEFAULT '{}'")
        
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
        in_transit REAL DEFAULT 0,
        amc REAL DEFAULT 0,
        request_qty REAL DEFAULT 0,
        category TEXT DEFAULT 'Medicines',
        custom_fields TEXT DEFAULT '{}'
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
    cur.execute("""
    CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS category_columns (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category_name TEXT,
        column_name TEXT,
        UNIQUE(category_name, column_name)
    )
    """)
    
    # Pre-seed standard categories
    cur.execute("INSERT OR IGNORE INTO categories (name) VALUES ('Medicines')")
    cur.execute("INSERT OR IGNORE INTO categories (name) VALUES ('Consumables')")
    
    conn.commit()
    conn.close()

# --- CATEGORY & CUSTOM COLUMNS MANAGERS ---

def get_all_categories():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT name FROM categories ORDER BY id ASC")
    db_categories = [row[0] for row in cur.fetchall()]
    
    cur.execute("SELECT DISTINCT category FROM medicines WHERE category IS NOT NULL")
    med_categories = [row[0] for row in cur.fetchall()]
    conn.close()
    
    all_cats = list(dict.fromkeys(["Medicines", "Consumables"] + db_categories + med_categories))
    return all_cats

def add_category(category_name):
    clean_name = category_name.strip()
    if not clean_name:
        return False
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("INSERT INTO categories (name) VALUES (?)", (clean_name,))
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        success = True
    conn.close()
    return success

def get_category_custom_columns(category_name):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT column_name FROM category_columns WHERE category_name = ? ORDER BY id ASC", (category_name,))
    cols = [row[0] for row in cur.fetchall()]
    conn.close()
    return cols

def add_category_custom_column(category_name, column_name):
    clean_col = column_name.strip()
    if not clean_col:
        return False
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("INSERT INTO category_columns (category_name, column_name) VALUES (?, ?)", (category_name, clean_col))
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        success = True
    conn.close()
    return success

def delete_category_custom_column(category_name, column_name):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM category_columns WHERE category_name = ? AND column_name = ?", (category_name, column_name))
    conn.commit()
    conn.close()

# --- MEDICINE CRUD OPERATIONS ---

IMPORT_NUMERIC_FIELDS = ["stock_on_hand", "in_transit", "amc", "request_qty"]

def _clean_num(val):
    """Returns a float, or None when the value is blank / not numeric."""
    try:
        if val is None or pd.isna(val):
            return None
        return float(val)
    except (TypeError, ValueError):
        return None

def get_existing_item_map():
    """READ-ONLY. Maps lower-cased item name -> (stored_name, category) for every record."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT medicine_name, COALESCE(category, 'Medicines') FROM medicines")
    rows = cur.fetchall()
    conn.close()
    return {str(n).strip().lower(): (n, c) for n, c in rows if n}

STATUS_NEW = "New"
STATUS_UPDATE = "Update (already in this category)"
STATUS_BLOCKED_PREFIX = "Skipped - already exists in "

def classify_import_rows(df, category):
    """READ-ONLY. Adds an 'import_status' column telling the user exactly what the import will do."""
    if df is None or df.empty:
        return df
    existing = get_existing_item_map()
    statuses = []
    for name in df["medicine_name"]:
        hit = existing.get(str(name).strip().lower())
        if hit is None:
            statuses.append(STATUS_NEW)
        elif hit[1] == category:
            statuses.append(STATUS_UPDATE)
        else:
            statuses.append(f"{STATUS_BLOCKED_PREFIX}'{hit[1]}'")
    out = df.copy()
    out.insert(0, "import_status", statuses)
    return out

def save_medicines(df, category=None, custom_cols=None):
    """
    Safely imports rows from an Excel / pasted / typed table.

    Safety rules (this is what protects existing data):
      * A record that already lives in ANOTHER category is never touched or re-labelled.
      * A record already in the SAME category only has the columns that were actually
        supplied updated (stock is never silently reset to 0).
      * Brand-new names are inserted with the target category exactly as given.

    Returns a summary dict: {"added", "updated", "skipped": [(name, existing_category)]}.
    """
    custom_cols = [c for c in (custom_cols or []) if c in df.columns]
    summary = {"added": 0, "updated": 0, "skipped": []}
    if df is None or df.empty:
        return summary

    conn = get_connection()
    cur = conn.cursor()

    target_default = category or "Medicines"
    if target_default:
        cur.execute("INSERT OR IGNORE INTO categories (name) VALUES (?)", (target_default,))

    cur.execute("SELECT medicine_name, COALESCE(category, 'Medicines'), COALESCE(custom_fields, '{}') FROM medicines")
    existing = {str(n).strip().lower(): (n, c, cf) for n, c, cf in cur.fetchall() if n}

    for _, row in df.iterrows():
        name = str(row.get("medicine_name", "")).strip()
        if not name or name.lower() == "nan":
            continue
        cat = category or row.get("category") or "Medicines"
        key = name.lower()

        row_custom = {}
        for c in custom_cols:
            v = row.get(c)
            if v is not None and not pd.isna(v) and str(v).strip() != "":
                row_custom[c] = str(v).strip()

        if key in existing:
            stored_name, existing_cat, existing_cf = existing[key]
            if existing_cat != cat:
                summary["skipped"].append((stored_name, existing_cat))
                continue

            sets, params = [], []
            for field in IMPORT_NUMERIC_FIELDS:
                if field in df.columns:
                    val = _clean_num(row.get(field))
                    if val is not None:
                        sets.append(f"{field} = ?")
                        params.append(val)
            if row_custom:
                try:
                    merged = json.loads(existing_cf) if existing_cf else {}
                except (TypeError, ValueError):
                    merged = {}
                merged.update(row_custom)
                sets.append("custom_fields = ?")
                params.append(json.dumps(merged))
            if sets:
                params.append(stored_name)
                cur.execute(f"UPDATE medicines SET {', '.join(sets)} WHERE medicine_name = ?", params)
            summary["updated"] += 1
        else:
            base = row.get("base_drug_name", "")
            if base is None or pd.isna(base) or str(base).strip() == "":
                base = name.split()[0].upper() if name.split() else name.upper()
            nums = [(_clean_num(row.get(f)) if f in df.columns else None) for f in IMPORT_NUMERIC_FIELDS]
            nums = [0 if v is None else v for v in nums]
            cur.execute("""
                INSERT INTO medicines (base_drug_name, medicine_name, stock_on_hand, in_transit, amc, request_qty, category, custom_fields)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (str(base).strip(), name, nums[0], nums[1], nums[2], nums[3], cat, json.dumps(row_custom)))
            existing[key] = (name, cat, json.dumps(row_custom))
            summary["added"] += 1

    conn.commit()
    conn.close()
    return summary

def update_stock_levels(medicine_name, soh, transit, amc=0, request_qty=0, custom_fields_dict=None):
    if custom_fields_dict is None:
        custom_fields_dict = {}
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
    UPDATE medicines 
    SET stock_on_hand = ?, in_transit = ?, amc = ?, request_qty = ?, custom_fields = ? 
    WHERE medicine_name = ?
    """, (soh, transit, amc, request_qty, json.dumps(custom_fields_dict), medicine_name))
    conn.commit()
    conn.close()

def add_new_medicine(base_drug_name, medicine_name, stock_on_hand=0, in_transit=0, amc=0, request_qty=0, category="Medicines", custom_fields_dict=None):
    if custom_fields_dict is None:
        custom_fields_dict = {}
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute("""
        INSERT INTO medicines (base_drug_name, medicine_name, stock_on_hand, in_transit, amc, request_qty, category, custom_fields)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (base_drug_name, medicine_name, stock_on_hand, in_transit, amc, request_qty, category, json.dumps(custom_fields_dict)))
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        success = False
    conn.close()
    return success

def delete_medicine(medicine_name, deleted_by_person, delete_date_time):
    """Logs deletion audit trail before removing record from medicines table."""
    conn = get_connection()
    cur = conn.cursor()
    
    # Write audit log entry
    cur.execute("""
    INSERT INTO medicine_updates (medicine_name, update_date, status, owner, comment)
    VALUES (?, ?, 'DELETED', ?, ?)
    """, (medicine_name, delete_date_time, deleted_by_person, f"Item permanently deleted by {deleted_by_person}."))
    
    # Remove record
    cur.execute("DELETE FROM medicines WHERE medicine_name = ?", (medicine_name,))
    
    conn.commit()
    conn.close()

def delete_medicines_bulk(medicine_names, deleted_by_person, delete_date_time, category=None, reason=""):
    """
    Deletes several items in ONE transaction and writes an audit-trail entry for each one
    (same 'DELETED' log format as delete_medicine). If 'category' is given, only items that
    really belong to that category can be removed. Returns the list of names actually deleted.
    """
    conn = get_connection()
    cur = conn.cursor()
    deleted = []
    try:
        for name in dict.fromkeys(medicine_names or []):
            if category:
                cur.execute(
                    "SELECT 1 FROM medicines WHERE medicine_name = ? AND COALESCE(category, 'Medicines') = ?",
                    (name, category))
            else:
                cur.execute("SELECT 1 FROM medicines WHERE medicine_name = ?", (name,))
            if not cur.fetchone():
                continue
            comment = f"Item permanently deleted by {deleted_by_person} (bulk clean-up)."
            if reason and reason.strip():
                comment += f" Reason: {reason.strip()}"
            cur.execute("""
            INSERT INTO medicine_updates (medicine_name, update_date, status, owner, comment)
            VALUES (?, ?, 'DELETED', ?, ?)
            """, (name, delete_date_time, deleted_by_person, comment))
            cur.execute("DELETE FROM medicines WHERE medicine_name = ?", (name,))
            deleted.append(name)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return deleted

def get_medicines_grid(category="Medicines"):
    conn = get_connection()
    query = """
    SELECT m.base_drug_name, m.medicine_name, 
           CAST(COALESCE(m.stock_on_hand, 0) AS INTEGER) as stock_on_hand, 
           CAST(COALESCE(m.in_transit, 0) AS INTEGER) as in_transit,
           CAST(COALESCE(m.amc, 0) AS INTEGER) as amc,
           CAST(COALESCE(m.request_qty, 0) AS INTEGER) as request_qty,
           COALESCE(m.custom_fields, '{}') as custom_fields,
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
    WHERE COALESCE(m.category, 'Medicines') = ? OR (? = 'Medicines' AND m.category IS NULL)
    ORDER BY m.base_drug_name, m.medicine_name
    """
    df = pd.read_sql_query(query, conn, params=(category, category))
    conn.close()
    return df

def get_latest_statuses(category="Medicines"):
    conn = get_connection()
    query = """
    SELECT m.base_drug_name as "Drug Group", 
           m.medicine_name as "Medicine", 
           CAST(COALESCE(m.stock_on_hand, 0) AS INTEGER) as "SOH", 
           CAST(COALESCE(m.in_transit, 0) AS INTEGER) as "Transit",
           CAST(COALESCE(m.amc, 0) AS INTEGER) as "AMC",
           CAST(COALESCE(m.request_qty, 0) AS INTEGER) as "Request",
           COALESCE(m.custom_fields, '{}') as "CustomFields",
           COALESCE(mu.status, 'Open') as "Status", 
           COALESCE(mu.owner, '-') as "Owner", 
           COALESCE(mu.comment, '-') as "Note",
           COALESCE(mu.update_date, '-') as "Updated"
    FROM medicines m
    LEFT JOIN (
        SELECT medicine_name, status, owner, comment, update_date,
               ROW_NUMBER() OVER (PARTITION BY medicine_name ORDER BY created_at DESC) as rn
        FROM medicine_updates
    ) mu ON m.medicine_name = mu.medicine_name AND mu.rn = 1
    WHERE COALESCE(m.category, 'Medicines') = ? OR (? = 'Medicines' AND m.category IS NULL)
    ORDER BY m.base_drug_name, m.medicine_name
    """
    df = pd.read_sql_query(query, conn, params=(category, category))
    conn.close()
    return df

def get_medicine_history(medicine_name):
    conn = get_connection()
    query = "SELECT update_date, status, owner, comment FROM medicine_updates WHERE medicine_name = ? ORDER BY created_at DESC"
    df = pd.read_sql_query(query, conn, params=(medicine_name,))
    conn.close()
    return df

def save_update(medicine_name, update_date, status, owner, comment):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("INSERT INTO medicine_updates (medicine_name, update_date, status, owner, comment) VALUES (?, ?, ?, ?, ?)", (medicine_name, update_date, status, owner, comment))
    conn.commit()
    conn.close()

def get_dashboard_stats(category="Medicines"):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM medicines WHERE COALESCE(category, 'Medicines') = ? OR (? = 'Medicines' AND category IS NULL)", (category, category))
    total = cur.fetchone()[0] or 0
    
    cur.execute("""
    SELECT COALESCE(mu.status, 'Open'), COUNT(*) 
    FROM medicines m 
    LEFT JOIN (
        SELECT medicine_name, status, ROW_NUMBER() OVER (PARTITION BY medicine_name ORDER BY created_at DESC) as rn 
        FROM medicine_updates
    ) mu ON m.medicine_name = mu.medicine_name AND mu.rn = 1 
    WHERE COALESCE(m.category, 'Medicines') = ? OR (? = 'Medicines' AND m.category IS NULL)
    GROUP BY COALESCE(mu.status, 'Open')
    """, (category, category))
    
    rows = cur.fetchall()
    conn.close()
    
    stats = {"total": total, "open": 0, "progress": 0, "supplier": 0, "escalated": 0, "completed": 0}
    mapping = {"Open": "open", "In Progress": "progress", "Waiting Supplier": "supplier", "Escalated": "escalated", "Completed": "completed"}
    for status_val, count_val in rows:
        if status_val in mapping: 
            stats[mapping[status_val]] = count_val
    return stats