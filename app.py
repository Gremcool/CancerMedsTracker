import os
import sys
import json
import re

# Ensure local imports resolve safely
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import pandas as pd
from datetime import datetime, date

from database import *
from excel_import import *

# Configure page layout
st.set_page_config(
    page_title="Medical Inventory System",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS styling for non-wrapping headers, clean badges, and Green/Red action buttons
st.markdown("""
    <style>
        /* Top padding to clear Streamlit header banner */
        .block-container {
            padding-top: 3.2rem !important;
            padding-bottom: 2rem !important;
            margin-top: 0rem !important;
        }
        /* Completely hide sidebar collapse control */
        [data-testid="collapsedControl"] {
            display: none !important;
        }
        /* Standard input field styling */
        div[data-testid="stInput"] input {
            border: 1px solid #d0d7de !important;
            border-radius: 6px !important;
            min-width: 95px !important;
        }
        /* Page title positioning */
        div[data-testid="stTitle"] {
            margin-top: -0.6rem !important;
            padding-top: 0rem !important;
        }
        /* Ensure selectbox options like Escalated fit cleanly */
        div[data-baseweb="select"] {
            min-width: 135px !important;
        }
        /* Prevent Column Headers from Wrapping */
        .grid-header-text {
            white-space: nowrap !important;
            font-weight: bold;
            font-size: 0.88rem;
            color: #333;
        }
        /* Search Box Highlight */
        .search-box-container input {
            border: 2px solid #ff4b4b !important;
            border-radius: 8px !important;
            font-size: 1rem !important;
        }
        /* KPI Cards */
        .kpi-card {
            border-radius: 10px;
            padding: 15px 10px;
            text-align: center;
            box-shadow: 0px 2px 6px rgba(0,0,0,0.05);
            font-family: system-ui, -apple-system, sans-serif;
            margin-bottom: 10px;
        }
        .kpi-title {
            font-size: 0.78rem;
            font-weight: 700;
            letter-spacing: 0.5px;
            text-transform: uppercase;
            margin-bottom: 6px;
        }
        .kpi-value {
            font-size: 1.8rem;
            font-weight: 800;
            line-height: 1.1;
        }
        .kpi-total { background-color: #f8f9fa; border: 1px solid #e9ecef; color: #212529; }
        .kpi-open { background-color: #e3f2fd; border: 1px solid #bbdefb; color: #0d47a1; }
        .kpi-progress { background-color: #fffde7; border: 1px solid #fff9c4; color: #f57f17; }
        .kpi-supplier { background-color: #f3e5f5; border: 1px solid #e1bee7; color: #4a148c; }
        .kpi-escalated { background-color: #ffebee; border: 1px solid #ffcdd2; color: #c62828; }
        .kpi-completed { background-color: #e8f5e9; border: 1px solid #c8e6c9; color: #2e7d32; }
        
        /* Clean Status Badges (No Emojis) */
        .badge-completed {
            background-color: #28a745 !important;
            color: #ffffff !important;
            padding: 4px 12px;
            border-radius: 6px;
            font-weight: bold;
            display: inline-block;
        }
        .badge-escalated {
            background-color: #dc3545 !important;
            color: #ffffff !important;
            padding: 4px 12px;
            border-radius: 6px;
            font-weight: bold;
            display: inline-block;
        }
        .badge-progress {
            background-color: #ffc107 !important;
            color: #212529 !important;
            padding: 4px 12px;
            border-radius: 6px;
            font-weight: bold;
            display: inline-block;
        }
        .badge-open {
            background-color: #17a2b8 !important;
            color: #ffffff !important;
            padding: 4px 12px;
            border-radius: 6px;
            font-weight: bold;
            display: inline-block;
        }
        .badge-supplier {
            background-color: #6c757d !important;
            color: #ffffff !important;
            padding: 4px 12px;
            border-radius: 6px;
            font-weight: bold;
            display: inline-block;
        }
        
        /* DISTINCT ACTION BUTTON STYLING (Green Save Icon, Red Delete Popover) */
        div.save-btn-container button {
            background-color: #28a745 !important;
            color: white !important;
            border: none !important;
            font-weight: bold !important;
            border-radius: 6px !important;
            padding: 6px 0px !important;
            width: 100% !important;
            min-width: 36px !important;
            font-size: 1.1rem !important;
        }
        div.save-btn-container button:hover {
            background-color: #218838 !important;
            box-shadow: 0px 2px 4px rgba(0,0,0,0.2) !important;
        }
        div.delete-btn-container div[data-testid="stPopover"] button {
            background-color: #dc3545 !important;
            color: white !important;
            border: none !important;
            font-weight: bold !important;
            border-radius: 6px !important;
            padding: 6px 0px !important;
            width: 100% !important;
            min-width: 36px !important;
            font-size: 1.1rem !important;
        }
        div.delete-btn-container div[data-testid="stPopover"] button:hover {
            background-color: #c82333 !important;
            box-shadow: 0px 2px 4px rgba(0,0,0,0.2) !important;
        }
    </style>
""", unsafe_allow_html=True)

# Connect and activate database tables
create_tables()
update_schema()

# Helper function to parse user typed numbers (e.g. "100,000" or "100000") safely into integers
def parse_number(val):
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return int(val)
    clean_str = re.sub(r'[^\d-]', '', str(val))
    return int(clean_str) if clean_str else 0

# Helper function to render clean color status badges
def render_status_badge(status_str):
    if status_str == "Completed":
        return '<span class="badge-completed">Completed</span>'
    elif status_str == "Escalated":
        return '<span class="badge-escalated">Escalated</span>'
    elif status_str == "In Progress":
        return '<span class="badge-progress">In Progress</span>'
    elif status_str == "Waiting Supplier":
        return '<span class="badge-supplier">Waiting Supplier</span>'
    else:
        return '<span class="badge-open">Open</span>'

# ====================================================
# DYNAMIC CATEGORY BAR & CONTROL PANEL
# ====================================================
categories_list = get_all_categories()

c_bar1, c_bar2, c_bar3 = st.columns([3.5, 1.0, 1.0])

active_category = c_bar1.radio(
    "Active Category",
    options=categories_list,
    horizontal=True,
    key="active_category_selector",
    label_visibility="collapsed"
)

with c_bar2.popover("➕ New Category"):
    st.markdown("### Create Category")
    new_cat_input = st.text_input("Category Name", placeholder="e.g. Surgical Supplies...", key="input_new_category").strip()
    if st.button("🚀 Create Category", type="primary"):
        if new_cat_input:
            add_category(new_cat_input)
            st.session_state["input_new_category"] = ""
            st.rerun()
        else:
            st.error("Please enter a category name.")

custom_cols = get_category_custom_columns(active_category)

with c_bar3.popover("⚙️ Custom Columns"):
    st.markdown(f"### Columns for '{active_category}'")
    st.caption("Add dynamic custom attributes to this category without changing code.")
    
    if custom_cols:
        st.write("**Existing Custom Columns:**")
        for col in custom_cols:
            col_del1, col_del2 = st.columns([3, 1])
            col_del1.write(f"• {col}")
            if col_del2.button("❌", key=f"del_col_{active_category}_{col}"):
                delete_category_custom_column(active_category, col)
                st.rerun()
    else:
        st.info("No custom columns added yet.")
        
    st.markdown("---")
    new_col_input = st.text_input("New Column Name", placeholder="e.g. Expiry Date, Batch No...", key="input_new_column").strip()
    if st.button("➕ Add Field", type="primary", key="add_custom_col_btn"):
        if new_col_input:
            add_category_custom_column(active_category, new_col_input)
            st.session_state["input_new_column"] = ""
            st.rerun()

st.title(f"📂 {active_category} Management")

# ====================================================
# QUICK ITEM & STATUS SEARCH BAR
# ====================================================
st.markdown('<div class="search-box-container">', unsafe_allow_html=True)
search_query = st.text_input(
    "🔍 Search Item or Status",
    placeholder="🔎 Search by Item Name, Group, Owner, or Status (e.g. Escalated, Completed, Open)...",
    key="global_item_search",
    label_visibility="collapsed"
).strip().lower()
st.markdown('</div>', unsafe_allow_html=True)

tabs = st.tabs([
    "📊 Summary Action Dashboard",
    "🗓️ Live Meeting Review",
    "📥 Master Import Portal"
])

def get_board_column_weights(custom_cols_count):
    weights = [0.4, 2.2, 1.1, 1.1, 1.1, 1.1]
    for _ in range(custom_cols_count):
        weights.append(1.0)
    weights.extend([1.4, 1.2, 1.8, 1.8, 0.6, 0.6])
    return weights

# ====================================================
# TAB 1: ACTION DASHBOARD
# ====================================================
with tabs[0]:
    st.subheader(f"Meeting Action Progress KPIs — {active_category}")
    stats = get_dashboard_stats(active_category)

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.markdown(f'<div class="kpi-card kpi-total"><div class="kpi-title">Total Items</div><div class="kpi-value">{stats["total"]:,}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card kpi-open"><div class="kpi-title">Open Actions</div><div class="kpi-value">{stats["open"]:,}</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card kpi-progress"><div class="kpi-title">In Progress</div><div class="kpi-value">{stats["progress"]:,}</div></div>', unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card kpi-supplier"><div class="kpi-title">Waiting Supplier</div><div class="kpi-value">{stats["supplier"]:,}</div></div>', unsafe_allow_html=True)
    k5.markdown(f'<div class="kpi-card kpi-escalated"><div class="kpi-title">Escalated</div><div class="kpi-value">{stats["escalated"]:,}</div></div>', unsafe_allow_html=True)
    k6.markdown(f'<div class="kpi-card kpi-completed"><div class="kpi-title">Completed</div><div class="kpi-value">{stats["completed"]:,}</div></div>', unsafe_allow_html=True)

    st.markdown("---")
    st.write(f"### 📋 Master Meeting Tracking Log — {active_category}")
    
    master_grid = get_latest_statuses(active_category)
    
    if search_query and not master_grid.empty:
        master_grid = master_grid[
            master_grid["Drug Group"].astype(str).str.lower().str.contains(search_query, na=False) |
            master_grid["Medicine"].astype(str).str.lower().str.contains(search_query, na=False) |
            master_grid["Owner"].astype(str).str.lower().str.contains(search_query, na=False) |
            master_grid["Status"].astype(str).str.lower().str.contains(search_query, na=False)
        ]
    
    if len(master_grid) == 0:
        if search_query:
            st.warning(f"No items match your search for **'{search_query}'** in {active_category}.")
        else:
            st.info(f"No records under '{active_category}'. You can add items manually in Tab 2 or import an Excel file in Tab 3.")
    else:
        if "expanded_summary_groups" not in st.session_state:
            st.session_state.expanded_summary_groups = set()
            
        sum_weights = [0.4, 1.8, 2.2, 0.8, 0.8, 0.8, 0.8]
        for _ in custom_cols:
            sum_weights.append(0.8)
        sum_weights.extend([1.5, 1.2, 2.0])
        
        sum_cols_elements = st.columns(sum_weights)
        sum_cols_elements[0].markdown("")
        sum_cols_elements[1].markdown('<span class="grid-header-text">Group Name</span>', unsafe_allow_html=True)
        sum_cols_elements[2].markdown('<span class="grid-header-text">Item Name</span>', unsafe_allow_html=True)
        sum_cols_elements[3].markdown('<span class="grid-header-text">SOH</span>', unsafe_allow_html=True)
        sum_cols_elements[4].markdown('<span class="grid-header-text">Transit</span>', unsafe_allow_html=True)
        sum_cols_elements[5].markdown('<span class="grid-header-text">AMC</span>', unsafe_allow_html=True)
        sum_cols_elements[6].markdown('<span class="grid-header-text">Request</span>', unsafe_allow_html=True)
        
        col_idx = 7
        for c_name in custom_cols:
            sum_cols_elements[col_idx].markdown(f'<span class="grid-header-text">{c_name}</span>', unsafe_allow_html=True)
            col_idx += 1
            
        sum_cols_elements[col_idx].markdown('<span class="grid-header-text">Status</span>', unsafe_allow_html=True)
        sum_cols_elements[col_idx+1].markdown('<span class="grid-header-text">Owner</span>', unsafe_allow_html=True)
        sum_cols_elements[col_idx+2].markdown('<span class="grid-header-text">Last Note</span>', unsafe_allow_html=True)
        
        st.markdown("<hr style='margin:0px 0px 10px 0px; border-top: 2px solid #555;' />", unsafe_allow_html=True)
        
        unique_groups = master_grid["Drug Group"].unique()
        
        for group_name in unique_groups:
            group_df = master_grid[master_grid["Drug Group"] == group_name]
            is_repeating = len(group_df) > 1
            
            r_elements = st.columns(sum_weights)
            
            if is_repeating:
                sum_soh = int(group_df["SOH"].sum())
                sum_transit = int(group_df["Transit"].sum())
                sum_amc = int(group_df["AMC"].sum())
                sum_request = int(group_df["Request"].sum())
                
                statuses = group_df["Status"].unique()
                status_text = ", ".join(statuses) if len(statuses) <= 2 else f"{len(statuses)} Statuses"
                
                is_expanded = group_name in st.session_state.expanded_summary_groups
                btn_label = "➖" if is_expanded else "➕"
                
                if r_elements[0].button(btn_label, key=f"sum_exp_{group_name}", help="Click to expand/collapse details"):
                    if is_expanded:
                        st.session_state.expanded_summary_groups.remove(group_name)
                    else:
                        st.session_state.expanded_summary_groups.add(group_name)
                    st.rerun()
                
                r_elements[1].markdown(f"**{group_name}**")
                r_elements[2].markdown(f"*Multiple items ({len(group_df)} items)*")
                r_elements[3].markdown(f"**{sum_soh:,}**")
                r_elements[4].markdown(f"**{sum_transit:,}**")
                r_elements[5].markdown(f"**{sum_amc:,}**")
                r_elements[6].markdown(f"**{sum_request:,}**")
                
                c_idx = 7
                for _ in custom_cols:
                    r_elements[c_idx].markdown("-")
                    c_idx += 1
                    
                r_elements[c_idx].markdown(f"`{status_text}`")
                r_elements[c_idx+1].markdown("-")
                r_elements[c_idx+2].markdown("-")
                
                if is_expanded:
                    for idx, row in group_df.iterrows():
                        sub_elements = st.columns(sum_weights)
                        sub_elements[0].write("")
                        sub_elements[1].markdown(f"<span style='color: #777;'>{group_name}</span>", unsafe_allow_html=True)
                        sub_elements[2].write(f"• {row['Medicine']}")
                        sub_elements[3].write(f"{int(row['SOH']):,}")
                        sub_elements[4].write(f"{int(row['Transit']):,}")
                        sub_elements[5].write(f"{int(row['AMC']):,}")
                        sub_elements[6].write(f"{int(row['Request']):,}")
                        
                        c_dict = json.loads(row['CustomFields']) if row['CustomFields'] else {}
                        c_idx = 7
                        for c_name in custom_cols:
                            sub_elements[c_idx].write(str(c_dict.get(c_name, "-")))
                            c_idx += 1
                            
                        sub_elements[c_idx].markdown(render_status_badge(row['Status']), unsafe_allow_html=True)
                        sub_elements[c_idx+1].write(row['Owner'])
                        sub_elements[c_idx+2].caption(f"⏱️ ({row['Updated']}): {row['Note']}")
            else:
                row = group_df.iloc[0]
                r_elements[0].write("")
                r_elements[1].markdown(f"**{group_name}**")
                r_elements[2].write(row["Medicine"])
                r_elements[3].write(f"{int(row['SOH']):,}")
                r_elements[4].write(f"{int(row['Transit']):,}")
                r_elements[5].write(f"{int(row['AMC']):,}")
                r_elements[6].write(f"{int(row['Request']):,}")
                
                c_dict = json.loads(row['CustomFields']) if row['CustomFields'] else {}
                c_idx = 7
                for c_name in custom_cols:
                    r_elements[c_idx].write(str(c_dict.get(c_name, "-")))
                    c_idx += 1
                    
                r_elements[c_idx].markdown(render_status_badge(row['Status']), unsafe_allow_html=True)
                r_elements[c_idx+1].write(row["Owner"])
                r_elements[c_idx+2].caption(f"⏱️ ({row['Updated']}): {row['Note']}")
                
            st.markdown("<hr style='margin:2px 0px; border-top: 1px dashed #444;' />", unsafe_allow_html=True)

# ====================================================
# TAB 2: LIVE REVIEW & ITEM-BY-ITEM CREATION / DELETE
# ====================================================
with tabs[1]:
    if "flash_message" in st.session_state:
        st.success(st.session_state["flash_message"])
        del st.session_state["flash_message"]

    st.subheader(f"Interactive Board Grid — {active_category}")
    
    records = get_medicines_grid(active_category)
    
    if search_query and not records.empty:
        records = records[
            records["base_drug_name"].astype(str).str.lower().str.contains(search_query, na=False) |
            records["medicine_name"].astype(str).str.lower().str.contains(search_query, na=False) |
            records["owner"].astype(str).str.lower().str.contains(search_query, na=False) |
            records["status"].astype(str).str.lower().str.contains(search_query, na=False)
        ]
    
    st.caption(f"Add a new item row directly below under **{active_category}**, or modify/delete existing items.")
    
    grid_weights = get_board_column_weights(len(custom_cols))
    h_elements = st.columns(grid_weights)
    
    h_elements[0].markdown('<span class="grid-header-text">Select</span>', unsafe_allow_html=True)
    h_elements[1].markdown('<span class="grid-header-text">Item Description</span>', unsafe_allow_html=True)
    h_elements[2].markdown('<span class="grid-header-text">SOH</span>', unsafe_allow_html=True)
    h_elements[3].markdown('<span class="grid-header-text">Inbound</span>', unsafe_allow_html=True)
    h_elements[4].markdown('<span class="grid-header-text">AMC</span>', unsafe_allow_html=True)
    h_elements[5].markdown('<span class="grid-header-text">Request</span>', unsafe_allow_html=True)
    
    col_idx = 6
    for c_name in custom_cols:
        h_elements[col_idx].markdown(f'<span class="grid-header-text">{c_name}</span>', unsafe_allow_html=True)
        col_idx += 1
        
    h_elements[col_idx].markdown('<span class="grid-header-text">Status</span>', unsafe_allow_html=True)
    h_elements[col_idx+1].markdown('<span class="grid-header-text">Owner</span>', unsafe_allow_html=True)
    h_elements[col_idx+2].markdown('<span class="grid-header-text">Last Note</span>', unsafe_allow_html=True)
    h_elements[col_idx+3].markdown('<span class="grid-header-text">New Update</span>', unsafe_allow_html=True)
    h_elements[col_idx+4].markdown('<span class="grid-header-text">Save</span>', unsafe_allow_html=True)
    h_elements[col_idx+5].markdown('<span class="grid-header-text">Delete</span>', unsafe_allow_html=True)
    
    st.markdown("<hr style='margin:0px 0px 10px 0px; border-top: 2px solid #555;' />", unsafe_allow_html=True)
    
    # INLINE QUICK ADD ROW
    add_elements = st.columns(grid_weights)
    
    add_elements[0].markdown("➕")
    new_group = add_elements[1].text_input("Add Group", key="add_group", placeholder="GROUP / CLASS...", label_visibility="collapsed")
    new_med = add_elements[1].text_input("Add Item Name", key="add_med", placeholder="Item Description / Spec...", label_visibility="collapsed")
    
    new_soh_raw = add_elements[2].text_input("Add SOH", value="0", key="add_soh", label_visibility="collapsed")
    new_transit_raw = add_elements[3].text_input("Add Transit", value="0", key="add_transit", label_visibility="collapsed")
    new_amc_raw = add_elements[4].text_input("Add AMC", value="0", key="add_amc", label_visibility="collapsed")
    new_req_raw = add_elements[5].text_input("Add Request", value="0", key="add_req", label_visibility="collapsed")
    
    new_soh = parse_number(new_soh_raw)
    new_transit = parse_number(new_transit_raw)
    new_amc = parse_number(new_amc_raw)
    new_req = parse_number(new_req_raw)
    
    new_custom_inputs = {}
    col_idx = 6
    for c_name in custom_cols:
        new_custom_inputs[c_name] = add_elements[col_idx].text_input(f"Add {c_name}", key=f"add_custom_{c_name}", placeholder=f"{c_name}...", label_visibility="collapsed")
        col_idx += 1
        
    status_options = ["Open", "In Progress", "Waiting Supplier", "Escalated", "Completed"]
    new_status = add_elements[col_idx].selectbox("Add Status", status_options, index=0, key="add_status", label_visibility="collapsed")
    new_owner = add_elements[col_idx+1].text_input("Add Owner", placeholder="Owner...", key="add_owner", label_visibility="collapsed")
    
    add_elements[col_idx+2].markdown(f"<span style='color: #ff4b4b; font-weight: bold;'>New Entry</span>", unsafe_allow_html=True)
    new_comment = add_elements[col_idx+3].text_input("Add Comment", placeholder="Initial update note...", key="add_comment", label_visibility="collapsed")
    
    if add_elements[col_idx+4].button("➕", key="btn_add_new_row", help=f"Insert entry into {active_category}"):
        if not new_group.strip() or not new_med.strip():
            st.error("Both Group Name and Item Description are required!")
        else:
            success = add_new_medicine(new_group.strip().upper(), new_med.strip(), new_soh, new_transit, new_amc, new_req, category=active_category, custom_fields_dict=new_custom_inputs)
            if success:
                comment_to_save = new_comment.strip() if new_comment.strip() else "Item manually added."
                save_update(new_med.strip(), datetime.now().strftime("%Y-%m-%d %H:%M"), new_status, new_owner.strip(), comment_to_save)
                
                st.session_state["add_group"] = ""
                st.session_state["add_med"] = ""
                st.session_state["add_soh"] = "0"
                st.session_state["add_transit"] = "0"
                st.session_state["add_amc"] = "0"
                st.session_state["add_req"] = "0"
                st.session_state["add_owner"] = ""
                st.session_state["add_comment"] = ""
                for c_name in custom_cols:
                    if f"add_custom_{c_name}" in st.session_state:
                        st.session_state[f"add_custom_{c_name}"] = ""
                
                st.success(f"Added {new_med.strip()} to {active_category}!")
                st.rerun()
            else:
                st.error(f"Error: '{new_med.strip()}' already exists in your database.")

    st.markdown("<hr style='margin:10px 0px 15px 0px; border-top: 3px double #888;' />", unsafe_allow_html=True)

    # EXISTING RECORDS LOOP
    if len(records) == 0:
        if search_query:
            st.warning(f"No items match your search for **'{search_query}'** in {active_category}.")
        else:
            st.info(f"No items currently registered under '{active_category}'. Use the inline row above to add one.")
    else:
        for idx, row in records.iterrows():
            medicine = row["medicine_name"]
            group_label = row["base_drug_name"]
            
            # PRE-INSTANTIATION RESET CHECK:
            # Unchecks checkbox and clears cached input state BEFORE widgets render
            if st.session_state.get(f"reset_row_{medicine}"):
                st.session_state[f"chk_{medicine}"] = False
                for k in [f"soh_{medicine}", f"transit_{medicine}", f"amc_{medicine}", f"request_{medicine}", 
                          f"status_{medicine}", f"owner_{medicine}", f"comment_{medicine}",
                          f"adv_group_{medicine}", f"adv_name_{medicine}", f"adv_status_{medicine}", 
                          f"adv_owner_{medicine}", f"adv_soh_{medicine}", f"adv_transit_{medicine}", 
                          f"adv_amc_{medicine}", f"adv_req_{medicine}", f"adv_note_{medicine}"]:
                    if k in st.session_state:
                        del st.session_state[k]
                for c_name in custom_cols:
                    for ck in [f"custom_{c_name}_{medicine}", f"adv_custom_{c_name}_{medicine}"]:
                        if ck in st.session_state:
                            del st.session_state[ck]
                del st.session_state[f"reset_row_{medicine}"]

            # SPLIT ROW COLUMNS: Checkbox (Outside Form) | Inline Form Inputs | Delete Popover (Outside Form)
            chk_w = grid_weights[0]
            form_weights = grid_weights[1:-1]
            del_w = grid_weights[-1]
            
            col_chk, col_form, col_del = st.columns([chk_w, sum(form_weights), del_w])
            
            # 1. Checkbox outside form so clicking it toggles the expander immediately!
            with col_chk:
                is_selected = st.checkbox("", key=f"chk_{medicine}", label_visibility="collapsed")

            # 2. Row Form contains all inline inputs and the green save button
            with col_form:
                with st.form(key=f"row_form_{medicine}", border=False):
                    f_elements = st.columns(form_weights)
                    
                    f_elements[0].write(f"**{group_label}**\n\n{medicine}")
                    
                    soh_raw = f_elements[1].text_input("SOH", value=f"{int(row['stock_on_hand']):,}", key=f"soh_{medicine}", label_visibility="collapsed")
                    transit_raw = f_elements[2].text_input("Inbound", value=f"{int(row['in_transit']):,}", key=f"transit_{medicine}", label_visibility="collapsed")
                    amc_raw = f_elements[3].text_input("AMC", value=f"{int(row['amc']):,}", key=f"amc_{medicine}", label_visibility="collapsed")
                    request_qty_raw = f_elements[4].text_input("Request", value=f"{int(row['request_qty']):,}", key=f"request_{medicine}", label_visibility="collapsed")
                    
                    soh = parse_number(soh_raw)
                    transit = parse_number(transit_raw)
                    amc = parse_number(amc_raw)
                    request_qty = parse_number(request_qty_raw)
                    
                    existing_custom_dict = json.loads(row["custom_fields"]) if row["custom_fields"] else {}
                    updated_custom_dict = {}
                    
                    col_idx = 5
                    for c_name in custom_cols:
                        val = existing_custom_dict.get(c_name, "")
                        updated_custom_dict[c_name] = f_elements[col_idx].text_input(c_name, value=str(val), key=f"custom_{c_name}_{medicine}", label_visibility="collapsed")
                        col_idx += 1
                        
                    def_idx = status_options.index(row["status"]) if row["status"] in status_options else 0
                    selected_status = f_elements[col_idx].selectbox("Status", status_options, index=def_idx, key=f"status_{medicine}", label_visibility="collapsed")
                    current_owner = f_elements[col_idx+1].text_input("Owner", value=row["owner"], key=f"owner_{medicine}", label_visibility="collapsed")
                    
                    f_elements[col_idx+2].caption(f"⏱ *({row['last_updated']})*:\n{row['last_comment']}")
                    new_comment = f_elements[col_idx+3].text_input("New Comment", key=f"comment_{medicine}", label_visibility="collapsed", placeholder="Type update...")
                    
                    # Green Save Form Submit Button
                    f_elements[col_idx+4].markdown('<div class="save-btn-container">', unsafe_allow_html=True)
                    save_submitted = f_elements[col_idx+4].form_submit_button("💾", help=f"Save changes for {medicine}")
                    f_elements[col_idx+4].markdown('</div>', unsafe_allow_html=True)
                    
                    if save_submitted:
                        update_stock_levels(medicine, soh, transit, amc, request_qty, custom_fields_dict=updated_custom_dict)
                        
                        comment_to_save = new_comment.strip()
                        if not comment_to_save:
                            changes = []
                            if selected_status != row["status"]: changes.append(f"Status: {row['status']} ➔ {selected_status}")
                            if current_owner != row["owner"]: changes.append(f"Owner: {row['owner'] or 'None'} ➔ {current_owner}")
                            if soh != int(row["stock_on_hand"]): changes.append(f"SOH: {int(row['stock_on_hand']):,} ➔ {soh:,}")
                            if transit != int(row["in_transit"]): changes.append(f"Inbound: {int(row['in_transit']):,} ➔ {transit:,}")
                            comment_to_save = ", ".join(changes) if changes else "Updated parameters."
                        
                        save_update(medicine, datetime.now().strftime("%Y-%m-%d %H:%M"), selected_status, current_owner, comment_to_save)
                        st.session_state[f"reset_row_{medicine}"] = True
                        st.success(f"Saved & updated details for '{medicine}'!")
                        st.rerun()

            # 3. Red Delete Popover Button (Outside form)
            with col_del:
                col_del.markdown('<div class="delete-btn-container">', unsafe_allow_html=True)
                with col_del.popover("🗑️", help=f"Delete {medicine}"):
                    st.markdown(f"### Delete Item")
                    st.warning(f"Are you sure you want to delete **'{medicine}'**?")
                    deleter_name = st.text_input("Your Name (Required for audit log)", key=f"deleter_name_{medicine}").strip()
                    
                    if st.button("Confirm Delete", key=f"confirm_del_{medicine}", type="primary"):
                        if not deleter_name:
                            st.error("Please enter your name to confirm deletion.")
                        else:
                            delete_date_time = datetime.now().strftime("%Y-%m-%d %H:%M")
                            delete_medicine(medicine, deleter_name, delete_date_time)
                            st.success(f"Item '{medicine}' deleted by {deleter_name}.")
                            st.rerun()
                col_del.markdown('</div>', unsafe_allow_html=True)

            # ADVANCED DETAILS EDITOR FORM
            if is_selected:
                with st.expander(f"✏️ Advanced Details Editor — {medicine}", expanded=True):
                    with st.form(key=f"adv_form_{medicine}"):
                        st.markdown("#### Edit Standard Details")
                        
                        adv_c1, adv_c2, adv_c3, adv_c4 = st.columns(4)
                        adv_group = adv_c1.text_input("Group / Class Name", value=group_label, key=f"adv_group_{medicine}")
                        adv_name = adv_c2.text_input("Item Specification / Full Description", value=medicine, key=f"adv_name_{medicine}")
                        adv_status = adv_c3.selectbox("Current Status", status_options, index=def_idx, key=f"adv_status_{medicine}")
                        adv_owner = adv_c4.text_input("Responsible Officer / Owner", value=row["owner"], key=f"adv_owner_{medicine}")
                        
                        st.markdown("##### Stock & Quantities")
                        q_c1, q_c2, q_c3, q_c4 = st.columns(4)
                        adv_soh_raw = q_c1.text_input("Stock On Hand (SOH)", value=f"{int(row['stock_on_hand']):,}", key=f"adv_soh_{medicine}")
                        adv_transit_raw = q_c2.text_input("Inbound / In Transit", value=f"{int(row['in_transit']):,}", key=f"adv_transit_{medicine}")
                        adv_amc_raw = q_c3.text_input("Average Monthly Consumption (AMC)", value=f"{int(row['amc']):,}", key=f"adv_amc_{medicine}")
                        adv_req_raw = q_c4.text_input("Request Quantity", value=f"{int(row['request_qty']):,}", key=f"adv_req_{medicine}")
                        
                        adv_soh = parse_number(adv_soh_raw)
                        adv_transit = parse_number(adv_transit_raw)
                        adv_amc = parse_number(adv_amc_raw)
                        adv_req = parse_number(adv_req_raw)
                        
                        adv_custom_dict = {}
                        if custom_cols:
                            st.markdown("##### Dynamic Custom Attributes")
                            cust_cols_elements = st.columns(min(len(custom_cols), 4))
                            for idx_c, c_name in enumerate(custom_cols):
                                val_c = existing_custom_dict.get(c_name, "")
                                target_col = cust_cols_elements[idx_c % 4]
                                adv_custom_dict[c_name] = target_col.text_input(c_name, value=str(val_c), key=f"adv_custom_{c_name}_{medicine}")
                        else:
                            adv_custom_dict = updated_custom_dict

                        adv_note = st.text_area("Update Log Note", placeholder="Type detailed update note...", key="adv_note_" + medicine)
                        
                        adv_submitted = st.form_submit_button("💾 Save All Details", type="primary", help="Save all details")
                        
                        if adv_submitted:
                            update_stock_levels(medicine, adv_soh, adv_transit, adv_amc, adv_req, custom_fields_dict=adv_custom_dict)
                            comment_to_save = adv_note.strip() if adv_note.strip() else "Updated via Advanced Details Editor."
                            save_update(medicine, datetime.now().strftime("%Y-%m-%d %H:%M"), adv_status, adv_owner, comment_to_save)
                            
                            # Flag row reset on rerun to uncheck box (collapse window) & sync inline row textboxes
                            st.session_state[f"reset_row_{medicine}"] = True
                            st.success(f"Saved & updated inline view for '{medicine}'!")
                            st.rerun()

            with st.expander("📜 View Full Past Change Log Timeline", expanded=False):
                history_df = get_medicine_history(medicine)
                if len(history_df) <= 1:
                    st.caption("No older historical timeline records found.")
                else:
                    for _, log in history_df.iloc[1:].iterrows():
                        st.markdown(f"🗓️ **{log['update_date']}** | Status: `{log['status']}` | Owner: `{log['owner'] or '-'}`")
                        st.markdown(f"> *{log['comment']}*")
            
            st.markdown("<hr style='margin:5px 0px; border-top: 1px dashed #444;' />", unsafe_allow_html=True)

# ====================================================
# TAB 3: MASTER IMPORT PORTAL
# ====================================================
with tabs[2]:
    st.subheader(f"Seed Baseline Directory — {active_category}")
    st.write(f"Upload an Excel sheet formatted identically to seed or update records inside **{active_category}**.")
    
    uploaded_file = st.file_uploader(f"Upload Workbook for {active_category}", type=["xls", "xlsx"])

    if uploaded_file:
        try:
            sheets = load_excel_sheets(uploaded_file)
            selected_sheet = st.selectbox("Select Target Sheet Tab to Sync:", sheets)
            
            df = load_excel(uploaded_file, sheet_name=selected_sheet)
            df_prepared = prepare_medicines(df)
            
            df_prepared["category"] = active_category
            
            if st.button(f"🚀 Import Records into {active_category}", type="primary"):
                save_medicines(df_prepared)
                st.success(f"Successfully loaded records into '{active_category}'!")
                st.rerun()
            st.dataframe(df_prepared, use_container_width=True)
        except Exception as e:
            st.error(f"Excel Parser Error: {e}")