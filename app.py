import os
import sys
import json
import re
import io

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

# Render transient toast notification if saved
if "toast_msg" in st.session_state:
    st.toast(st.session_state["toast_msg"], icon="✅")
    del st.session_state["toast_msg"]

# Custom CSS styling
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
        /* Ensure selectbox options fit cleanly */
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
        
        /* Clean Status Badges */
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
        
        .adv-header-banner {
            background-color: #dc3545 !important;
            color: #ffffff !important;
            padding: 8px 14px !important;
            border-radius: 6px 6px 0px 0px !important;
            font-weight: bold !important;
            font-size: 1rem !important;
            margin-top: 8px !important;
            margin-bottom: 0px !important;
        }

        /* FAINT BACKGROUND BOX FOR LOG NOTES */
        .last-note-box {
            background-color: #f0f4f8 !important;
            border: 1px solid #d9e2ec !important;
            border-radius: 6px !important;
            padding: 6px 10px !important;
            font-size: 0.82rem !important;
            line-height: 1.35 !important;
            word-break: break-word !important;
            color: #334e68 !important;
        }

        /* TOP CONTROL BAR POPOVERS (Yellow Accent) */
        div.top-control-bar div[data-testid="stPopover"] button {
            background-color: #ffc107 !important;
            color: #212529 !important;
            border: 1px solid #e0a800 !important;
            font-weight: 600 !important;
            border-radius: 8px !important;
            padding: 6px 16px !important;
            white-space: nowrap !important;
            min-width: max-content !important;
            font-size: 0.92rem !important;
        }
        div.top-control-bar div[data-testid="stPopover"] button:hover {
            background-color: #e0a800 !important;
            color: #000000 !important;
            box-shadow: 0px 2px 6px rgba(0,0,0,0.15) !important;
        }
        div.top-control-bar div[data-testid="stPopover"] button svg {
            display: inline-block !important;
            margin-left: 6px !important;
        }

        /* FLEX CONTAINERS FOR FLUSH TOP ALIGNMENT */
        div.save-btn-container, div.delete-btn-container {
            display: flex !important;
            align-items: flex-start !important;
            margin-top: 0px !important;
            padding-top: 0px !important;
        }
        div.save-btn-container div[data-testid="stElementContainer"],
        div.delete-btn-container div[data-testid="stElementContainer"] {
            margin-top: 0px !important;
            padding-top: 0px !important;
        }
        
        /* HARD-OVERRIDE SOLID GREEN SAVE BUTTONS */
        div.save-btn-container button,
        div.save-btn-container button[data-testid="stBaseButton-secondary"],
        div.save-btn-container button[data-testid="stBaseButton-primary"] {
            background-color: #28a745 !important;
            background: #28a745 !important;
            color: #ffffff !important;
            border: none !important;
            font-weight: bold !important;
            border-radius: 8px !important;
            height: 40px !important;
            min-width: 48px !important;
            width: 100% !important;
            font-size: 1.2rem !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            margin-top: 0px !important;
        }
        div.save-btn-container button *,
        div.save-btn-container button p,
        div.save-btn-container button span {
            color: #ffffff !important;
        }
        div.save-btn-container button:hover {
            background-color: #218838 !important;
            background: #218838 !important;
            box-shadow: 0px 2px 6px rgba(40,167,69,0.3) !important;
        }

        /* HARD-OVERRIDE SOLID RED DELETE BUTTONS */
        div.delete-btn-container button,
        div.delete-btn-container button[data-testid="stBaseButton-secondary"],
        div.delete-btn-container button[data-testid="stBaseButton-primary"],
        div.delete-btn-container button[data-testid="stPopoverButton"] {
            background-color: #dc3545 !important;
            background: #dc3545 !important;
            color: #ffffff !important;
            border: none !important;
            font-weight: bold !important;
            border-radius: 8px !important;
            height: 40px !important;
            min-width: 48px !important;
            width: 100% !important;
            font-size: 1.2rem !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            padding: 0px !important;
            margin-top: 0px !important;
        }
        div.delete-btn-container button *,
        div.delete-btn-container button p,
        div.delete-btn-container button span {
            color: #ffffff !important;
        }
        div.delete-btn-container button:hover {
            background-color: #c82333 !important;
            background: #c82333 !important;
            box-shadow: 0px 2px 6px rgba(220,53,69,0.3) !important;
        }
        div.delete-btn-container button svg {
            display: none !important;
        }
    </style>
""", unsafe_allow_html=True)

# Connect and activate database tables
create_tables()
update_schema()

def parse_number(val):
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return int(val)
    clean_str = re.sub(r'[^\d-]', '', str(val))
    return int(clean_str) if clean_str else 0

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Inventory Export')
    return output.getvalue()

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

st.markdown('<div class="top-control-bar">', unsafe_allow_html=True)
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
            st.session_state["toast_msg"] = f"Category '{new_cat_input}' created. Open the Master Import Portal tab to download its template."
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
st.markdown('</div>', unsafe_allow_html=True)


# ====================================================
# QUICK ITEM SEARCH BAR & GLOBAL EXPORT BUTTON
# ====================================================
try:
    s_col0, s_col1, s_col2 = st.columns([2.7, 3.4, 1.3], vertical_alignment="center")
except TypeError:  # very old Streamlit without vertical_alignment
    s_col0, s_col1, s_col2 = st.columns([2.7, 3.4, 1.3])

with s_col0:
    st.title(f"📂 {active_category} Management")

with s_col1:
    search_query = st.text_input(
        "🔍 Search Item or Status",
        placeholder="🔎 Search item, group, owner or status (e.g. Escalated, Open)...",
        key="global_item_search",
        label_visibility="collapsed"
    ).strip().lower()

with s_col2:
    export_data_df = get_latest_statuses(active_category)
    if search_query and not export_data_df.empty:
        export_data_df = export_data_df[
            export_data_df["Drug Group"].astype(str).str.lower().str.contains(search_query, na=False) |
            export_data_df["Medicine"].astype(str).str.lower().str.contains(search_query, na=False) |
            export_data_df["Owner"].astype(str).str.lower().str.contains(search_query, na=False) |
            export_data_df["Status"].astype(str).str.lower().str.contains(search_query, na=False)
        ]

    if not export_data_df.empty:
        excel_bytes = convert_df_to_excel(export_data_df)
        st.download_button(
            label="📊 Export Excel",
            data=excel_bytes,
            file_name=f"{active_category}_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            key="global_excel_export_btn"
        )

st.markdown("""
    <style>
        /* ---- Compact top section (less empty space, nothing hidden) ---- */
        header[data-testid="stHeader"] { height: 2.75rem !important; }
        .block-container { padding-top: 3.1rem !important; padding-bottom: 1.5rem !important; }
        .block-container > div[data-testid="stVerticalBlock"] { gap: 0.55rem !important; }
        .block-container h1 {
            font-size: 1.75rem !important;
            line-height: 1.25 !important;
            padding: 0 !important;
            margin: 0 !important;
        }
        /* style-only helper blocks should not take up a blank row */
        div[data-testid="stElementContainer"]:has(> div[data-testid="stMarkdownContainer"] > style),
        .element-container:has(> div[data-testid="stMarkdownContainer"] > style) { display: none !important; }
        div[data-baseweb="tab-panel"] { padding-top: 0.4rem !important; }
        .kpi-card { padding: 9px 8px !important; margin-bottom: 2px !important; }
        .kpi-title { margin-bottom: 2px !important; }
        .kpi-value { font-size: 1.55rem !important; }
        div[data-testid="stMarkdownContainer"] > hr:not([style]) { margin: 0.5rem 0 !important; }

        /* Expand / collapse (+ / -) buttons in the Master Meeting Tracking Log */
        [class*="st-key-sum_exp_closed_"] button,
        [class*="st-key-sum_exp_open_"] button {
            width: 30px !important;
            min-width: 30px !important;
            height: 30px !important;
            padding: 0 !important;
            border-radius: 8px !important;
            box-shadow: 0px 1px 3px rgba(0,0,0,0.25) !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
        }
        [class*="st-key-sum_exp_closed_"] button {
            background-color: #0d6efd !important;
            border: 1px solid #0b5ed7 !important;
        }
        [class*="st-key-sum_exp_open_"] button {
            background-color: #495057 !important;
            border: 1px solid #343a40 !important;
        }
        [class*="st-key-sum_exp_closed_"] button:hover,
        [class*="st-key-sum_exp_open_"] button:hover {
            filter: brightness(1.15);
            border-color: #ffffff !important;
        }
        [class*="st-key-sum_exp_closed_"] button *,
        [class*="st-key-sum_exp_open_"] button * {
            color: #ffffff !important;
            font-size: 1.35rem !important;
            font-weight: 800 !important;
            line-height: 1 !important;
            margin: 0 !important;
        }
    </style>
""", unsafe_allow_html=True)

tabs = st.tabs([
    "📊 Summary Action Dashboard",
    "🗓️ Live Meeting Review",
    "📥 Master Import Portal"
])

def get_board_column_weights(custom_cols_count):
    weights = [0.4, 2.2, 1.1, 1.1, 1.1, 1.1]
    for _ in range(custom_cols_count):
        weights.append(1.0)
    weights.extend([1.4, 1.2, 1.8, 1.8, 0.9, 0.9])
    return weights

# ====================================================
# TAB 1: ACTION DASHBOARD
# ====================================================
with tabs[0]:
    t1_head1, t1_head2 = st.columns([3.5, 1.2])
    with t1_head1:
        st.subheader(f"Meeting Action Progress KPIs — {active_category}")
    with t1_head2:
        if not export_data_df.empty:
            st.download_button(
                label="📊 Export Dashboard Data",
                data=excel_bytes,
                file_name=f"{active_category}_Dashboard_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="tab1_excel_export_btn",
                use_container_width=True
            )

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
    st.caption("Groups with several items have a blue **+** button on the left. Click it to show the individual items, and click **−** to hide them again.")
    
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
            
        sum_weights = [0.6, 1.8, 2.2, 0.8, 0.8, 0.8, 0.8]
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
                btn_label = "−" if is_expanded else "+"
                
                if r_elements[0].button(
                    btn_label,
                    key=f"sum_exp_{'open' if is_expanded else 'closed'}_{group_name}",
                    help="Click to hide the individual items" if is_expanded else "Click to show the individual items in this group",
                ):
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
                    r_elements[c_idx].markdown("–")
                    c_idx += 1
                    
                r_elements[c_idx].markdown(f"`{status_text}`")
                r_elements[c_idx+1].markdown("–")
                r_elements[c_idx+2].markdown("–")
                
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
                        sub_elements[c_idx+1].write(row['Owner'] if row['Owner'] not in ('-', '') else '–')
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
                r_elements[c_idx+1].write(row["Owner"] if row["Owner"] not in ("-", "") else "–")
                r_elements[c_idx+2].caption(f"⏱️ ({row['Updated']}): {row['Note']}")
                
            st.markdown("<hr style='margin:2px 0px; border-top: 1px dashed #444;' />", unsafe_allow_html=True)

# ====================================================
# TAB 2 ISOLATED FRAGMENT FOR LIGHTNING-FAST GRID INTERACTION
# ====================================================
@st.fragment
def render_review_board_grid(records, custom_cols, active_category):
    status_options = ["Open", "In Progress", "Waiting Supplier", "Escalated", "Completed"]
    grid_weights = get_board_column_weights(len(custom_cols))
    
    # Header row
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

    # Inline Quick Add Row
    add_elements = st.columns(grid_weights)
    add_elements[0].markdown("➕")
    new_group = add_elements[1].text_input("Add Group", key="add_group", placeholder="GROUP / CLASS...", label_visibility="collapsed")
    new_med = add_elements[1].text_input("Add Item Name", key="add_med", placeholder="Item Description / Spec...", label_visibility="collapsed")
    
    new_soh_raw = add_elements[2].text_input("Add SOH", value="0", key="add_soh", label_visibility="collapsed")
    new_transit_raw = add_elements[3].text_input("Add Transit", value="0", key="add_transit", label_visibility="collapsed")
    new_amc_raw = add_elements[4].text_input("Add AMC", value="0", key="add_amc", label_visibility="collapsed")
    new_req_raw = add_elements[5].text_input("Add Request", value="0", key="add_req", label_visibility="collapsed")
    
    new_custom_inputs = {}
    col_idx = 6
    for c_name in custom_cols:
        new_custom_inputs[c_name] = add_elements[col_idx].text_input(f"Add {c_name}", key=f"add_custom_{c_name}", placeholder=f"{c_name}...", label_visibility="collapsed")
        col_idx += 1
        
    new_status = add_elements[col_idx].selectbox("Add Status", status_options, index=0, key="add_status", label_visibility="collapsed")
    new_owner = add_elements[col_idx+1].text_input("Add Owner", placeholder="Owner...", key="add_owner", label_visibility="collapsed")
    
    add_elements[col_idx+2].markdown(f"<span style='color: #ff4b4b; font-weight: bold;'>New Entry</span>", unsafe_allow_html=True)
    new_comment = add_elements[col_idx+3].text_input("Add Comment", placeholder="Initial update note...", key="add_comment", label_visibility="collapsed")
    
    if add_elements[col_idx+4].button("➕", key="btn_add_new_row", help=f"Insert entry into {active_category}"):
        if not new_group.strip() or not new_med.strip():
            st.error("Both Group Name and Item Description are required!")
        else:
            success = add_new_medicine(
                new_group.strip().upper(), new_med.strip(), 
                parse_number(new_soh_raw), parse_number(new_transit_raw), 
                parse_number(new_amc_raw), parse_number(new_req_raw), 
                category=active_category, custom_fields_dict=new_custom_inputs
            )
            if success:
                comment_to_save = new_comment.strip() if new_comment.strip() else "Item manually added."
                save_update(new_med.strip(), datetime.now().strftime("%Y-%m-%d %H:%M"), new_status, new_owner.strip(), comment_to_save)
                
                st.session_state["toast_msg"] = "All Data is saved"
                st.rerun()
            else:
                st.error(f"Error: '{new_med.strip()}' already exists in your database.")

    st.markdown("<hr style='margin:10px 0px 15px 0px; border-top: 3px double #888;' />", unsafe_allow_html=True)

    chk_w = grid_weights[0]
    form_weights = grid_weights[1:-1]
    del_w = grid_weights[-1]

    # Render items
    for idx, row in records.iterrows():
        medicine = row["medicine_name"]
        group_label = row["base_drug_name"]

        # SAFE PRE-INSTANTIATION STATE RESETS BEFORE WIDGETS ARE RENDERED
        if st.session_state.get(f"reset_chk_{medicine}"):
            st.session_state[f"chk_{medicine}"] = False
            del st.session_state[f"reset_chk_{medicine}"]

        if st.session_state.get(f"reset_comment_{medicine}"):
            st.session_state[f"comment_{medicine}"] = ""
            del st.session_state[f"reset_comment_{medicine}"]

        if st.session_state.get(f"reset_adv_note_{medicine}"):
            st.session_state[f"adv_note_{medicine}"] = ""
            del st.session_state[f"reset_adv_note_{medicine}"]

        col_chk, col_form, col_del = st.columns([chk_w, sum(form_weights), del_w])
        
        # 1. Instant Checkbox
        with col_chk:
            is_selected = st.checkbox("", key=f"chk_{medicine}", label_visibility="collapsed")

        # 2. Main Row Form
        with col_form:
            with st.form(key=f"row_form_{medicine}", border=False):
                f_elements = st.columns(form_weights)
                
                f_elements[0].write(f"**{group_label}**\n\n{medicine}")
                
                soh_raw = f_elements[1].text_input("SOH", value=f"{int(row['stock_on_hand']):,}", key=f"soh_{medicine}", label_visibility="collapsed")
                transit_raw = f_elements[2].text_input("Inbound", value=f"{int(row['in_transit']):,}", key=f"transit_{medicine}", label_visibility="collapsed")
                amc_raw = f_elements[3].text_input("AMC", value=f"{int(row['amc']):,}", key=f"amc_{medicine}", label_visibility="collapsed")
                request_qty_raw = f_elements[4].text_input("Request", value=f"{int(row['request_qty']):,}", key=f"request_{medicine}", label_visibility="collapsed")
                
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
                
                # STYLED LAST NOTE BOX (Faint background color for visual prominence)
                last_note_html = f"""
                <div class="last-note-box">
                    <span style="font-size: 0.72rem; color: #555; font-weight: bold;">⏱ ({row['last_updated']})</span><br/>
                    <span style="font-size: 0.82rem; color: #111;">{row['last_comment'] or 'No notes'}</span>
                </div>
                """
                f_elements[col_idx+2].markdown(last_note_html, unsafe_allow_html=True)
                
                new_comment = f_elements[col_idx+3].text_input("New Comment", key=f"comment_{medicine}", label_visibility="collapsed", placeholder="Type update...")
                
                # Solid Green Save Button Container (Aligned top)
                f_elements[col_idx+4].markdown('<div class="save-btn-container">', unsafe_allow_html=True)
                save_submitted = f_elements[col_idx+4].form_submit_button("💾", help=f"Save changes for {medicine}")
                f_elements[col_idx+4].markdown('</div>', unsafe_allow_html=True)
                
                if save_submitted:
                    comment_to_save = new_comment.strip()
                    # STRICT GUARD: Block saving if comment box is empty
                    if not comment_to_save:
                        st.error("⚠️ Please enter an update comment before saving!")
                    else:
                        soh = parse_number(soh_raw)
                        transit = parse_number(transit_raw)
                        amc = parse_number(amc_raw)
                        request_qty = parse_number(request_qty_raw)

                        update_stock_levels(medicine, soh, transit, amc, request_qty, custom_fields_dict=updated_custom_dict)
                        save_update(medicine, datetime.now().strftime("%Y-%m-%d %H:%M"), selected_status, current_owner, comment_to_save)
                        
                        # QUEUE TEXTBOX CLEAR BEFORE RERUN
                        st.session_state[f"reset_comment_{medicine}"] = True
                        st.session_state["toast_msg"] = "All Data is saved"
                        st.rerun()

        # 3. Solid Red Delete Popover Container (Aligned top)
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
                        st.session_state["toast_msg"] = f"Item '{medicine}' deleted"
                        st.rerun()
            col_del.markdown('</div>', unsafe_allow_html=True)

        # Instant Advanced Editor Window
        if is_selected:
            st.markdown(f'<div class="adv-header-banner">✏️ Advanced Details Editor — {medicine}</div>', unsafe_allow_html=True)
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
                    adv_soh = parse_number(adv_soh_raw)
                    adv_transit = parse_number(adv_transit_raw)
                    adv_amc = parse_number(adv_amc_raw)
                    adv_req = parse_number(adv_req_raw)

                    update_stock_levels(medicine, adv_soh, adv_transit, adv_amc, adv_req, custom_fields_dict=adv_custom_dict)
                    comment_to_save = adv_note.strip() if adv_note.strip() else "Updated via Advanced Details Editor."
                    save_update(medicine, datetime.now().strftime("%Y-%m-%d %H:%M"), adv_status, adv_owner, comment_to_save)
                    
                    # QUEUE RESETS BEFORE RERUN
                    st.session_state[f"reset_comment_{medicine}"] = True
                    st.session_state[f"reset_adv_note_{medicine}"] = True
                    st.session_state[f"reset_chk_{medicine}"] = True
                    st.session_state["toast_msg"] = "All Data is saved"
                    st.rerun()

        # History timeline expander
        with st.expander("📜 View Full Past Change Log Timeline", expanded=False):
            if is_selected:
                history_df = get_medicine_history(medicine)
                if len(history_df) <= 1:
                    st.caption("No older historical timeline records found.")
                else:
                    for _, log in history_df.iloc[1:].iterrows():
                        st.markdown(f"🗓️ **{log['update_date']}** | Status: `{log['status']}` | Owner: `{log['owner'] or '-'}`")
                        st.markdown(f"> *{log['comment']}*")
            else:
                st.caption("Select checkbox to view complete change history.")
        
        st.markdown("<hr style='margin:5px 0px; border-top: 1px dashed #444;' />", unsafe_allow_html=True)

# ====================================================
# TAB 2: LIVE REVIEW & ITEM-BY-ITEM CREATION / DELETE
# ====================================================
with tabs[1]:
    t2_head1, t2_head2 = st.columns([3.5, 1.2])
    with t2_head1:
        st.subheader(f"Interactive Board Grid — {active_category}")
    with t2_head2:
        if not export_data_df.empty:
            st.download_button(
                label="📊 Export Grid Data",
                data=excel_bytes,
                file_name=f"{active_category}_Grid_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="tab2_excel_export_btn",
                use_container_width=True
            )

    records = get_medicines_grid(active_category)
    
    if search_query and not records.empty:
        records = records[
            records["base_drug_name"].astype(str).str.lower().str.contains(search_query, na=False) |
            records["medicine_name"].astype(str).str.lower().str.contains(search_query, na=False) |
            records["owner"].astype(str).str.lower().str.contains(search_query, na=False) |
            records["status"].astype(str).str.lower().str.contains(search_query, na=False)
        ]
    
    st.caption(f"Add a new item row directly below under **{active_category}**, or modify/delete existing items.")

    # ---- Bulk delete / clean-up (collapsed by default so the grid layout is unchanged) ----
    with st.expander("🧹 Bulk Delete — remove several items at once", expanded=False):
        bulk_all = get_medicines_grid(active_category)
        sel_key = f"bulk_del_sel_{active_category}"
        name_key = f"bulk_del_name_{active_category}"
        reason_key = f"bulk_del_reason_{active_category}"
        confirm_key = f"bulk_del_confirm_{active_category}"

        # Clear the widgets safely BEFORE they are drawn (after a successful delete)
        if st.session_state.pop(f"reset_bulk_del_{active_category}", False):
            for _k in (sel_key, name_key, reason_key, confirm_key):
                st.session_state.pop(_k, None)

        if bulk_all.empty:
            st.info(f"No items in '{active_category}' to delete.")
        else:
            bulk_labels = {
                r["medicine_name"]: f"{r['medicine_name']}   [{r['status']} | last update: {r['last_updated']}]"
                for _, r in bulk_all.iterrows()
            }
            previously_deleted = bulk_all.loc[bulk_all["status"] == "DELETED", "medicine_name"].tolist()

            def _pick_previously_deleted(key, names):
                st.session_state[key] = names

            st.caption("Pick the items that should not be here, enter your name for the audit log, tick the confirmation and delete. "
                       "Every deleted item gets its own audit entry.")
            st.button(
                f"⚡ Select the {len(previously_deleted)} item(s) whose last status is DELETED (previously deleted, came back)",
                key=f"bulk_pick_deleted_{active_category}",
                on_click=_pick_previously_deleted,
                args=(sel_key, previously_deleted),
                disabled=not previously_deleted,
            )
            selected_to_delete = st.multiselect(
                "Items to delete",
                options=list(bulk_labels.keys()),
                format_func=lambda n: bulk_labels.get(n, n),
                key=sel_key,
                placeholder="Search and pick one or more items...",
            )

            if selected_to_delete:
                preview_df = bulk_all[bulk_all["medicine_name"].isin(selected_to_delete)][
                    ["base_drug_name", "medicine_name", "stock_on_hand", "status", "last_updated"]
                ].rename(columns={"base_drug_name": "Group", "medicine_name": "Item", "stock_on_hand": "SOH",
                                  "status": "Status", "last_updated": "Last updated"})
                st.dataframe(preview_df, hide_index=True, use_container_width=True)
                if len(selected_to_delete) == len(bulk_all):
                    st.warning(f"This selects EVERY item in '{active_category}'.")

            bd_c1, bd_c2 = st.columns(2)
            bulk_name = bd_c1.text_input("Your Name (Required for audit log)", key=name_key).strip()
            bulk_reason = bd_c2.text_input("Reason (optional)", key=reason_key,
                                           placeholder="e.g. Added by mistake during import")
            bulk_confirm = st.checkbox(
                f"I confirm I want to permanently delete {len(selected_to_delete)} item(s) from '{active_category}'",
                key=confirm_key,
            )
            if st.button(f"🗑️ Delete {len(selected_to_delete)} selected item(s)", type="primary",
                         key=f"bulk_del_btn_{active_category}"):
                if not selected_to_delete:
                    st.error("Please select at least one item.")
                elif not bulk_name:
                    st.error("Please enter your name to confirm deletion.")
                elif not bulk_confirm:
                    st.error("Please tick the confirmation box.")
                else:
                    done = delete_medicines_bulk(
                        selected_to_delete, bulk_name, datetime.now().strftime("%Y-%m-%d %H:%M"),
                        category=active_category, reason=bulk_reason)
                    st.session_state[f"reset_bulk_del_{active_category}"] = True
                    st.session_state["toast_msg"] = f"{len(done)} item(s) deleted from '{active_category}' by {bulk_name}"
                    st.rerun()
    
    if len(records) == 0:
        if search_query:
            st.warning(f"No items match your search for **'{search_query}'** in {active_category}.")
        else:
            st.info(f"No items currently registered under '{active_category}'. Use the inline row above to add one.")
    else:
        render_review_board_grid(records, custom_cols, active_category)

# ====================================================
# TAB 3: MASTER IMPORT PORTAL
# ====================================================
def render_import_preview(df_prepared, target_category, cat_custom_cols, key_prefix):
    """Shows exactly what an import will do, then imports safely (never re-labels existing items)."""
    if df_prepared is None or df_prepared.empty:
        st.warning("No importable item rows were found. Make sure the data has an 'ITEM DESCRIPTION' column (see the template above).")
        return

    df_prepared = df_prepared.copy()
    df_prepared["category"] = target_category
    df_preview = classify_import_rows(df_prepared, target_category)

    n_new = int((df_preview["import_status"] == STATUS_NEW).sum())
    n_update = int((df_preview["import_status"] == STATUS_UPDATE).sum())
    n_blocked = int(df_preview["import_status"].str.startswith(STATUS_BLOCKED_PREFIX).sum())

    m1, m2, m3 = st.columns(3)
    m1.metric("New items to add", n_new)
    m2.metric(f"Existing in {target_category} (update)", n_update)
    m3.metric("Skipped (belong to another category)", n_blocked)

    if n_blocked:
        st.warning(
            f"{n_blocked} row(s) already exist in another category (for example Medicines). "
            f"They will NOT be changed or re-labelled as '{target_category}'. "
            "If most rows are skipped, you probably selected the wrong sheet."
        )

    if st.button(f"🚀 Import Records into {target_category}", type="primary",
                 key=f"{key_prefix}_import_btn", disabled=(n_new + n_update) == 0):
        result = save_medicines(df_prepared, category=target_category, custom_cols=cat_custom_cols)
        st.session_state["toast_msg"] = (
            f"'{target_category}': added {result['added']}, updated {result['updated']}, "
            f"skipped {len(result['skipped'])} (belong to other categories, left unchanged)."
        )
        st.rerun()
    st.dataframe(df_preview, use_container_width=True)


with tabs[2]:
    st.subheader(f"Seed Baseline Directory — {active_category}")
    st.write(f"Upload an Excel sheet formatted identically to seed or update records inside **{active_category}**.")

    # ---- Template / column order -------------------------------------------------
    template_cols = get_template_columns(custom_cols)
    st.info("**Column order:** " + "  →  ".join(template_cols) + "  \n Only **ITEM DESCRIPTION** is required; the rest are optional.")

    t_col1, t_col2 = st.columns([1, 3])
    t_col1.download_button(
        "⬇️ Download Excel Template",
        data=build_template_excel(active_category, custom_cols),
        file_name=f"{active_category.replace(' ', '_')}_import_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_template_{active_category}",
    )
    with t_col2.expander("📋 See what each column means"):
        st.dataframe(get_template_guide(custom_cols), hide_index=True, use_container_width=True)
        st.caption("Item names are imported exactly as written. Items that already exist in a different category are never changed.")

    # ---- Input method ------------------------------------------------------------
    input_mode = st.radio(
        "How would you like to add items?",
        ["📁 Upload Excel", "📋 Paste from Excel", "⌨️ Type in a table"],
        horizontal=True,
        key=f"import_mode_{active_category}",
    )

    if input_mode == "📁 Upload Excel":
        uploaded_file = st.file_uploader(f"Upload Workbook for {active_category}", type=["xls", "xlsx"])

        if uploaded_file:
            try:
                file_bytes = uploaded_file.getvalue()
                sheets = load_excel_sheets(io.BytesIO(file_bytes))

                # Work out how many genuinely NEW items each sheet would add, so the right sheet is pre-selected
                sheet_new_counts = {}
                for sh in sheets:
                    try:
                        prep = prepare_medicines(load_excel(io.BytesIO(file_bytes), sheet_name=sh), custom_cols)
                        if prep.empty:
                            sheet_new_counts[sh] = 0
                        else:
                            prep["category"] = active_category
                            classified = classify_import_rows(prep, active_category)
                            sheet_new_counts[sh] = int((classified["import_status"] == STATUS_NEW).sum())
                    except Exception:
                        sheet_new_counts[sh] = 0
                best_sheet = max(sheets, key=lambda s: sheet_new_counts[s]) if any(sheet_new_counts.values()) else sheets[0]

                selected_sheet = st.selectbox(
                    "Select Target Sheet Tab to Sync:",
                    sheets,
                    index=sheets.index(best_sheet),
                    format_func=lambda s: f"{s}  ({sheet_new_counts.get(s, 0)} new item(s))",
                    key=f"sheet_sel_{active_category}_{uploaded_file.name}_{uploaded_file.size}",
                )

                df = load_excel(io.BytesIO(file_bytes), sheet_name=selected_sheet)
                df_prepared = prepare_medicines(df, custom_cols)
                render_import_preview(df_prepared, active_category, custom_cols, "upload")
            except Exception as e:
                st.error(f"Excel Parser Error: {e}")

    elif input_mode == "📋 Paste from Excel":
        st.caption("Copy the cells in Excel (include the header row in the column order above) and paste them here. "
                   "A plain list with one item name per line also works.")
        pasted = st.text_area(
            "Paste data here",
            height=220,
            placeholder="\t".join(template_cols),
            key=f"paste_area_{active_category}",
        )
        if pasted.strip():
            try:
                df_prepared = prepare_medicines(parse_pasted_text(pasted), custom_cols)
                render_import_preview(df_prepared, active_category, custom_cols, "paste")
            except Exception as e:
                st.error(f"Could not read the pasted data: {e}")

    else:
        st.caption("Type or paste directly into the table. Use the + row at the bottom to add more lines.")
        edited_df = st.data_editor(
            build_blank_entry_df(custom_cols),
            num_rows="dynamic",
            hide_index=True,
            use_container_width=True,
            key=f"typed_entry_{active_category}_{'|'.join(template_cols)}",
        )
        try:
            df_prepared = prepare_medicines(edited_df, custom_cols)
            if not df_prepared.empty:
                render_import_preview(df_prepared, active_category, custom_cols, "typed")
        except Exception as e:
            st.error(f"Could not read the table: {e}")
