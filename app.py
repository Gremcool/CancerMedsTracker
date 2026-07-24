import os
import sys
import json

# Ensure local imports resolve safely
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import pandas as pd
from datetime import datetime, date

from database import *
from excel_import import *

# Configure page and hide sidebar by default
st.set_page_config(
    page_title="Medical Inventory System",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Inject custom CSS to lower top controls below Streamlit header banner and tighten spacing
st.markdown("""
    <style>
        /* Increase top padding to clear Streamlit's header banner */
        .block-container {
            padding-top: 3.2rem !important;
            padding-bottom: 0rem !important;
            margin-top: 0rem !important;
        }
        /* Completely hide sidebar collapse button/chevron */
        [data-testid="collapsedControl"] {
            display: none !important;
        }
        /* Highlight input borders */
        div[data-testid="stInput"] input {
            border: 1px solid #ff4b4b !important;
        }
        /* Pull page title closer to top controls */
        div[data-testid="stTitle"] {
            margin-top: -0.6rem !important;
            padding-top: 0rem !important;
        }
    </style>
""", unsafe_allow_html=True)

# Connect and activate schemas
create_tables()
update_schema()

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
            st.session_state["input_new_category"] = ""  # Clear input
            st.rerun()  # Triggers rerun and automatically closes popover
        else:
            st.error("Please enter a category name.")

# Custom Columns Manager Popover
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
            st.session_state["input_new_column"] = ""  # Clear input
            st.rerun()  # Triggers rerun and automatically closes popover

st.title(f"📂 {active_category} Management")

tabs = st.tabs([
    "📊 Summary Action Dashboard",
    "🗓️ Live Meeting Review",
    "📥 Master Import Portal"
])

# Helper function to compute dynamic column layout weights
def get_column_weights(custom_cols_count):
    weights = [1.8, 0.5, 0.9, 0.5, 0.8]
    for _ in range(custom_cols_count):
        weights.append(0.8)
    weights.extend([1.0, 1.0, 1.4, 1.8, 0.8])
    return weights

# ====================================================
# TAB 1: ACTION DASHBOARD
# ====================================================
with tabs[0]:
    st.subheader(f"Meeting Action Progress KPIs — {active_category}")
    stats = get_dashboard_stats(active_category)

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total Items", f"{stats['total']}")
    c2.metric("Open Actions", f"{stats['open']}")
    c3.metric("In Progress", f"{stats['progress']}")
    c4.metric("Waiting Supplier", f"{stats['supplier']}")
    c5.metric("Escalated 🔥", f"{stats['escalated']}")
    c6.metric("Completed ✅", f"{stats['completed']}")

    st.markdown("---")
    st.write(f"### 📋 Master Meeting Tracking Log — {active_category}")
    
    master_grid = get_latest_statuses(active_category)
    if len(master_grid) == 0:
        st.info(f"No records under '{active_category}'. You can add items manually in Tab 2 or import an Excel file in Tab 3.")
    else:
        if "expanded_summary_groups" not in st.session_state:
            st.session_state.expanded_summary_groups = set()
            
        sum_weights = [0.4, 1.8, 2.2, 0.6, 0.6, 0.6, 0.6]
        for _ in custom_cols:
            sum_weights.append(0.8)
        sum_weights.extend([1.0, 1.0, 2.0])
        
        sum_cols_elements = st.columns(sum_weights)
        sum_cols_elements[0].markdown("")
        sum_cols_elements[1].markdown("**Group Name**")
        sum_cols_elements[2].markdown("**Item Name**")
        sum_cols_elements[3].markdown("**SOH**")
        sum_cols_elements[4].markdown("**Transit**")
        sum_cols_elements[5].markdown("**AMC**")
        sum_cols_elements[6].markdown("**Request**")
        
        col_idx = 7
        for c_name in custom_cols:
            sum_cols_elements[col_idx].markdown(f"**{c_name}**")
            col_idx += 1
            
        sum_cols_elements[col_idx].markdown("**Status**")
        sum_cols_elements[col_idx+1].markdown("**Owner**")
        sum_cols_elements[col_idx+2].markdown("**Last Note**")
        
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
                r_elements[3].markdown(f"**{sum_soh}**")
                r_elements[4].markdown(f"**{sum_transit}**")
                r_elements[5].markdown(f"**{sum_amc}**")
                r_elements[6].markdown(f"**{sum_request}**")
                
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
                        sub_elements[3].write(str(row['SOH']))
                        sub_elements[4].write(str(row['Transit']))
                        sub_elements[5].write(str(row['AMC']))
                        sub_elements[6].write(str(row['Request']))
                        
                        c_dict = json.loads(row['CustomFields']) if row['CustomFields'] else {}
                        c_idx = 7
                        for c_name in custom_cols:
                            sub_elements[c_idx].write(str(c_dict.get(c_name, "-")))
                            c_idx += 1
                            
                        sub_elements[c_idx].write(row['Status'])
                        sub_elements[c_idx+1].write(row['Owner'])
                        sub_elements[c_idx+2].caption(f"⏱️ ({row['Updated']}): {row['Note']}")
            else:
                row = group_df.iloc[0]
                r_elements[0].write("")
                r_elements[1].markdown(f"**{group_name}**")
                r_elements[2].write(row["Medicine"])
                r_elements[3].write(str(row["SOH"]))
                r_elements[4].write(str(row["Transit"]))
                r_elements[5].write(str(row["AMC"]))
                r_elements[6].write(str(row["Request"]))
                
                c_dict = json.loads(row['CustomFields']) if row['CustomFields'] else {}
                c_idx = 7
                for c_name in custom_cols:
                    r_elements[c_idx].write(str(c_dict.get(c_name, "-")))
                    c_idx += 1
                    
                r_elements[c_idx].write(row["Status"])
                r_elements[c_idx+1].write(row["Owner"])
                r_elements[c_idx+2].caption(f"⏱️ ({row['Updated']}): {row['Note']}")
                
            st.markdown("<hr style='margin:2px 0px; border-top: 1px dashed #444;' />", unsafe_allow_html=True)

# ====================================================
# TAB 2: LIVE REVIEW & ITEM-BY-ITEM CREATION / DELETE
# ====================================================
with tabs[1]:
    st.subheader(f"Interactive Board Grid — {active_category}")
    
    records = get_medicines_grid(active_category)
    
    st.caption(f"Add a new item row directly below under **{active_category}**, or modify/delete existing items.")
    
    grid_weights = get_column_weights(len(custom_cols))
    h_elements = st.columns(grid_weights)
    
    h_elements[0].markdown("**Item Description**")
    h_elements[1].markdown("**SOH**")
    h_elements[2].markdown("**Inbound**")
    h_elements[3].markdown("**AMC**")
    h_elements[4].markdown("**Request**")
    
    col_idx = 5
    for c_name in custom_cols:
        h_elements[col_idx].markdown(f"**{c_name}**")
        col_idx += 1
        
    h_elements[col_idx].markdown("**Status**")
    h_elements[col_idx+1].markdown("**Owner**")
    h_elements[col_idx+2].markdown("**Last Note**")
    h_elements[col_idx+3].markdown("**New Update**")
    h_elements[col_idx+4].markdown("**Action**")
    
    st.markdown("<hr style='margin:0px 0px 10px 0px; border-top: 2px solid #555;' />", unsafe_allow_html=True)
    
    # ====================================================
    # INLINE QUICK ADD ROW
    # ====================================================
    add_elements = st.columns(grid_weights)
    
    new_group = add_elements[0].text_input("Add Group", key="add_group", placeholder="GROUP / CLASS...", label_visibility="collapsed")
    new_med = add_elements[0].text_input("Add Item Name", key="add_med", placeholder="Item Description / Spec...", label_visibility="collapsed")
    
    new_soh = add_elements[1].number_input("Add SOH", min_value=0, step=1, value=0, key="add_soh", label_visibility="collapsed")
    new_transit = add_elements[2].number_input("Add Transit", min_value=0, step=1, value=0, key="add_transit", label_visibility="collapsed")
    new_amc = add_elements[3].number_input("Add AMC", min_value=0, step=1, value=0, key="add_amc", label_visibility="collapsed")
    new_req = add_elements[4].number_input("Add Request", min_value=0, step=1, value=0, key="add_req", label_visibility="collapsed")
    
    new_custom_inputs = {}
    col_idx = 5
    for c_name in custom_cols:
        new_custom_inputs[c_name] = add_elements[col_idx].text_input(f"Add {c_name}", key=f"add_custom_{c_name}", placeholder=f"{c_name}...", label_visibility="collapsed")
        col_idx += 1
        
    status_options = ["Open", "In Progress", "Waiting Supplier", "Escalated", "Completed"]
    new_status = add_elements[col_idx].selectbox("Add Status", status_options, index=0, key="add_status", label_visibility="collapsed")
    new_owner = add_elements[col_idx+1].text_input("Add Owner", placeholder="Owner...", key="add_owner", label_visibility="collapsed")
    
    add_elements[col_idx+2].markdown(f"<span style='color: #ff4b4b; font-weight: bold;'>New {active_category} Entry</span>", unsafe_allow_html=True)
    new_comment = add_elements[col_idx+3].text_input("Add Comment", placeholder="Initial update note...", key="add_comment", label_visibility="collapsed")
    
    if add_elements[col_idx+4].button("➕", key="btn_add_new_row", help=f"Insert this entry into {active_category}"):
        if not new_group.strip() or not new_med.strip():
            st.error("Both Group Name and Item Description are required!")
        else:
            success = add_new_medicine(new_group.strip().upper(), new_med.strip(), new_soh, new_transit, new_amc, new_req, category=active_category, custom_fields_dict=new_custom_inputs)
            if success:
                comment_to_save = new_comment.strip() if new_comment.strip() else "Item manually added."
                save_update(new_med.strip(), datetime.now().strftime("%Y-%m-%d %H:%M"), new_status, new_owner.strip(), comment_to_save)
                
                # RESET / CLEAR QUICK ADD INPUTS IN SESSION STATE
                st.session_state["add_group"] = ""
                st.session_state["add_med"] = ""
                st.session_state["add_soh"] = 0
                st.session_state["add_transit"] = 0
                st.session_state["add_amc"] = 0
                st.session_state["add_req"] = 0
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

    # ====================================================
    # EXISTING RECORDS LOOP
    # ====================================================
    if len(records) == 0:
        st.info(f"No items currently registered under '{active_category}'. Use the inline row above to add one.")
    else:
        for idx, row in records.iterrows():
            medicine = row["medicine_name"]
            group_label = row["base_drug_name"]
            
            r_elements = st.columns(grid_weights)
            
            r_elements[0].write(f"**{group_label}**\n\n{medicine}")
            
            soh = r_elements[1].number_input("SOH", value=int(row["stock_on_hand"]), step=1, key=f"soh_{medicine}", label_visibility="collapsed")
            transit = r_elements[2].number_input("Transit", value=int(row["in_transit"]), step=1, key=f"transit_{medicine}", label_visibility="collapsed")
            amc = r_elements[3].number_input("AMC", value=int(row["amc"]), step=1, key=f"amc_{medicine}", label_visibility="collapsed")
            request_qty = r_elements[4].number_input("Request", value=int(row["request_qty"]), step=1, key=f"request_{medicine}", label_visibility="collapsed")
            
            existing_custom_dict = json.loads(row["custom_fields"]) if row["custom_fields"] else {}
            updated_custom_dict = {}
            
            col_idx = 5
            for c_name in custom_cols:
                val = existing_custom_dict.get(c_name, "")
                updated_custom_dict[c_name] = r_elements[col_idx].text_input(c_name, value=str(val), key=f"custom_{c_name}_{medicine}", label_visibility="collapsed")
                col_idx += 1
                
            status_options = ["Open", "In Progress", "Waiting Supplier", "Escalated", "Completed"]
            def_idx = status_options.index(row["status"]) if row["status"] in status_options else 0
            selected_status = r_elements[col_idx].selectbox("Status", status_options, index=def_idx, key=f"status_{medicine}", label_visibility="collapsed")
            current_owner = r_elements[col_idx+1].text_input("Owner", value=row["owner"], key=f"owner_{medicine}", label_visibility="collapsed")
            
            r_elements[col_idx+2].caption(f"⏱ *({row['last_updated']})*: {row['last_comment']}")
            new_comment = r_elements[col_idx+3].text_input("New Comment", key=f"comment_{medicine}", label_visibility="collapsed", placeholder="Type update...")
            
            act_col1, act_col2 = r_elements[col_idx+4].columns(2)
            
            # SAVE ACTION
            if act_col1.button("💾", key=f"save_{medicine}", help=f"Save changes for {medicine}"):
                update_stock_levels(medicine, soh, transit, amc, request_qty, custom_fields_dict=updated_custom_dict)
                
                changed_metrics = (soh != int(row["stock_on_hand"]) or 
                                   transit != int(row["in_transit"]) or 
                                   amc != int(row["amc"]) or 
                                   request_qty != int(row["request_qty"]) or
                                   updated_custom_dict != existing_custom_dict)
                changed_meta = (selected_status != row["status"] or current_owner != row["owner"] or new_comment.strip() != "")
                
                if changed_metrics or changed_meta:
                    comment_to_save = new_comment.strip()
                    if not comment_to_save:
                        changes = []
                        if selected_status != row["status"]: changes.append(f"Status: {row['status']} ➔ {selected_status}")
                        if current_owner != row["owner"]: changes.append(f"Owner: {row['owner'] or 'None'} ➔ {current_owner}")
                        if soh != int(row["stock_on_hand"]): changes.append(f"SOH: {row['stock_on_hand']} ➔ {soh}")
                        if transit != int(row["in_transit"]): changes.append(f"Transit: {row['in_transit']} ➔ {transit}")
                        comment_to_save = ", ".join(changes) if changes else "Updated parameters."
                    
                    save_update(medicine, datetime.now().strftime("%Y-%m-%d %H:%M"), selected_status, current_owner, comment_to_save)
                
                st.success(f"Updated {medicine}!")
                st.rerun()
                
            # DELETE ACTION WITH AUDIT CONFIRMATION
            with act_col2.popover("🗑️"):
                st.markdown(f"### Delete Item")
                st.warning(f"Are you sure you want to delete **'{medicine}'**?")
                deleter_name = st.text_input("Your Name (Required for audit log)", key=f"deleter_name_{medicine}").strip()
                
                if st.button("🔴 Confirm Delete", key=f"confirm_del_{medicine}", type="primary"):
                    if not deleter_name:
                        st.error("Please enter your name to confirm deletion.")
                    else:
                        delete_date_time = datetime.now().strftime("%Y-%m-%d %H:%M")
                        delete_medicine(medicine, deleter_name, delete_date_time)
                        st.success(f"Item '{medicine}' deleted by {deleter_name}.")
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