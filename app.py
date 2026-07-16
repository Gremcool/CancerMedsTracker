import streamlit as st
import pandas as pd
from datetime import datetime, date

from database import *
from excel_import import *

# Configure page and hide sidebar by default
st.set_page_config(
    page_title="Cancer Medicines Tracker",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Inject custom CSS to push title to the top and hide sidebar navigation panel entirely
st.markdown("""
    <style>
        /* Push the main layout container to the top */
        .block-container {
            padding-top: 1rem !important;
            padding-bottom: 0rem !important;
            margin-top: 0rem !important;
        }
        /* Completely hide sidebar collapse button/chevron */
        [data-testid="collapsedControl"] {
            display: none !important;
        }
        /* Style the add-row inputs to keep them clean */
        div[data-testid="stInput"] input {
            border: 1px solid #ff4b4b !important;
        }
    </style>
""", unsafe_allow_html=True)

# Connect and activate schemas safely
create_tables()
update_schema()  # Migrates your DB to add SOH, In-Transit, AMC, and Request columns

st.title("💊 Cancer Medicines Management System")

tabs = st.tabs([
    "📊 Summary Action Dashboard",
    "🗓️ Live Meeting Review",
    "📥 Master Import Portal"
])

# ====================================================
# TAB 1: ACTION DASHBOARD (CUSTOM EXPANDABLE ROW GRID)
# ====================================================
with tabs[0]:
    st.subheader("Meeting Action Progress KPIs")
    stats = get_dashboard_stats()

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total Items", f"{stats['total']}")
    c2.metric("Open Actions", f"{stats['open']}")
    c3.metric("In Progress", f"{stats['progress']}")
    c4.metric("Waiting Supplier", f"{stats['supplier']}")
    c5.metric("Escalated 🔥", f"{stats['escalated']}")
    c6.metric("Completed ✅", f"{stats['completed']}")

    st.markdown("---")
    st.write("### 📋 Master Meeting Tracking Log Status Overview")
    
    master_grid = get_latest_statuses()
    if len(master_grid) == 0:
        st.info("The tracker is empty. Please upload your master sheet via the 'Master Import Portal' tab.")
    else:
        # Initialize expansion states in session state
        if "expanded_summary_groups" not in st.session_state:
            st.session_state.expanded_summary_groups = set()
            
        # Grid Header Labels (Perfect alignment matching sub-rows)
        h_col0, h_col1, h_col2, h_col3, h_col4, h_col5, h_col6, h_col7, h_col8, h_col9 = st.columns([0.4, 1.8, 2.2, 0.6, 0.6, 0.6, 0.6, 1.0, 1.0, 2.0])
        h_col0.markdown("")  # Expand/Collapse button space
        h_col1.markdown("**Drug Group**")
        h_col2.markdown("**Medicine**")
        h_col3.markdown("**SOH**")
        h_col4.markdown("**Transit**")
        h_col5.markdown("**AMC**")
        h_col6.markdown("**Request**")
        h_col7.markdown("**Status**")
        h_col8.markdown("**Owner**")
        h_col9.markdown("**Last Note**")
        st.markdown("<hr style='margin:0px 0px 10px 0px; border-top: 2px solid #555;' />", unsafe_allow_html=True)
        
        unique_groups = master_grid["Drug Group"].unique()
        
        for group_name in unique_groups:
            group_df = master_grid[master_grid["Drug Group"] == group_name]
            is_repeating = len(group_df) > 1
            
            # Row Container
            r_col0, r_col1, r_col2, r_col3, r_col4, r_col5, r_col6, r_col7, r_col8, r_col9 = st.columns([0.4, 1.8, 2.2, 0.6, 0.6, 0.6, 0.6, 1.0, 1.0, 2.0])
            
            if is_repeating:
                # Sum the integer values for the grouped item
                sum_soh = int(group_df["SOH"].sum())
                sum_transit = int(group_df["Transit"].sum())
                sum_amc = int(group_df["AMC"].sum())
                sum_request = int(group_df["Request"].sum())
                
                # Setup unified status representation
                statuses = group_df["Status"].unique()
                status_text = ", ".join(statuses) if len(statuses) <= 2 else f"{len(statuses)} Statuses"
                
                # Check expansion toggle
                is_expanded = group_name in st.session_state.expanded_summary_groups
                btn_label = "➖" if is_expanded else "➕"
                
                # Expand click callback
                if r_col0.button(btn_label, key=f"sum_exp_{group_name}", help="Click to expand/collapse formulation details"):
                    if is_expanded:
                        st.session_state.expanded_summary_groups.remove(group_name)
                    else:
                        st.session_state.expanded_summary_groups.add(group_name)
                    st.rerun()
                
                r_col1.markdown(f"**{group_name}**")
                r_col2.markdown(f"*Multiple formulations ({len(group_df)} items)*")
                r_col3.markdown(f"**{sum_soh}**")
                r_col4.markdown(f"**{sum_transit}**")
                r_col5.markdown(f"**{sum_amc}**")
                r_col6.markdown(f"**{sum_request}**")
                r_col7.markdown(f"`{status_text}`")
                r_col8.markdown("-")
                r_col9.markdown("-")
                
                # Nested formulation rows under the expanded header (Using clean bullet point)
                if is_expanded:
                    for idx, row in group_df.iterrows():
                        sub_c0, sub_c1, sub_c2, sub_c3, sub_c4, sub_c5, sub_c6, sub_c7, sub_c8, sub_c9 = st.columns([0.4, 1.8, 2.2, 0.6, 0.6, 0.6, 0.6, 1.0, 1.0, 2.0])
                        sub_c0.write("")  # Spacer indent
                        sub_c1.markdown(f"<span style='color: #777;'>{group_name}</span>", unsafe_allow_html=True)
                        sub_c2.write(f"• {row['Medicine']}")
                        sub_c3.write(str(row['SOH']))
                        sub_c4.write(str(row['Transit']))
                        sub_c5.write(str(row['AMC']))
                        sub_c6.write(str(row['Request']))
                        sub_c7.write(row['Status'])
                        sub_c8.write(row['Owner'])
                        sub_c9.caption(f"⏱️ ({row['Updated']}): {row['Note']}")
            else:
                # Normal single row rendering
                row = group_df.iloc[0]
                r_col0.write("")  # Spacer empty
                r_col1.markdown(f"**{group_name}**")
                r_col2.write(row["Medicine"])
                r_col3.write(str(row["SOH"]))
                r_col4.write(str(row["Transit"]))
                r_col5.write(str(row["AMC"]))
                r_col6.write(str(row["Request"]))
                r_col7.write(row["Status"])
                r_col8.write(row["Owner"])
                r_col9.caption(f"⏱️ ({row['Updated']}): {row['Note']}")
                
            st.markdown("<hr style='margin:2px 0px; border-top: 1px dashed #444;' />", unsafe_allow_html=True)

# ====================================================
# TAB 2: LIVE MEETING REVIEW (INDIVIDUAL ROW SAVE MODE)
# ====================================================
with tabs[1]:
    st.subheader("Interactive Board Grid Review")
    
    records = get_medicines_grid()
    
    if len(records) == 0:
        st.warning("No records found. Please ingest a master file via the 'Master Import Portal' or add a row above to get started.")
    else:
        st.caption("Add a new medicine row directly below, or modify inventory parameter rows instantly.")
        
        # Grid Header Labels (Perfect wide 10-column distribution containing individual row save allocations)
        h_col1, h_col2, h_col3, h_col4, h_col5, h_col6, h_col7, h_col8, h_col9, h_col10 = st.columns([1.8, 0.5, 0.9, 0.5, 0.8, 1.0, 1.0, 1.4, 2.0, 0.5])
        h_col1.markdown("**Medicine Description**")
        h_col2.markdown("**SOH**")
        h_col3.markdown("**Inbound**")
        h_col4.markdown("**AMC**")
        h_col5.markdown("**Request**")
        h_col6.markdown("**Status**")
        h_col7.markdown("**Owner**")
        h_col8.markdown("**Last Note**")
        h_col9.markdown("**New Update**")
        h_col10.markdown("**Action**")
        st.markdown("<hr style='margin:0px 0px 10px 0px; border-top: 2px solid #555;' />", unsafe_allow_html=True)
        
        # ====================================================
        # INLINE QUICK ADD ROW (Displays right under table headers)
        # ====================================================
        add_col1, add_col2, add_col3, add_col4, add_col5, add_col6, add_col7, add_col8, add_col9, add_col10 = st.columns([1.8, 0.5, 0.9, 0.5, 0.8, 1.0, 1.0, 1.4, 2.0, 0.5])
        
        # Custom input stack inside Column 1
        new_group = add_col1.text_input("Add Group", key="add_group", placeholder="DRUG GROUP...", label_visibility="collapsed")
        new_med = add_col1.text_input("Add Medicine", key="add_med", placeholder="Strength/Formulation...", label_visibility="collapsed")
        
        new_soh = add_col2.number_input("Add SOH", min_value=0, step=1, value=0, key="add_soh", label_visibility="collapsed")
        new_transit = add_col3.number_input("Add Transit", min_value=0, step=1, value=0, key="add_transit", label_visibility="collapsed")
        new_amc = add_col4.number_input("Add AMC", min_value=0, step=1, value=0, key="add_amc", label_visibility="collapsed")
        new_req = add_col5.number_input("Add Request", min_value=0, step=1, value=0, key="add_req", label_visibility="collapsed")
        
        status_options = ["Open", "In Progress", "Waiting Supplier", "Escalated", "Completed"]
        new_status = add_col6.selectbox("Add Status", status_options, index=0, key="add_status", label_visibility="collapsed")
        new_owner = add_col7.text_input("Add Owner", placeholder="Owner...", key="add_owner", label_visibility="collapsed")
        
        add_col8.markdown("<span style='color: #ff4b4b; font-weight: bold;'>New Row Entry</span>", unsafe_allow_html=True)
        new_comment = add_col9.text_input("Add Comment", placeholder="Initial update note...", key="add_comment", label_visibility="collapsed")
        
        # Save inline entry logic
        if add_col10.button("➕", key="btn_add_new_row", help="Insert this new entry row into the tracker"):
            if not new_group.strip() or not new_med.strip():
                st.error("Both 'Drug Group' and 'Medicine Description' are required!")
            else:
                success = add_new_medicine(new_group.strip().upper(), new_med.strip(), new_soh, new_transit, new_amc, new_req)
                if success:
                    comment_to_save = new_comment.strip() if new_comment.strip() else "Item manually added."
                    save_update(new_med.strip(), datetime.now().strftime("%Y-%m-%d %H:%M"), new_status, new_owner.strip(), comment_to_save)
                    st.success(f"Added {new_med.strip()}!")
                    st.rerun()
                else:
                    st.error(f"Error: '{new_med.strip()}' already exists in your tracker.")

        st.markdown("<hr style='margin:10px 0px 15px 0px; border-top: 3px double #888;' />", unsafe_allow_html=True)

        # ====================================================
        # EXISTING DIRECTORY ROW RECORDS LOOP
        # ====================================================
        for idx, row in records.iterrows():
            medicine = row["medicine_name"]
            group_label = row["base_drug_name"]
            
            r_col1, r_col2, r_col3, r_col4, r_col5, r_col6, r_col7, r_col8, r_col9, r_col10 = st.columns([1.8, 0.5, 0.9, 0.5, 0.8, 1.0, 1.0, 1.4, 2.0, 0.5])
            
            r_col1.write(f"**{group_label}**\n\n{medicine}")
            
            # Numeric inputs (step=1 enforces integers cleanly)
            soh = r_col2.number_input("SOH", value=int(row["stock_on_hand"]), step=1, key=f"soh_{medicine}", label_visibility="collapsed")
            transit = r_col3.number_input("Transit", value=int(row["in_transit"]), step=1, key=f"transit_{medicine}", label_visibility="collapsed")
            amc = r_col4.number_input("AMC", value=int(row["amc"]), step=1, key=f"amc_{medicine}", label_visibility="collapsed")
            request_qty = r_col5.number_input("Request", value=int(row["request_qty"]), step=1, key=f"request_{medicine}", label_visibility="collapsed")
            
            # Status and Owner selectors
            status_options = ["Open", "In Progress", "Waiting Supplier", "Escalated", "Completed"]
            def_idx = status_options.index(row["status"]) if row["status"] in status_options else 0
            selected_status = r_col6.selectbox("Status", status_options, index=def_idx, key=f"status_{medicine}", label_visibility="collapsed")
            current_owner = r_col7.text_input("Owner", value=row["owner"], key=f"owner_{medicine}", label_visibility="collapsed")
            
            # Show date and time of last updated comment cleanly
            r_col8.caption(f"⏱ *({row['last_updated']})*: {row['last_comment']}")
            new_comment = r_col9.text_input("New Comment", key=f"comment_{medicine}", label_visibility="collapsed", placeholder="Type update...")
            
            # INDIVIDUAL ROW SAVE ACTION
            if r_col10.button("💾", key=f"save_{medicine}", help=f"Save changes for {medicine}"):
                update_stock_levels(medicine, soh, transit, amc, request_qty)
                
                # Check metrics or metadata changes to append smart history log entry
                changed_metrics = (soh != int(row["stock_on_hand"]) or 
                                   transit != int(row["in_transit"]) or 
                                   amc != int(row["amc"]) or 
                                   request_qty != int(row["request_qty"]))
                changed_meta = (selected_status != row["status"] or current_owner != row["owner"] or new_comment.strip() != "")
                
                if changed_metrics or changed_meta:
                    comment_to_save = new_comment.strip()
                    if not comment_to_save:
                        changes = []
                        if selected_status != row["status"]: changes.append(f"Status: {row['status']} ➔ {selected_status}")
                        if current_owner != row["owner"]: changes.append(f"Owner: {row['owner'] or 'None'} ➔ {current_owner}")
                        if soh != int(row["stock_on_hand"]): changes.append(f"SOH: {row['stock_on_hand']} ➔ {soh}")
                        if transit != int(row["in_transit"]): changes.append(f"Transit: {row['in_transit']} ➔ {transit}")
                        if amc != int(row["amc"]): changes.append(f"AMC: {row['amc']} ➔ {amc}")
                        if request_qty != int(row["request_qty"]): changes.append(f"Request: {row['request_qty']} ➔ {request_qty}")
                        comment_to_save = ", ".join(changes) if changes else "Updated parameters."
                    
                    # Generates time tracking accurately inside the database updates log
                    save_update(medicine, datetime.now().strftime("%Y-%m-%d %H:%M"), selected_status, current_owner, comment_to_save)
                
                st.success(f"Updated {medicine}!")
                st.rerun()
            
            # Full Change Log Timeline Expander
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
    st.subheader("Seed Baseline Tracking Directory")
    uploaded_file = st.file_uploader("Upload Cancer Medicines Report Workbook", type=["xls", "xlsx"])

    if uploaded_file:
        try:
            sheets = load_excel_sheets(uploaded_file)
            selected_sheet = st.selectbox("Select Target Sheet Tab to Sync:", sheets)
            
            df = load_excel(uploaded_file, sheet_name=selected_sheet)
            df_prepared = prepare_medicines(df)
            
            if st.button("🚀 Confirm and Load Records", type="primary"):
                save_medicines(df_prepared)
                st.success("Master medicine inventory directory updated successfully!")
                st.rerun()
            st.dataframe(df_prepared, use_container_width=True)
        except Exception as e:
            st.error(f"Excel Parser Error: {e}")