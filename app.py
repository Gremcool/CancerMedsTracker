import streamlit as st
import pandas as pd
from datetime import date

from database import *
from excel_import import *

st.set_page_config(
    page_title="Cancer Medicines Tracker",
    layout="wide"
)

# Connect and activate schemas safely
create_tables()

st.title("💊 Cancer Medicines Management System")

# SYSTEM STORAGE PERSISTENCE AUDIT CHECK
if not is_storage_permanent():
    st.sidebar.warning(
        "⚠️ **Temporary Storage Warning**\n\n"
        "The app is currently saving data to a temporary fallback file. "
        "To ensure your meeting notes are saved permanently, please go to your "
        "Railway dashboard, add a persistent **Volume**, and set its mount path to `/data`."
    )
else:
    st.sidebar.success("🔒 Permanent Cloud Storage Connected")

tabs = st.tabs([
    "📊 Summary Action Dashboard",
    "🗓️ Live Meeting Review",
    "📥 Master Import Portal"
])

# ====================================================
# TAB 1: ACTION DASHBOARD
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
        st.dataframe(master_grid, use_container_width=True, height=500)

# ====================================================
# TAB 2: LIVE MEETING REVIEW (EXCEL INLINE LAYOUT)
# ====================================================
with tabs[1]:
    st.subheader("Interactive Board Grid Review")
    
    records = get_medicines_grid()
    
    if len(records) == 0:
        st.warning("No data found. Please run a baseline file ingest on the Import tab first.")
    else:
        st.caption("Modify statuses, assign owners, and input notes directly on any item row below. Save individual rows instantly.")
        
        # Grid Header Labels
        h_col1, h_col2, h_col3, h_col4, h_col5, h_col6 = st.columns([3.5, 1.8, 1.8, 2.5, 3.5, 0.8])
        h_col1.markdown("**Medicine Formulation Description**")
        h_col2.markdown("**Status**")
        h_col3.markdown("**Action Owner**")
        h_col4.markdown("**Last Updated Note Context**")
        h_col5.markdown("**New Meeting Update Notes**")
        h_col6.markdown("**Save**")
        st.markdown("<hr style='margin:0px 0px 10px 0px; border-top: 2px solid #555;' />", unsafe_allow_html=True)
        
        for idx, row in records.iterrows():
            medicine = row["medicine_name"]
            group_label = row["base_drug_name"]
            
            r_col1, r_col2, r_col3, r_col4, r_col5, r_col6 = st.columns([3.5, 1.8, 1.8, 2.5, 3.5, 0.8])
            
            r_col1.write(f"**{group_label}**\n\n{medicine}")
            
            status_options = ["Open", "In Progress", "Waiting Supplier", "Escalated", "Completed"]
            try:
                def_idx = status_options.index(row["status"])
            except ValueError:
                def_idx = 0
            selected_status = r_col2.selectbox(
                "Status", status_options, index=def_idx, 
                key=f"status_{medicine}", label_visibility="collapsed"
            )
            
            current_owner = r_col3.text_input(
                "Owner", value=row["owner"], 
                key=f"owner_{medicine}", label_visibility="collapsed", placeholder="Unassigned"
            )
            
            note_preview = f"⏱️ *({row['last_updated']})*:\n\n{row['last_comment']}"
            r_col4.caption(note_preview)
            
            new_comment = r_col5.text_input(
                "New Comment", key=f"comment_{medicine}", 
                label_visibility="collapsed", placeholder="Type minutes update..."
            )
            
            if r_col6.button("💾", key=f"save_{medicine}", help=f"Save changes for {medicine}"):
                if not new_comment.strip():
                    st.error("Please provide an action description note comment before updating this row.")
                else:
                    save_update(medicine, str(date.today()), selected_status, current_owner, new_comment)
                    st.success(f"Saved: {medicine}")
                    st.rerun()
                    
            with st.expander("📜 View Full Past Change Log Timeline", expanded=False):
                history_df = get_medicine_history(medicine)
                if len(history_df) <= 1:
                    st.caption("No older historical timeline records found for this formulation.")
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
            
            default_index = 0
            for idx, s in enumerate(sheets):
                if "ORDER" in s.upper() or "ASSESSMENT" in s.upper():
                    default_index = idx
                    break
            
            selected_sheet = st.selectbox("Select Target Sheet Tab to Sync:", sheets, index=default_index)
            
            df = load_excel(uploaded_file, sheet_name=selected_sheet)
            df_prepared = prepare_medicines(df)
            
            if len(df_prepared) == 0:
                st.error("No valid medicine rows matching an 'ITEM DESCRIPTION' layout were identified.")
            else:
                st.info(f"Detected {len(df_prepared)} unique medicine variations inside tab '{selected_sheet}'.")
                if st.button("🚀 Confirm and Load Records", type="primary"):
                    save_medicines(df_prepared)
                    st.success("Master medicine inventory directory updated successfully!")
                    st.rerun()
                
                st.write("### Data Columns Preview Matrix:")
                st.dataframe(df_prepared, use_container_width=True)
        except Exception as e:
            st.error(f"Excel File Parser Exception Encountered: {e}")