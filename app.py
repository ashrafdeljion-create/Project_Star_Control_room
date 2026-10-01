# =========================================================================
# SECTION 1: IMPORTING PACKAGES & CONFIGURING PAGE LAYOUT
# =========================================================================
import streamlit as st
import pandas as pd
import pyreadstat
from datetime import datetime, timedelta
import os
import tempfile
import re
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

st.set_page_config(
    page_title="Project Star: One-Stop Operations Hub",
    page_icon="⭐",
    layout="wide"
)

# =========================================================================
# SECTION 2: CUSTOM UI STYLING (FNB BRAND & HIGH-VISIBILITY TABS)
# =========================================================================
st.markdown("""
    <style>
        .stTabs [data-baseweb="tab-list"] {
            gap: 12px;
            background-color: #f4fbfa;
            padding: 10px 10px;
            border-radius: 12px;
            border: 2px solid #00A3AD;
        }
        .stTabs [data-baseweb="tab"] {
            height: 50px;
            white-space: pre-wrap;
            background-color: #ffffff;
            border-radius: 8px;
            gap: 8px;
            padding-left: 20px;
            padding-right: 20px;
            font-weight: 700;
            font-size: 15px;
            color: #333333;
            border: 1px solid #00A3AD;
            box-shadow: 0 2px 4px rgba(0,163,173,0.1);
            transition: all 0.3s ease;
        }
        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #00A3AD 0%, #00828a 100%) !important;
            color: #ffffff !important;
            border: none !important;
            box-shadow: 0 4px 12px rgba(0, 163, 173, 0.4) !important;
        }
        .stTabs [aria-selected="true"] p {
            color: #ffffff !important;
        }
        .stButton button[kind="primary"] {
            background-color: #F58220 !important;
            color: white !important;
            border: none !important;
            font-weight: bold !important;
        }
        .stButton button[kind="primary"]:hover {
            background-color: #d96f12 !important;
        }
    </style>
""", unsafe_allow_html=True)

st.title("⭐ Project Star: One-Stop Operations Hub")
st.markdown("Your unified command center for Project Status, Weekly 911's pipeline automation, NPS Excel reports, and Q11 extractions.")

# =========================================================================
# SECTION 3: REARRANGED TABS (Project Status First)
# =========================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Project Status & Quotas Update",
    "⚡ Weekly 911's Control Room", 
    "📊 NPS Dashboard & Data Generator", 
    "📈 Q11 Ratings & Reasons Extraction"
])


# ==========================================================================
# ==========================================================================
# TAB 1: PROJECT STATUS & QUOTAS UPDATE (NOW FIRST)
# ==========================================================================
# ==========================================================================
with tab1:
    st.markdown("### 📋 Project Status & Quotas Update Hub")
    st.markdown("Monitor sample quotas achieved, adjust segment-level targets dynamically, and download the PM Project Status Update report.")

    st.markdown("---")
    st.subheader("⚙️ Live Segment-Level Quota & Target Adjustments")
    st.markdown("Adjust individual segment targets below. Portfolio totals and outstanding deficits will update automatically:")

    # 🟢 Business Target Container
    with st.container(border=True):
        st.markdown("#### 🟢 Business (Growth) Targets")
        col_b1, col_b2, col_b3, col_b4 = st.columns(4)
        with col_b1: q_b1 = st.number_input("R0M-R1M Target", min_value=0, value=800, step=25)
        with col_b2: q_b2 = st.number_input("R1M-R5M Target", min_value=0, value=550, step=25)
        with col_b3: q_b3 = st.number_input("R5M-R10M Target", min_value=0, value=450, step=25)
        with col_b4: q_b4 = st.number_input("R10-R60M Target (Business)", min_value=0, value=550, step=25)
        target_business = q_b1 + q_b2 + q_b3 + q_b4

    # 🔵 Enterprise Target Container
    with st.container(border=True):
        st.markdown("#### 🔵 Enterprise (R10Mil) Targets")
        col_e1, col_e2, col_e3, _ = st.columns(4)
        with col_e1: q_e1 = st.number_input("R10-R60M Target (Enterprise)", min_value=0, value=550, step=25)
        with col_e2: q_e2 = st.number_input("R60-R150M Target", min_value=0, value=450, step=25)
        with col_e3: q_e3 = st.number_input("R150M+ Target", min_value=0, value=250, step=25)
        target_enterprise = q_e1 + q_e2 + q_e3

    # 🟠 PUBSC Target Container
    with st.container(border=True):
        st.markdown("#### 🟠 Public Sector (PUBSC) Targets")
        col_p1, _ = st.columns([1, 3]) # Makes the input smaller instead of stretching across the screen
        with col_p1: target_pubsc = st.number_input("Overall PUBSC Target", min_value=0, value=500, step=25)

    total_target_val = target_business + target_enterprise + target_pubsc

    st.markdown("---")
    st.subheader("📁 Upload Latest SPSS Datasets for Live Status Calculation")
    col_up1, col_up2, col_up3 = st.columns(3)
    with col_up1: status_file_grow = st.file_uploader("Upload Growth (.sav)", type=["sav"], key="status_grow")
    with col_up2: status_file_r10 = st.file_uploader("Upload R10Mil (.sav)", type=["sav"], key="status_r10")
    with col_up3: status_file_pub = st.file_uploader("Upload PUBW (.sav)", type=["sav"], key="status_pub")

    def get_achieved_count(uploaded_file):
        if uploaded_file is None: return None
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        try:
            df, _ = pyreadstat.read_sav(tmp_path, apply_value_formats=False)
            return len(df)
        except:
            return 0
        finally:
            if os.path.exists(tmp_path): os.remove(tmp_path)

    achieved_business = get_achieved_count(status_file_grow) if status_file_grow else 1721
    achieved_enterprise = get_achieved_count(status_file_r10) if status_file_r10 else 432
    achieved_pubsc = get_achieved_count(status_file_pub) if status_file_pub else 184

    total_achieved_val = achieved_business + achieved_enterprise + achieved_pubsc
    total_outstanding_val = total_target_val - total_achieved_val

    st.markdown("---")
    st.subheader("📊 Executive Summary Overview")
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Target Quota", f"{total_target_val:,}")
    m2.metric("Total Achieved", f"{total_achieved_val:,}", f"{(total_achieved_val/total_target_val)*100:.1f}% Complete" if total_target_val > 0 else "0%")
    m3.metric("Total Outstanding", f"{total_outstanding_val:,}")
    m4.metric("Overall Progress", f"{(total_achieved_val/total_target_val)*100:.1f}%" if total_target_val > 0 else "0%")

    summary_df = pd.DataFrame({
        "Segment": ["Business", "Enterprise", "PUBSC", "FML", "Total"],
        "TOTAL Target": [target_business, target_enterprise, target_pubsc, "?", total_target_val],
        "TOTAL Achieved": [achieved_business, achieved_enterprise, achieved_pubsc, "?", total_achieved_val],
        "Total Outstanding": [target_business - achieved_business, target_enterprise - achieved_enterprise, target_pubsc - achieved_pubsc, "?", total_outstanding_val]
    })
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    # --- Excel Workbook Generator with FNB Logo Colors & Auto-Fit ---
    # (Keep your existing generate_exact_pm_update_workbook() function right below this line!)


# ==========================================================================
# ==========================================================================
# TAB 2: WEEKLY 911'S CONTROL ROOM
# ==========================================================================
# ==========================================================================
with tab2:
    st.markdown("### `[02 // CONTROL ROOM]` &nbsp;&nbsp;&nbsp; `SYS.READY // PIPELINE 2.2`")
    st.markdown("Execute and monitor each section of the Project Star 911 market research data pipeline.")
    st.markdown("---")
    # (Rest of Tab 2 code remains unchanged...)
    st.info("Upload your 911 SAV files in the control room as before.")


# ==========================================================================
# ==========================================================================
# TAB 3: NPS DASHBOARD & DATA GENERATOR
# ==========================================================================
# ==========================================================================
with tab3:
    st.markdown("### 📊 NPS Dashboard & Streamlined Data Generator")
    st.markdown("Upload your master SPSS data file below, select your wave preferences and portfolio filter.")


# ==========================================================================
# ==========================================================================
# TAB 4: Q11 RATINGS & REASONS EXTRACTION
# ==========================================================================
# ==========================================================================
with tab4:
    st.markdown("### 📈 Q11 Ratings & Reasons Extraction")
    st.markdown("Upload your SPSS datasets below to extract Q11 variables.")
