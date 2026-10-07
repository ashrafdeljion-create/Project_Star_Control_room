# =========================================================================
# SECTION 1: IMPORTING PACKAGES & CONFIGURING PAGE LAYOUT
# =========================================================================
import io
import os
import re
import tempfile
from datetime import datetime, timedelta
from dateutil.relativedelta import FR, relativedelta
from itertools import groupby

import numpy as np
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import pyreadstat
import streamlit as st

# Configures browser tab title, icon, and sets layout to wide mode
st.set_page_config(
    page_title="Project Star: One-Stop Operations Hub", page_icon="⭐", layout="wide"
)

# =========================================================================
# SECTION 2: CUSTOM UI STYLING (FNB BRAND & HIGH-VISIBILITY TABS)
# =========================================================================
st.markdown(
    """
    <style>
        /* Styles the main background container holding the navigation tabs */
        .stTabs [data-baseweb="tab-list"] {
            gap: 12px;
            background-color: #f4fbfa;
            padding: 10px 10px;
            border-radius: 12px;
            border: 2px solid #00A3AD;
        }
        /* Styles each individual unselected tab button */
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
        /* Styles the currently active/selected tab */
        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #00A3AD 0%, #00828a 100%) !important;
            color: #ffffff !important;
            border: none !important;
            box-shadow: 0 4px 12px rgba(0, 163, 173, 0.4) !important;
        }
        .stTabs [aria-selected="true"] p {
            color: #ffffff !important;
        }
        /* Styles primary action buttons (FNB Orange) */
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
""",
    unsafe_allow_html=True,
)

# Main title and subtitle displayed at the top of the application web page
st.title("⭐ Project Star: One-Stop Operations Hub")
st.markdown(
    "Your unified command center for Project Star, Weekly 911's pipeline automation, BM/RM NPS Portfolio Generator, Q11 extractions, Yearly Dashboard generation, and SME/ENT tables."
)

# =========================================================================
# SECTION 3: DEFINING MAIN APP NAVIGATION TABS (6 Tabs Total)
# =========================================================================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "📊 Project Status & Quotas Update",
        "⚙️ Weekly 911's Control Room",
        "📈 BM/RM NPS Portfolio Generator",
        "📋 Q11 Ratings & Reasons Extraction",
        "📅 NPS Yearly Dashboard",
        "🏢 SME/ENT Tables",
    ]
)

# ==========================================================================
# ==========================================================================
# TAB 1: PROJECT STATUS & QUOTAS UPDATE
# ==========================================================================
# ==========================================================================
with tab1:
    st.markdown("### 📊 Live Status Calculation & Quota Tracker")
    st.markdown("---")

    # SECTION 1: LIVE QUOTA TARGET ADJUSTMENTS
    st.markdown("### 🎯 Live Quota Target Adjustments")
    col1, col2, col3 = st.columns(3)

    with col1:
        business_target = st.number_input(
            "Business (Growth) Target", min_value=0, value=4700, step=50, key="bus_tgt_1"
        )
    with col2:
        enterprise_target = st.number_input(
            "Enterprise (R10Mil) Target", min_value=0, value=1400, step=50, key="ent_tgt_1"
        )
    with col3:
        pubsc_target = st.number_input(
            "PUBSC Target", min_value=0, value=500, step=50, key="pub_tgt_1"
        )

    total_target_quota = business_target + enterprise_target + pubsc_target

    st.markdown("---")

    # SECTION 2: SEGMENT-LEVEL QUOTA BREAKDOWN
    st.markdown("### 🔢 Segment-Level Quota Breakdown Inputs")
    st.markdown("Specify exact individual segment quotas below for detailed tracking:")

    seg_cols = st.columns(4)
    with seg_cols[0]:
        q_rom_r1m = st.number_input("R0M-R1M Quota", min_value=0, value=1600, step=25, key="q_r0_r1_1")
    with seg_cols[1]:
        q_r1m_r5m = st.number_input("R1M-R5M Quota", min_value=0, value=1100, step=25, key="q_r1_r5_1")
    with seg_cols[2]:
        q_r5m_r10m = st.number_input(
            "R5M-R10M Quota", min_value=0, value=900, step=25, key="q_r5_r10_1"
        )
    with seg_cols[3]:
        q_r10_r60m = st.number_input(
            "R10-R60M Quota", min_value=0, value=1100, step=25, key="q_r10_r60_1"
        )

    seg_cols_2 = st.columns(2)
    with seg_cols_2[0]:
        q_r60_r150m = st.number_input(
            "R60-R150M Quota", min_value=0, value=900, step=25, key="q_r60_r150_1"
        )
    with seg_cols_2[1]:
        q_r150m_plus = st.number_input(
            "R150M+ Quota", min_value=0, value=500, step=25, key="q_r150_plus_1"
        )

    total_segment_quota = (
        q_rom_r1m
        + q_r1m_r5m
        + q_r5m_r10m
        + q_r10_r60m
        + q_r60_r150m
        + q_r150m_plus
    )

    st.markdown("---")

    # SECTION 3: UPLOAD SPSS DATASETS (.sav)
    st.markdown("### 📁 Upload Latest SPSS Datasets for Live Status Calculation")
    up_col1, up_col2, up_col3 = st.columns(3)

    with up_col1:
        file_grow = st.file_uploader("Upload Growth (.sav)", type=["sav"], key="up_grow_1")
    with up_col2:
        file_rmw = st.file_uploader("Upload R10Mil (.sav)", type=["sav"], key="up_rmw_1")
    with up_col3:
        file_pubw = st.file_uploader("Upload PUBSC (.sav)", type=["sav"], key="up_pub_1")

    def load_spss_data(uploaded_file):
        if uploaded_file is not None:
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_path = tmp_file.name

                df, meta = pyreadstat.read_sav(tmp_path, apply_value_formats=True)
                df.columns = [str(c).upper() for c in df.columns]
                return df
            except Exception as e:
                st.error(f"Error reading file: {e}")
                return None
        return None

    df_grow = load_spss_data(file_grow)
    df_rmw = load_spss_data(file_rmw)
    df_pubw = load_spss_data(file_pubw)

    def count_completions(df):
        if df is not None and "V9999" in df.columns:
            str_val = df["V9999"].astype(str).str.lower()
            completed_filter = str_val.str.contains("continue", na=False)
            return int(completed_filter.sum())
        return 0

    achieved_growth = count_completions(df_grow)
    achieved_rmw = count_completions(df_rmw)
    achieved_pubw = count_completions(df_pubw)

    total_achieved = achieved_growth + achieved_rmw + achieved_pubw
    total_outstanding = max(0, total_target_quota - total_achieved)
    overall_progress = (
        (total_achieved / total_target_quota * 100)
        if total_target_quota > 0
        else 0.0
    )

    st.markdown("---")

    # SECTION 4: EXECUTIVE SUMMARY OVERVIEW
    st.markdown("### 📋 Executive Summary Overview")

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric(label="Total Target Quota", value=f"{total_target_quota:,.0f}")
    with m2:
        st.metric(
            label="Total Achieved",
            value=f"{total_achieved:,.0f}",
            delta=f"{overall_progress:.1f}% Complete",
        )
    with m3:
        st.metric(label="Total Outstanding", value=f"{total_outstanding:,.0f}")
    with m4:
        st.metric(label="Overall Progress", value=f"{overall_progress:.1f}%")

    summary_data = {
        "Segment": ["Business", "Enterprise", "PUBSC", "Total"],
        "TOTAL Target": [
            business_target,
            enterprise_target,
            pubsc_target,
            total_target_quota,
        ],
        "TOTAL Achieved": [
            achieved_growth,
            achieved_rmw,
            achieved_pubw,
            total_achieved,
        ],
        "TOTAL Outstanding": [
            max(0, business_target - achieved_growth),
            max(0, enterprise_target - achieved_rmw),
            max(0, pubsc_target - achieved_pubw),
            total_outstanding,
        ],
    }
    df_summary = pd.DataFrame(summary_data)
    st.dataframe(df_summary, use_container_width=True, hide_index=True)

    st.markdown("---")

    # SECTION 5: SEGMENT QUOTAS EXECUTIVE SUMMARY BREAKDOWN
    st.markdown("### 📊 Segment Quotas Executive Summary Breakdown")

    def get_segment_achieved(target_df, segment_keywords):
        count = 0
        if target_df is not None and "V9999" in target_df.columns and "V44011" in target_df.columns:
            str_val = target_df["V9999"].astype(str).str.lower()
            completed = str_val.str.contains("continue", na=False)

            matched_seg = pd.Series(False, index=target_df.index)
            for kw in segment_keywords:
                matched_seg = matched_seg | (
                    target_df["V44011"]
                    .astype(str)
                    .str.contains(kw, case=False, na=False)
                )
            count = int((completed & matched_seg).sum())
        return count

    ach_rom_r1m = get_segment_achieved(df_grow, ["r0m-r1m"])
    ach_r1m_r5m = get_segment_achieved(df_grow, ["r1m-r5m"])
    ach_r5m_r10m = get_segment_achieved(df_grow, ["r5m-r10"])

    ach_r10_r60m = get_segment_achieved(df_rmw, ["r10m-r60m"])
    ach_r60_r150m = get_segment_achieved(df_rmw, ["r60m-r150"])
    ach_r150m_plus = get_segment_achieved(df_rmw, ["r150m+"])

    total_seg_achieved = (
        ach_rom_r1m
        + ach_r1m_r5m
        + ach_r5m_r10m
        + ach_r10_r60m
        + ach_r60_r150m
        + ach_r150m_plus
    )
    total_seg_outstanding = max(0, total_segment_quota - total_seg_achieved)

    seg_breakdown_data = {
        "Segment": [
            "R0M-R1M",
            "R1M-R5M",
            "R5M-R10M",
            "R10-R60M",
            "R60-R150M",
            "R150M+",
            "Total",
        ],
        "TOTAL Target": [
            q_rom_r1m,
            q_r1m_r5m,
            q_r5m_r10m,
            q_r10_r60m,
            q_r60_r150m,
            q_r150m_plus,
            total_segment_quota,
        ],
        "TOTAL Achieved": [
            ach_rom_r1m,
            ach_r1m_r5m,
            ach_r5m_r10m,
            ach_r10_r60m,
            ach_r60_r150m,
            ach_r150m_plus,
            total_seg_achieved,
        ],
        "TOTAL Outstanding": [
            max(0, q_rom_r1m - ach_rom_r1m),
            max(0, q_r1m_r5m - ach_r1m_r5m),
            max(0, q_r5m_r10m - ach_r5m_r10m),
            max(0, q_r10_r60m - ach_r10_r60m),
            max(0, q_r60_r150m - ach_r60_r150m),
            max(0, q_r150m_plus - ach_r150m_plus),
            total_seg_outstanding,
        ],
    }

    df_seg_breakdown = pd.DataFrame(seg_breakdown_data)
    st.dataframe(df_seg_breakdown, use_container_width=True, hide_index=True)

    st.markdown("---")

    # SECTION 6: BUSINESS REGIONAL & SEGMENT BREAKDOWN
    standard_subregions = [
        "Eastern Cape",
        "Free State",
        "Gauteng East",
        "Gauteng South Central/Klipriver",
        "Gauteng Tshwane East",
        "Gauteng Tshwane North",
        "Gauteng Tshwane South",
        "Gauteng West-Rand",
        "Greater Sandton",
        "Gauteng Midrand",
        "KZN North",
        "KZN South",
        "KZN West",
        "Limpopo",
        "Mpumalanga",
        "North West",
        "Northern Cape",
        "Western Cape",
    ]

    standard_regions = ["Cape", "Gauteng North", "Gauteng South Central", "Inland", "KwaZulu-Natal"]

    def map_subregion_and_region(row_sub, row_reg):
        sub_str = str(row_sub).strip() if pd.notnull(row_sub) else ""
        reg_str = str(row_reg).strip() if pd.notnull(row_reg) else ""
        sub_lower = sub_str.lower()
        
        if "eastern cape" in sub_lower:
            return "Eastern Cape", "Cape"
        elif "free state" in sub_lower:
            return "Free State", "Inland"
        elif "gauteng east" in sub_lower:
            return "Gauteng East", "Gauteng South Central"
        elif "klipriver" in sub_lower or "south central" in sub_lower:
            return "Gauteng South Central/Klipriver", "Gauteng South Central"
        elif "tshwane east" in sub_lower:
            return "Gauteng Tshwane East", "Gauteng North"
        elif "tshwane north" in sub_lower:
            return "Gauteng Tshwane North", "Gauteng North"
        elif "tshwane south" in sub_lower:
            return "Gauteng Tshwane South", "Gauteng North"
        elif "west-rand" in sub_lower or "west rand" in sub_lower:
            return "Gauteng West-Rand", "Gauteng South Central"
        elif "greater sandton" in sub_lower or "sandton" in sub_lower:
            if "gn growth" in sub_lower or "north" in reg_str.lower():
                return "Greater Sandton", "Gauteng North"
            return "Greater Sandton", "Gauteng South Central"
        elif "midrand" in sub_lower:
            if "gn growth" in sub_lower or "north" in reg_str.lower():
                return "Gauteng Midrand", "Gauteng North"
            return "Gauteng Midrand", "Gauteng South Central"
        elif "kzn north" in sub_lower:
            return "KZN North", "KwaZulu-Natal"
        elif "kzn south" in sub_lower:
            return "KZN South", "KwaZulu-Natal"
        elif "kzn west" in sub_lower or "kzn" in sub_lower:
            return "KZN West", "KwaZulu-Natal"
        elif "limpopo" in sub_lower:
            return "Limpopo", "Inland"
        elif "mpumalanga" in sub_lower or "moumalanga" in sub_lower:
            return "Mpumalanga", "Inland"
        elif "north west" in sub_lower:
            return "North West", "Inland"
        elif "northern cape" in sub_lower:
            return "Northern Cape", "Inland"
        elif "western cape" in sub_lower:
            return "Western Cape", "Cape"
        
        return sub_str if sub_str else "Unknown", reg_str if reg_str else "Unknown"

    bus_reg_matrix = pd.DataFrame(0, index=standard_subregions, columns=standard_regions + ["Total"])

    if df_grow is not None and "V9999" in df_grow.columns and "V13290" in df_grow.columns:
        str_val = df_grow["V9999"].astype(str).str.lower()
        completed_df = df_grow[str_val.str.contains("continue", na=False)].copy()
        
        for idx, row in completed_df.iterrows():
            raw_sub = row.get("V13290", "")
            raw_reg = row.get("V12290", "")
            mapped_sub, mapped_reg = map_subregion_and_region(raw_sub, raw_reg)
            if mapped_sub in bus_reg_matrix.index and mapped_reg in bus_reg_matrix.columns:
                bus_reg_matrix.loc[mapped_sub, mapped_reg] += 1

    bus_reg_matrix["Total"] = bus_reg_matrix[standard_regions].sum(axis=1)

    standard_segments = ["R0M-R1M", "R1M-R5M", "R5M-R10M", "R10-R60M"]
    bus_seg_matrix = pd.DataFrame(0, index=standard_segments, columns=standard_regions + ["Total", "Quota", "Outstanding"])
    bus_quotas = {"R0M-R1M": 1600, "R1M-R5M": 1100, "R5M-R10M": 900, "R10-R60M": 1100}

    if df_grow is not None and "V9999" in df_grow.columns and "V13290" in df_grow.columns and "V44011" in df_grow.columns:
        str_val = df_grow["V9999"].astype(str).str.lower()
        completed_df = df_grow[str_val.str.contains("continue", na=False)].copy()
        
        for idx, row in completed_df.iterrows():
            raw_sub = row.get("V13290", "")
            raw_reg = row.get("V12290", "")
            _, mapped_reg = map_subregion_and_region(raw_sub, raw_reg)
            
            v44 = str(row.get("V44011", "")).lower()
            seg_name = None
            if "r0m-r1m" in v44:
                seg_name = "R0M-R1M"
            elif "r1m-r5m" in v44:
                seg_name = "R1M-R5M"
            elif "r5m-r10" in v44:
                seg_name = "R5M-R10M"
            elif "r10m-r60m" in v44:
                seg_name = "R10-R60M"
                
            if seg_name and seg_name in bus_seg_matrix.index and mapped_reg in bus_seg_matrix.columns:
                bus_seg_matrix.loc[seg_name, mapped_reg] += 1

    bus_seg_matrix["Total"] = bus_seg_matrix[standard_regions].sum(axis=1)
    for seg in standard_segments:
        q_val = bus_quotas.get(seg, 0)
        ach_val = bus_seg_matrix.loc[seg, "Total"]
        bus_seg_matrix.loc[seg, "Quota"] = q_val
        bus_seg_matrix.loc[seg, "Outstanding"] = max(0, q_val - ach_val)

    crosstab_segments = ["R0m-R1m", "R1m-R5m", "R5m-R10"]
    bus_reg_seg_crosstab = pd.DataFrame(0, index=standard_subregions, columns=crosstab_segments + ["TOTAL"])

    if df_grow is not None and "V9999" in df_grow.columns and "V13290" in df_grow.columns and "V44011" in df_grow.columns:
        str_val = df_grow["V9999"].astype(str).str.lower()
        completed_df = df_grow[str_val.str.contains("continue", na=False)].copy()
        
        for idx, row in completed_df.iterrows():
            raw_sub = row.get("V13290", "")
            raw_reg = row.get("V12290", "")
            mapped_sub, _ = map_subregion_and_region(raw_sub, raw_reg)
            
            v44 = str(row.get("V44011", "")).lower()
            col_name = None
            if "r0m-r1m" in v44:
                col_name = "R0m-R1m"
            elif "r1m-r5m" in v44:
                col_name = "R1m-R5m"
            elif "r5m-r10" in v44:
                col_name = "R5m-R10"
                
            if mapped_sub in bus_reg_seg_crosstab.index and col_name in bus_reg_seg_crosstab.columns:
                bus_reg_seg_crosstab.loc[mapped_sub, col_name] += 1

    bus_reg_seg_crosstab["TOTAL"] = bus_reg_seg_crosstab[crosstab_segments].sum(axis=1)
    total_row = bus_reg_seg_crosstab.sum(numeric_only=True)
    bus_reg_seg_crosstab.loc["TOTAL"] = total_row

    # SECTION 7: ENTERPRISE (R10MIL) REGIONAL & SEGMENT BREAKDOWN
    ent_subregions = [
        "Eastern Cape",
        "Free State",
        "Gauteng East",
        "Gauteng Klipriver",
        "Gauteng North",
        "Gauteng South-West",
        "Gauteng Tshwane",
        "Greater Sandton",
        "KZN Coastal",
        "KZN Inland",
        "Limpopo",
        "Midrand",
        "Mpumalanga",
        "North West",
        "Northern Cape",
        "Western Cape Inland",
        "Western Cape Metro",
    ]

    ent_regions = ["Cape", "Gauteng South and Central", "Gauteng-North", "Inland", "KwaZulu-Natal"]

    def map_ent_subregion_and_region(row_sub, row_reg):
        sub_str = str(row_sub).strip() if pd.notnull(row_sub) else ""
        reg_str = str(row_reg).strip() if pd.notnull(row_reg) else ""
        sub_upper = sub_str.upper()
        
        if "EASTERN CAPE" in sub_upper:
            return "Eastern Cape", "Cape"
        elif "FREE STATE" in sub_upper:
            return "Free State", "Inland"
        elif "GAUTENG EAST" in sub_upper:
            return "Gauteng East", "Gauteng-North"
        elif "GAUTENG KLIPRIVER" in sub_upper:
            return "Gauteng Klipriver", "Gauteng-North"
        elif "GAUTENG NORTH" in sub_upper:
            return "Gauteng North", "Gauteng-North"
        elif "GAUTENG WEST" in sub_upper or "SOUTH-WEST" in sub_upper:
            return "Gauteng South-West", "Gauteng-North"
        elif "GAUTENG TSHWANE" in sub_upper:
            return "Gauteng Tshwane", "Gauteng South and Central"
        elif "GREATER SANDTON" in sub_upper:
            return "Greater Sandton", "Gauteng South and Central"
        elif "KZN COASTAL" in sub_upper:
            return "KZN Coastal", "KwaZulu-Natal"
        elif "KZN INLAND" in sub_upper:
            return "KZN Inland", "KwaZulu-Natal"
        elif "LIMPOPO" in sub_upper:
            return "Limpopo", "Inland"
        elif "MIDRAND" in sub_upper:
            return "Midrand", "Gauteng South and Central"
        elif "MPUMALANGA" in sub_upper:
            return "Mpumalanga", "Inland"
        elif "NORTH WEST" in sub_upper:
            return "North West", "Inland"
        elif "NORTHERN CAPE" in sub_upper:
            return "Northern Cape", "Inland"
        elif "WESTERN CAPE INLAND" in sub_upper:
            return "Western Cape Inland", "Cape"
        elif "WESTERN CAPE METRO" in sub_upper:
            return "Western Cape Metro", "Cape"
            
        return sub_str if sub_str else "Unknown", reg_str if reg_str else "Unknown"

    ent_reg_matrix = pd.DataFrame(0, index=ent_subregions, columns=ent_regions + ["Total"])

    if df_rmw is not None and "V9999" in df_rmw.columns and "V13290" in df_rmw.columns:
        str_val = df_rmw["V9999"].astype(str).str.lower()
        completed_ent = df_rmw[str_val.str.contains("continue", na=False)].copy()
        
        for idx, row in completed_ent.iterrows():
            raw_sub = row.get("V13290", "")
            raw_reg = row.get("V12290", "")
            mapped_sub, mapped_reg = map_ent_subregion_and_region(raw_sub, raw_reg)
            if mapped_sub in ent_reg_matrix.index and mapped_reg in ent_reg_matrix.columns:
                ent_reg_matrix.loc[mapped_sub, mapped_reg] += 1

    ent_reg_matrix["Total"] = ent_reg_matrix[ent_regions].sum(axis=1)

    ent_segments = ["R10-R60M", "R60-R150M", "R150M+"]
    ent_seg_matrix = pd.DataFrame(0, index=ent_segments, columns=ent_regions + ["Total", "Quota", "Outstanding"])
    ent_quotas = {"R10-R60M": 550, "R60-R150M": 450, "R150M+": 250}

    if df_rmw is not None and "V9999" in df_rmw.columns and "V13290" in df_rmw.columns and "V44011" in df_rmw.columns:
        str_val = df_rmw["V9999"].astype(str).str.lower()
        completed_ent = df_rmw[str_val.str.contains("continue", na=False)].copy()
        
        for idx, row in completed_ent.iterrows():
            raw_sub = row.get("V13290", "")
            raw_reg = row.get("V12290", "")
            _, mapped_reg = map_ent_subregion_and_region(raw_sub, raw_reg)
            
            v44 = str(row.get("V44011", "")).lower()
            seg_name = None
            if "r10m-r60m" in v44 or "r10-r60m" in v44:
                seg_name = "R10-R60M"
            elif "r60m-r150" in v44 or "r60-r150m" in v44:
                seg_name = "R60-R150M"
            elif "r150m+" in v44:
                seg_name = "R150M+"
                
            if seg_name and seg_name in ent_seg_matrix.index and mapped_reg in ent_seg_matrix.columns:
                ent_seg_matrix.loc[seg_name, mapped_reg] += 1

    ent_seg_matrix["Total"] = ent_seg_matrix[ent_regions].sum(axis=1)
    for seg in ent_segments:
        q_val = ent_quotas.get(seg, 0)
        ach_val = ent_seg_matrix.loc[seg, "Total"]
        ent_seg_matrix.loc[seg, "Quota"] = q_val
        ent_seg_matrix.loc[seg, "Outstanding"] = max(0, q_val - ach_val)

    crosstab_ent_segments = ["R10m-R60m", "R150m+", "R60m-R150"]
    ent_reg_seg_crosstab = pd.DataFrame(0, index=ent_subregions, columns=crosstab_ent_segments + ["TOTAL"])

    if df_rmw is not None and "V9999" in df_rmw.columns and "V13290" in df_rmw.columns and "V44011" in df_rmw.columns:
        str_val = df_rmw["V9999"].astype(str).str.lower()
        completed_ent = df_rmw[str_val.str.contains("continue", na=False)].copy()
        
        for idx, row in completed_ent.iterrows():
            raw_sub = row.get("V13290", "")
            raw_reg = row.get("V12290", "")
            mapped_sub, _ = map_ent_subregion_and_region(raw_sub, raw_reg)
            
            v44 = str(row.get("V44011", "")).lower()
            col_name = None
            if "r10m-r60m" in v44 or "r10-r60m" in v44:
                col_name = "R10m-R60m"
            elif "r150m+" in v44:
                col_name = "R150m+"
            elif "r60m-r150" in v44 or "r60-r150m" in v44:
                col_name = "R60m-R150"
                
            if mapped_sub in ent_reg_seg_crosstab.index and col_name in ent_reg_seg_crosstab.columns:
                ent_reg_seg_crosstab.loc[mapped_sub, col_name] += 1

    ent_reg_seg_crosstab["TOTAL"] = ent_reg_seg_crosstab[crosstab_ent_segments].sum(axis=1)
    ent_total_row = ent_reg_seg_crosstab.sum(numeric_only=True)
    ent_reg_seg_crosstab.loc["TOTAL"] = ent_total_row

    # SECTION 8: PUBSC REGIONAL & SECTOR CROSSTAB
    pubsc_regions = [
        "EASTERN CAPE",
        "FREE STATE",
        "GAUTENG",
        "KWAZULU-NATAL",
        "LIMPOPO",
        "MPUMALANGA",
        "NORTH WEST",
        "NORTHERN CAPE",
        "WESTERN CAPE"
    ]

    pubsc_sectors = [
        "NON-PROFIT ORGANISATION",
        "PUBLIC SECTOR COLLEGES & FET'S",
        "PUBLIC SECTOR EMBASSIES",
        "PUBLIC SECTOR LOCAL GOVERMENT",
        "PUBLIC SECTOR PROVINCIAL GOVER",
        "PUBLIC SECTOR PUBLIC SCHOOLS",
        "PUBLIC SECTOR UNIONS & POLITIC"
    ]

    def map_pubsc_region_and_sector(row_reg, row_sec):
        reg_str = str(row_reg).strip().upper() if pd.notnull(row_reg) else ""
        sec_str = str(row_sec).strip().upper() if pd.notnull(row_sec) else ""
        
        mapped_reg = "UNKNOWN"
        if "EASTERN" in reg_str:
            mapped_reg = "EASTERN CAPE"
        elif "FREE" in reg_str:
            mapped_reg = "FREE STATE"
        elif "GAUTENG" in reg_str:
            mapped_reg = "GAUTENG"
        elif "KWAZULU" in reg_str or "KZN" in reg_str:
            mapped_reg = "KWAZULU-NATAL"
        elif "LIMPOPO" in reg_str:
            mapped_reg = "LIMPOPO"
        elif "MPUMALANGA" in reg_str:
            mapped_reg = "MPUMALANGA"
        elif "NORTH WEST" in reg_str:
            mapped_reg = "NORTH WEST"
        elif "NORTHERN" in reg_str:
            mapped_reg = "NORTHERN CAPE"
        elif "WESTERN" in reg_str:
            mapped_reg = "WESTERN CAPE"
            
        mapped_sec = "UNKNOWN"
        for s in pubsc_sectors:
            if s in sec_str or sec_str in s:
                mapped_sec = s
                break
                
        return mapped_reg, mapped_sec

    pubsc_crosstab = pd.DataFrame(0, index=pubsc_sectors, columns=pubsc_regions + ["TOTAL"])

    if df_pubw is not None and "V9999" in df_pubw.columns and "V12290" in df_pubw.columns and "V13290" in df_pubw.columns:
        str_val = df_pubw["V9999"].astype(str).str.lower()
        completed_pub = df_pubw[str_val.str.contains("continue", na=False)].copy()
        
        for idx, row in completed_pub.iterrows():
            raw_reg = row.get("V12290", "")
            raw_sec = row.get("V13290", "")
            mapped_reg, mapped_sec = map_pubsc_region_and_sector(raw_reg, raw_sec)
            
            if mapped_sec in pubsc_crosstab.index and mapped_reg in pubsc_crosstab.columns:
                pubsc_crosstab.loc[mapped_sec, mapped_reg] += 1

    pubsc_crosstab["TOTAL"] = pubsc_crosstab[pubsc_regions].sum(axis=1)
    pubsc_total_row = pubsc_crosstab.sum(numeric_only=True)
    pubsc_crosstab.loc["TOTAL"] = pubsc_total_row

    # UI TABS FOR BREAKDOWNS
    st.markdown("### 🔍 Live Regional & Segment Breakdown Tables")
    sub_tab1, sub_tab2, sub_tab3 = st.tabs(["Business Breakdown", "Enterprise Breakdown", "PUBSC Breakdown"])

    with sub_tab1:
        st.markdown("#### Business Regional Breakdown")
        st.dataframe(bus_reg_matrix, use_container_width=True)
        st.markdown("#### Business Segment Breakdown Matrix (Quota & Outstanding)")
        st.dataframe(bus_seg_matrix, use_container_width=True)
        st.markdown("#### Business Regional vs. Segments Crosstab (Sub-regions as Rows, Segments as Columns)")
        st.dataframe(bus_reg_seg_crosstab, use_container_width=True)

    with sub_tab2:
        st.markdown("#### Enterprise Regional Breakdown")
        st.dataframe(ent_reg_matrix, use_container_width=True)
        st.markdown("#### Enterprise Segment Breakdown Matrix (Quota & Outstanding)")
        st.dataframe(ent_seg_matrix, use_container_width=True)
        st.markdown("#### Enterprise Regional vs. Segments Crosstab (Sub-regions as Rows, Segments as Columns)")
        st.dataframe(ent_reg_seg_crosstab, use_container_width=True)

    with sub_tab3:
        st.markdown("#### Public Sector (PUBSC) Regional vs. Sector Crosstab")
        st.dataframe(pubsc_crosstab, use_container_width=True)

    # EXCEL WORKBOOK GENERATION
    st.markdown("---")
    st.markdown("### 📥 Download PM Update Workbook")

    def style_excel_sheet(ws):
        teal_fill = PatternFill(start_color="005E5D", end_color="005E5D", fill_type="solid")
        teal_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        teal_sub_fill = PatternFill(start_color="E0F2F1", end_color="E0F2F1", fill_type="solid")
        teal_sub_font = Font(name="Calibri", size=11, bold=True, color="003333")

        orange_fill = PatternFill(start_color="D35400", end_color="D35400", fill_type="solid")
        orange_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        orange_sub_fill = PatternFill(start_color="FDEBD0", end_color="FDEBD0", fill_type="solid")
        orange_sub_font = Font(name="Calibri", size=11, bold=True, color="7E5109")

        border_thin = Border(
            left=Side(style='thin', color='D3D3D3'),
            right=Side(style='thin', color='D3D3D3'),
            top=Side(style='thin', color='D3D3D3'),
            bottom=Side(style='thin', color='D3D3D3')
        )

        header_rows = [1]
        for r in range(2, ws.max_row + 1):
            val_col1 = ws.cell(row=r, column=1).value
            val_col2 = ws.cell(row=r, column=2).value
            if val_col1 in ["R0m-R1m", "R10m-R60m", "NON-PROFIT ORGANISATION"] or val_col2 in ["R0m-R1m", "R10m-R60m", "R1m-R5m", "R150m+", "Gauteng North", "Gauteng South and Central"]:
                header_rows.append(r)

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.value is not None:
                    cell.border = border_thin
                    is_lower_table = any(cell.row >= hr for hr in header_rows[1:]) if len(header_rows) > 1 else False
                    
                    if cell.row in header_rows:
                        if cell.row == 1:
                            cell.fill = teal_fill
                            cell.font = teal_font
                        else:
                            cell.fill = orange_fill
                            cell.font = orange_font
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif str(cell.value).upper() in ["TOTAL", "TOTAL INLC R10-R60MIL"]:
                        if is_lower_table:
                            cell.fill = orange_sub_fill
                            cell.font = orange_sub_font
                        else:
                            cell.fill = teal_sub_fill
                            cell.font = teal_sub_font

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    val_str = str(cell.value)
                    if len(val_str) > max_len:
                        max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    def create_pm_workbook():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            bus_reg_matrix.to_excel(writer, sheet_name='Update Business', startrow=0, startcol=0)
            bus_seg_matrix.to_excel(writer, sheet_name='Update Business', startrow=0, startcol=10)
            bus_reg_seg_crosstab.to_excel(writer, sheet_name='Update Business', startrow=22, startcol=0)
            
            ent_reg_matrix.to_excel(writer, sheet_name='Update Enterprise', startrow=0, startcol=0)
            ent_seg_matrix.to_excel(writer, sheet_name='Update Enterprise', startrow=0, startcol=10)
            ent_reg_seg_crosstab.to_excel(writer, sheet_name='Update Enterprise', startrow=22, startcol=0)
            
            pubsc_crosstab.to_excel(writer, sheet_name='Update PUBSC', startrow=0, startcol=0)

            for sheetname in writer.sheets:
                ws = writer.sheets[sheetname]
                style_excel_sheet(ws)

        return output.getvalue()

    excel_data = create_pm_workbook()
    st.download_button(
        label="📊 Generate & Download Styled PM Update Workbook",
        data=excel_data,
        file_name="Star_Detailed_Update_Live.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_pm_update_excel_final_tab1"
    )

# ==========================================================================
# ==========================================================================
# TAB 2: WEEKLY 911'S CONTROL ROOM
# ==========================================================================
# ==========================================================================
with tab2:
    st.markdown("### `[02 // CONTROL ROOM]` &nbsp;&nbsp;&nbsp; `SYS.READY // PIPELINE 2.2`")
    st.markdown("Execute and monitor each section of the Project Star 911 market research data pipeline.")
    st.markdown("---")

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.metric("Pipeline Status", "IDLE / READY", "Stable")
    col_m2.metric("Active Wave", "Wave 22", "2026")
    col_m3.metric("Modules Loaded", "3 / 3", "Growth, R10Mil, PUBSC")
    col_m4.metric("Environment", "Cloud Control Room", "Secure")

    st.markdown("---")
    st.subheader("📅 Global Execution Parameters")
    date_mode = st.radio("Select Date Filtering Mode for Runs:", ["Dynamic Past 7 Days (Auto Friday)", "Custom Date Range"], horizontal=True, key="911_date_mode")

    today = datetime.now()

    if date_mode == "Dynamic Past 7 Days (Auto Friday)":
        current_weekday = today.weekday()
        if current_weekday == 4:
            days_to_subtract = 7
        else:
            days_to_subtract = (current_weekday - 4) % 7
            if days_to_subtract == 0:
                days_to_subtract = 7
        last_friday = today - timedelta(days=days_to_subtract)
        last_friday = last_friday.replace(hour=0, minute=0, second=0, microsecond=0)
        st.info(f"🎯 Target active execution window: **{last_friday.strftime('%Y-%m-%d')}** to **{today.strftime('%Y-%m-%d')}**")
    else:
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            start_date_input = st.date_input("Start Date", value=today - timedelta(days=7), key="911_start")
        with col_d2:
            end_date_input = st.date_input("End Date", value=today, key="911_end")
        
        last_friday = datetime.combine(start_date_input, datetime.min.time())
        today = datetime.combine(end_date_input, datetime.max.time())

    st.markdown("---")

    def run_pipeline(uploaded_file, section_choice):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        try:
            df, meta = pyreadstat.read_sav(tmp_path)
            
            if 'V9999' in df.columns:
                df_filtered = df[df['V9999'] == 1].copy()
            else:
                df_filtered = df.copy()

            if 'STIME' in df_filtered.columns:
                df_filtered['STIME_CLEAN'] = df_filtered['STIME'].astype(str).str.strip().str[:8]
                def parse_stime_date(x):
                    try:
                        return datetime.strptime(x, "%Y%m%d")
                    except:
                        return None
                df_filtered['STIME_DATE'] = df_filtered['STIME_CLEAN'].apply(parse_stime_date)
                
                df_filtered = df_filtered[
                    (df_filtered['STIME_DATE'] >= last_friday) & 
                    (df_filtered['STIME_DATE'] <= today)
                ].copy()

            if df_filtered.empty:
                st.warning(f"⚠️ No records found matching the criteria for {section_choice}.")
                return None

            if 'INTNR' not in df_filtered.columns:
                df_filtered['INTNR'] = range(1, len(df_filtered) + 1)
                
            valid_intnr = df_filtered['INTNR'] > 0

            df_filtered.loc[valid_intnr, 'PARENT_TYPE'] = "Juristic"
            df_filtered.loc[valid_intnr, 'WAVE'] = "22"
            if 'V80116' in df_filtered.columns:
                df_filtered.loc[valid_intnr, 'CLIENT_UCN'] = df_filtered['V80116']
            df_filtered.loc[valid_intnr, 'CLIENT_TYPE'] = "Full Client"
            df_filtered.loc[valid_intnr, 'COMPANY_CODE'] = "15"
            df_filtered.loc[valid_intnr, 'CASE_SUBJECT'] = "Coverage"
            df_filtered.loc[valid_intnr, 'REQUEST_CATEGORY'] = "Care"
            df_filtered.loc[valid_intnr, 'TOPIC'] = "Complaints"

            tq13_open_clean = df_filtered['TQ13_OPEN'].fillna('').astype(str) if 'TQ13_OPEN' in df_filtered.columns else ""
            df_filtered.loc[valid_intnr, 'CASE_DESCRIPTION'] = "Improvement Area: " + tq13_open_clean

            df_filtered.loc[valid_intnr, 'CAMPAIGN'] = "CMP-01522-S7K7N3"
            df_filtered.loc[valid_intnr, 'ORIGIN'] = "Web"
            df_filtered.loc[valid_intnr, 'OWNER'] = r"FNBJNB01\Web"
            
            sub_region_col = 'V8013' if section_choice == 'PUBSC' else 'V13290'
            segment_col = 'V13290' if section_choice == 'PUBSC' else 'V44011'

            mapping_pairs = [
                ('PRIM_OFCR_IND', 'V8026'), ('OFFICER_NAME_AND_SURNAME', 'V8016'), 
                ('BUSINESS_NAME', 'V56011'), ('REGIONS', 'V12290'), 
                ('SUB_REGIONS', sub_region_col), ('SEGMENT', segment_col)
            ]

            for col_target, col_src in mapping_pairs:
                if col_src in df_filtered.columns:
                    df_filtered.loc[valid_intnr, col_target] = df_filtered[col_src]

            def calculate_nps_bucket(score):
                if pd.isna(score): return None
                if score < 7: return "NPS - Detractor"
                elif score in [7, 8]: return "NPS - Passive"
                elif score > 8: return "NPS - Promoter"
                return None

            if 'Q14_1' in df_filtered.columns:
                df_filtered['FNB_NPS'] = df_filtered['Q14_1'].apply(calculate_nps_bucket)
                df_filtered['FNB_NPS_OPEN_ENDED'] = df_filtered.apply(lambda r: r['TQ14_1_OPEN'] if 'TQ14_1_OPEN' in df_filtered.columns and pd.notna(r.get('TQ14_1_OPEN')) and r['Q14_1'] < 7 else None, axis=1)
            if 'Q14_2' in df_filtered.columns:
                df_filtered['RM_BM_NPS'] = df_filtered['Q14_2'].apply(calculate_nps_bucket)
                df_filtered['RM_BM_NPS_OPEN_ENDED'] = df_filtered.apply(lambda r: r['TQ14_2_OPEN'] if 'TQ14_2_OPEN' in df_filtered.columns and pd.notna(r.get('TQ14_2_OPEN')) and r['Q14_2'] < 7 else None, axis=1)

            if section_choice == 'PUBSC':
                bank_columns = [('Q16_1_1', 'Absa'), ('Q16_1_2', 'Investec'), ('Q16_1_3', 'Nedbank'), ('Q16_1_4', 'Standard Bank'), ('Q16_1_5', 'Capitec'), ('Q16_1_6', 'Refused')]
                loop_cols = ['TQ16_1C6', 'TQ16_1C7', 'TQ16_1C8']
            else:
                bank_columns = [('Q16_1_1', 'Absa'), ('Q16_1_2', 'Capitec'), ('Q16_1_3', 'Investec'), ('Q16_1_4', 'Mercantile'), ('Q16_1_5', 'Nedbank'), ('Q16_1_6', 'Sasfin'), ('Q16_1_7', 'Standard Bank')]
                loop_cols = ['TQ16_1C8', 'TQ16_1C9', 'TQ16_1C10']
            
            switch_compiled = []
            for idx, row in df_filtered.iterrows():
                matched_banks = []
                for col_flag, bank_label in bank_columns:
                    if col_flag in df_filtered.columns and row.get(col_flag) == 1:
                        matched_banks.append(bank_label)
                for loop_col in loop_cols:
                    if loop_col in df_filtered.columns and pd.notna(row.get(loop_col)) and str(row[loop_col]).strip() != '':
                        matched_banks.append(str(row[loop_col]).strip())
                switch_compiled.append(",".join(matched_banks))
                
            df_filtered['WOULD_CONSIDER_SWITCH_TO'] = switch_compiled
            if 'TQ16_OPEN' in df_filtered.columns:
                df_filtered.loc[valid_intnr, 'REASON'] = df_filtered['TQ16_OPEN']

            if 'STIME_CLEAN' in df_filtered.columns:
                df_filtered['NYEAR'] = df_filtered['STIME_CLEAN'].str[:4]
                df_filtered['NMONTH'] = df_filtered['STIME_CLEAN'].str[4:6]
                df_filtered['NDAY'] = df_filtered['STIME_CLEAN'].str[6:8]
                df_filtered['RECORDED_DATE'] = df_filtered['NYEAR'] + "-" + df_filtered['NMONTH'] + "-" + df_filtered['NDAY']

            df_filtered['Qualifier'] = "Not Priority"
            if 'Q14_2' in df_filtered.columns:
                df_filtered.loc[df_filtered['Q14_2'] < 7, 'Qualifier'] = "Priority"
            if 'Q16' in df_filtered.columns:
                df_filtered.loc[df_filtered['Q16'] == 1, 'Qualifier'] = "Priority"

            case_desc_upper = df_filtered['CASE_DESCRIPTION'].fillna('').str.upper() if 'CASE_DESCRIPTION' in df_filtered.columns else pd.Series([""]*len(df_filtered))
            people_keywords = ['BM', 'BUSINESS MANAGER', 'BUSINESS BANKERS', 'PRIVATE BANKER', 'RM', 'RELATIONSHIP MANAGER', 'STAFF', 'CLIENTS']
            process_keywords = ['SYSTEM', 'PROCESS', 'SERVICE', 'DELAY', 'QUERY', 'ACCESS', 'APP']
            product_keywords = ['FEE', 'CHARGES', 'LOAN', 'ACCOUNT', 'INVESTMENT', 'CARD']
            none_keywords = ['NO IMPROVEMENT', 'NONE', 'SATISFIED', 'ALL GOOD', 'N/A']

            def contains_keywords(text, kw_list):
                return 1 if any(kw in text for kw in kw_list) else 0

            df_filtered['People_1'] = case_desc_upper.apply(lambda x: contains_keywords(x, people_keywords))
            df_filtered['PROCESS_1'] = case_desc_upper.apply(lambda x: contains_keywords(x, process_keywords))
            df_filtered['PRODUCT_1'] = case_desc_upper.apply(lambda x: contains_keywords(x, product_keywords))
            df_filtered['NONE_OVERRIDE'] = case_desc_upper.apply(lambda x: contains_keywords(x, none_keywords))

            def apply_triple_p_logic(row):
                if row.get('NONE_OVERRIDE') == 1: return "NONE"
                p, pp, pr = row.get('PRODUCT_1') == 1, row.get('People_1') == 1, row.get('PROCESS_1') == 1
                if pp and pr and p: return "ALL"
                if p and pr: return "PRODUCT & PROCESS"
                if p and pp: return "PRODUCT & PEOPLE"
                if pp and pr: return "PEOPLE & PROCESS"
                if p: return "PRODUCT ONLY"
                if pr: return "PROCESS ONLY"
                if pp: return "PEOPLE ONLY"
                return ""

            df_filtered['PRODUCT_PEOPLE_PROCESS'] = df_filtered.apply(apply_triple_p_logic, axis=1)
            df_filtered = df_filtered.sort_values(by='INTNR', ascending=True).copy()

            verbatim_cols = ['CASE_DESCRIPTION', 'FNB_NPS_OPEN_ENDED', 'RM_BM_NPS_OPEN_ENDED', 'WOULD_CONSIDER_SWITCH_TO', 'REASON']
            for col in verbatim_cols:
                if col in df_filtered.columns:
                    df_filtered[col] = df_filtered[col].fillna('').astype(str).str.replace(',', '~', regex=False)

            df_priority = df_filtered[df_filtered['Qualifier'] == "Priority"].copy()

            keep_columns = [
                'PARENT_TYPE', 'WAVE', 'CLIENT_UCN', 'CLIENT_TYPE', 'COMPANY_CODE', 'CASE_SUBJECT',
                'REQUEST_CATEGORY', 'TOPIC', 'CASE_DESCRIPTION', 'PRODUCT_PEOPLE_PROCESS', 'CAMPAIGN',
                'ORIGIN', 'OWNER', 'PRIM_OFCR_IND', 'OFFICER_NAME_AND_SURNAME', 'BUSINESS_NAME',
                'REGIONS', 'SUB_REGIONS', 'SEGMENT', 'FNB_NPS', 'FNB_NPS_OPEN_ENDED', 'RM_BM_NPS',
                'RM_BM_NPS_OPEN_ENDED', 'WOULD_CONSIDER_SWITCH_TO', 'REASON', 'RECORDED_DATE'
            ]

            for c in keep_columns:
                if c not in df_priority.columns:
                    df_priority[c] = ""

            base_priority_data = df_priority[keep_columns].copy()
            run_date_file = today.strftime("%Y_%m_%d")

            if section_choice == "Growth": prefix = "Business_Client_911"
            elif section_choice == "R10Mil": prefix = "Enterprise_Client_911"
            else: prefix = "PUBSC_Client_911"

            f1_data = base_priority_data[(base_priority_data.get('FNB_NPS') == "NPS - Detractor") | (base_priority_data.get('RM_BM_NPS') == "NPS - Detractor")]
            f2_data = f1_data.drop(columns=['PRODUCT_PEOPLE_PROCESS'], errors='ignore')
            f3_data = base_priority_data
            f4_data = base_priority_data.drop(columns=['PRODUCT_PEOPLE_PROCESS'], errors='ignore')

            return {
                "f1": (f1_data.to_csv(sep='|', index=False, encoding='utf-8-sig').encode('utf-8-sig'), f"{prefix}_NPS_D_classification_{run_date_file}.csv"),
                "f2": (f2_data.to_csv(sep='|', index=False, encoding='utf-8-sig').encode('utf-8-sig'), f"{prefix}_NPS_D_NO_classification_{run_date_file}.csv"),
                "f3": (f3_data.to_csv(sep='|', index=False, encoding='utf-8-sig').encode('utf-8-sig'), f"{prefix}_NPS_D_S_classification_{run_date_file}.csv"),
                "f4": (f4_data.to_csv(sep='|', index=False, encoding='utf-8-sig').encode('utf-8-sig'), f"{prefix}_NPS_D_S_NO_classification_{run_date_file}.csv"),
                "count": len(df_filtered)
            }
        except Exception as e:
            st.error(f"❌ Error in {section_choice}: {e}")
            return None
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    st.subheader("⚡ Pipeline Execution Control Room")
    col1, col2, col3 = st.columns(3)

    with col1:
        with st.container(border=True):
            st.markdown("### 🟢 Growth Section")
            st.caption("Target: Business Client Pipeline")
            file_growth = st.file_uploader("Upload GROW SAV (.sav)", type=["sav"], key="growth_file")
            
            if st.button("▶ Run Growth Stage", key="btn_growth", type="primary", use_container_width=True):
                if file_growth is None:
                    st.error("Upload a SAV file first.")
                else:
                    with st.spinner("Processing Growth..."):
                        res = run_pipeline(file_growth, "Growth")
                        if res:
                            st.success(f"Processed {res['count']} records!")
                            st.download_button("📥 Output 1", res['f1'][0], file_name=res['f1'][1], mime="text/csv", key="g1")
                            st.download_button("📥 Output 2", res['f2'][0], file_name=res['f2'][1], mime="text/csv", key="g2")
                            st.download_button("📥 Output 3", res['f3'][0], file_name=res['f3'][1], mime="text/csv", key="g3")
                            st.download_button("📥 Output 4", res['f4'][0], file_name=res['f4'][1], mime="text/csv", key="g4")

    with col2:
        with st.container(border=True):
            st.markdown("### 🔵 R10Mil Section")
            st.caption("Target: Enterprise Client Pipeline")
            file_r10 = st.file_uploader("Upload RMW SAV (.sav)", type=["sav"], key="r10_file")
            
            if st.button("▶ Run R10Mil Stage", key="btn_r10", type="primary", use_container_width=True):
                if file_r10 is None:
                    st.error("Upload a SAV file first.")
                else:
                    with st.spinner("Processing R10Mil..."):
                        res = run_pipeline(file_r10, "R10Mil")
                        if res:
                            st.success(f"Processed {res['count']} records!")
                            st.download_button("📥 Output 1", res['f1'][0], file_name=res['f1'][1], mime="text/csv", key="r1")
                            st.download_button("📥 Output 2", res['f2'][0], file_name=res['f2'][1], mime="text/csv", key="r2")
                            st.download_button("📥 Output 3", res['f3'][0], file_name=res['f3'][1], mime="text/csv", key="r3")
                            st.download_button("📥 Output 4", res['f4'][0], file_name=res['f4'][1], mime="text/csv", key="r4")

    with col3:
        with st.container(border=True):
            st.markdown("### 🟠 PUBSC Section")
            st.caption("Target: Public Sector Pipeline")
            file_pub = st.file_uploader("Upload PUBW SAV (.sav)", type=["sav"], key="pub_file")
            
            if st.button("▶ Run PUBSC Stage", key="btn_pub", type="primary", use_container_width=True):
                if file_pub is None:
                    st.error("Upload a SAV file first.")
                else:
                    with st.spinner("Processing PUBSC..."):
                        res = run_pipeline(file_pub, "PUBSC")
                        if res:
                            st.success(f"Processed {res['count']} records!")
                            st.download_button("📥 Output 1", res['f1'][0], file_name=res['f1'][1], mime="text/csv", key="p1")
                            st.download_button("📥 Output 2", res['f2'][0], file_name=res['f2'][1], mime="text/csv", key="p2")
                            st.download_button("📥 Output 3", res['f3'][0], file_name=res['f3'][1], mime="text/csv", key="p3")
                            st.download_button("📥 Output 4", res['f4'][0], file_name=res['f4'][1], mime="text/csv", key="p4")

# ==========================================================================
# ==========================================================================
# TAB 3: BM/RM NPS PORTFOLIO GENERATOR
# ==========================================================================
# ==========================================================================
with tab3:
    st.markdown("### 📈 BM/RM NPS Portfolio Generator")
    st.markdown("Upload your master SPSS (`.sav`) data file below, select your wave preferences and portfolio filter, then click **Run Processing** to generate your reports.")

    if "reports_ready" not in st.session_state:
        st.session_state.reports_ready = False
    if "report_files" not in st.session_state:
        st.session_state.report_files = {}

    uploaded_file_tab3 = st.file_uploader("Upload Master SPSS Data File (.sav)", type=["sav"], key="tab3_master_file")

    portfolio_mode = st.selectbox(
        "Select Portfolio Filter Mode:",
        [
            "Generate All (Combined, Growth, and R10M Separately)",
            "Combined (Growth & R10M)",
            "Growth Only",
            "R10M Only"
        ],
        key="tab3_portfolio_mode"
    )

    filter_option = st.radio("Select Wave Filter Option:", ["All Waves", "Custom Range (e.g., Wave 1 to 10)", "Specific Waves List"], key="tab3_filter_option")

    selected_waves_filter = 'ALL'

    if filter_option == "Custom Range (e.g., Wave 1 to 10)":
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            start_w = st.number_input("Start Wave Number", min_value=1, max_value=30, value=1, key="tab3_start_w")
        with col_w2:
            end_w = st.number_input("End Wave Number", min_value=1, max_value=30, value=10, key="tab3_end_w")
        selected_waves_filter = [f'Wave {i}' for i in range(int(start_w), int(end_w) + 1)]

    elif filter_option == "Specific Waves List":
        waves_input = st.text_input("Enter waves separated by commas:", "Wave 20, Wave 21, Wave 22", key="tab3_waves_input")
        selected_waves_filter = [w.strip() for w in waves_input.split(',')]

    def generate_report_bytes(df_subset, prefix_label):
        is_combined = (prefix_label == "Combined")
        
        excel_name = f"Overall NPS Rating per BM RM Portfolio_{prefix_label}.xlsx"
        sav_name = f"Project Star_NPS_Streamlined_{prefix_label}.sav"

        temp_excel = f"temp_{prefix_label}.xlsx"
        temp_sav = f"temp_{prefix_label}.sav"

        pyreadstat.write_sav(df_subset, temp_sav)
        with open(temp_sav, "rb") as f:
            sav_bytes = f.read()

        with pd.ExcelWriter(temp_excel, engine='openpyxl') as writer:
            df_subset.to_excel(writer, sheet_name='data', index=False)

        wb = openpyxl.load_workbook(temp_excel)
        ws_toc = wb.create_sheet(title='TOC', index=0)
        ws_nps = wb.create_sheet(title='NPS', index=1)
        ws_data = wb['data']
        ws_data.sheet_state = 'hidden'

        TEAL_HEADER_FILL = PatternFill(start_color="00A3AD", end_color="00A3AD", fill_type="solid")
        LIGHT_TEAL_FILL = PatternFill(start_color="D9F2F4", end_color="D9F2F4", fill_type="solid")
        ORANGE_HEADER_FILL = PatternFill(start_color="F58220", end_color="F58220", fill_type="solid")
        LIGHT_ORANGE_FILL = PatternFill(start_color="FDF3EC", end_color="FDF3EC", fill_type="solid")
        BANNER_FILL = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
        ZEBRA_FILL = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")

        WHITE_BOLD_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        BOLD_FONT = Font(name="Calibri", size=11, bold=True)
        REGULAR_FONT = Font(name="Calibri", size=11)
        THIN_BORDER = Border(left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'), top=Side(style='thin', color='DDDDDD'), bottom=Side(style='thin', color='DDDDDD'))

        ws_toc.cell(row=1, column=2, value=f"FNB Customer Satisfaction Study 2026: NPS Rating per BM/RM Portfolio ({prefix_label})")
        ws_toc.cell(row=1, column=2).fill = TEAL_HEADER_FILL
        ws_toc.cell(row=1, column=2).font = TITLE_FONT
        ws_toc.cell(row=1, column=2).alignment = Alignment(horizontal="center", vertical="center")

        ws_toc.append(["", "Select Wave:", "All Waves"])
        ws_toc.cell(row=2, column=2).font = BOLD_FONT
        ws_toc.cell(row=2, column=2).alignment = Alignment(horizontal="right")
        ws_toc.cell(row=2, column=3).fill = LIGHT_TEAL_FILL
        ws_toc.cell(row=2, column=3).border = THIN_BORDER

        def wave_sort_key(val):
            match = re.search(r'\d+', str(val))
            return int(match.group()) if match else 0

        raw_waves = df_subset['WAVE'].dropna().unique() if 'WAVE' in df_subset.columns else []
        available_waves = sorted(raw_waves, key=wave_sort_key)
        wave_list_str = '"All Waves,' + ','.join([str(w) for w in available_waves]) + '"'
        wave_dv = DataValidation(type="list", formula1=wave_list_str, allow_blank=False)
        ws_toc.add_data_validation(wave_dv)
        wave_dv.add(ws_toc['C2'])

        if is_combined:
            ws_toc.append(["", "Select Type:", "All Types"])
            ws_toc.cell(row=3, column=2).font = BOLD_FONT
            ws_toc.cell(row=3, column=2).alignment = Alignment(horizontal="right")
            ws_toc.cell(row=3, column=3).fill = LIGHT_TEAL_FILL
            ws_toc.cell(row=3, column=3).border = THIN_BORDER

            available_types = sorted(df_subset['TYPE'].dropna().unique()) if 'TYPE' in df_subset.columns else []
            type_list_str = '"All Types,' + ','.join([str(t) for t in available_types]) + '"'
            type_dv = DataValidation(type="list", formula1=type_list_str, allow_blank=False)
            ws_toc.add_data_validation(type_dv)
            type_dv.add(ws_toc['C3'])
            toc_header_row = 5
        else:
            toc_header_row = 4

        ws_toc.append([])
        ws_toc.append(["", "Table of Contents", ""])
        ws_toc.cell(row=toc_header_row, column=2).font = Font(name="Calibri", size=12, bold=True, color="F58220")

        toc_entries = []

        def write_block(ws, title_text, officers_subset):
            start_row = ws.max_row + 1 if ws.max_row > 1 else 1
            if ws.max_row == 1 and ws['A1'].value is None:
                start_row = 1
            ws.append([title_text] + [""] * 13)
            ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=14)
            banner_cell = ws.cell(row=ws.max_row, column=1)
            banner_cell.fill = BANNER_FILL
            banner_cell.font = WHITE_BOLD_FONT
            banner_cell.alignment = Alignment(horizontal="center", vertical="center")
            
            ws.append([""] * 14)
            back_row = ws.max_row
            back_cell = ws.cell(row=back_row, column=1, value="Back to TOC")
            back_cell.hyperlink = f"#TOC!B{toc_header_row}"
            back_cell.font = Font(name="Calibri", size=11, color="00A3AD", underline="single")
            
            h1_row = ws.max_row + 1
            ws.append(["BM/RM Name", "Officer Code", "Sub-Region", "NPS Segments_BM", "", "", "", "Count", "NPS Segments_FNB Business", "", "", "", "Count", "Officer Code"])
            h2_row = ws.max_row + 1
            ws.append(["", "", "", "NPS Score", "Detractors", "Passives", "Promoters", "", "NPS Score", "Detractors", "Passives", "Promoters", "", ""])
            
            for c in [1, 2, 3, 8, 14]:
                ws.cell(row=h1_row, column=c).fill = BANNER_FILL
                ws.cell(row=h1_row, column=c).font = WHITE_BOLD_FONT
                ws.cell(row=h1_row, column=c).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=h2_row, column=c).fill = BANNER_FILL
                ws.cell(row=h2_row, column=c).font = WHITE_BOLD_FONT
                ws.cell(row=h2_row, column=c).alignment = Alignment(horizontal="center", vertical="center")

            for c in range(4, 8):
                ws.cell(row=h1_row, column=c).fill = TEAL_HEADER_FILL
                ws.cell(row=h1_row, column=c).font = WHITE_BOLD_FONT
                ws.cell(row=h1_row, column=c).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=h2_row, column=c).fill = LIGHT_TEAL_FILL
                ws.cell(row=h2_row, column=c).font = BOLD_FONT
                ws.cell(row=h2_row, column=c).alignment = Alignment(horizontal="center", vertical="center")

            for c in range(9, 14):
                ws.cell(row=h1_row, column=c).fill = ORANGE_HEADER_FILL
                ws.cell(row=h1_row, column=c).font = WHITE_BOLD_FONT
                ws.cell(row=h1_row, column=c).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=h2_row, column=c).fill = LIGHT_ORANGE_FILL
                ws.cell(row=h2_row, column=c).font = BOLD_FONT
                ws.cell(row=h2_row, column=c).alignment = Alignment(horizontal="center", vertical="center")

            row_counter = 0
            for _, row in officers_subset.iterrows():
                formula_row = ws.max_row + 1
                row_counter += 1
                rm_name = row['OFFICER_NAME']
                prim_code = row['PRIM_OFCR_IND']
                col_c_val = str(title_text) if title_text != "TOTAL" else str(row['SUBREG'])
                office_code = row['OFFICE_CODE']
                
                if is_combined:
                    bm_count_formula = f'=IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$H:$H, ">0", data!$M:$M, N{formula_row}), IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$H:$H, ">0", data!$M:$M, N{formula_row}), IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, ">0", data!$M:$M, N{formula_row}), COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$H:$H, ">0", data!$M:$M, N{formula_row}))))'
                    fnb_count_formula = f'=IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$I:$I, ">0", data!$M:$M, N{formula_row}), IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$I:$I, ">0", data!$M:$M, N{formula_row}), IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, ">0", data!$M:$M, N{formula_row}), COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$I:$I, ">0", data!$M:$M, N{formula_row}))))'
                    
                    bm_det_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$H:$H, 1, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$H:$H, 1, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, 1, data!$M:$M, N{formula_row})/H{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$H:$H, 1, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'
                    bm_pas_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$H:$H, 2, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$H:$H, 2, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, 2, data!$M:$M, N{formula_row})/H{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$H:$H, 2, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'
                    bm_pro_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$H:$H, 3, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$H:$H, 3, data!$M:$M, N{formula_row})/H{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, 3, data!$M:$M, N{formula_row})/H{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$H:$H, 3, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'
                    
                    fnb_det_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$I:$I, 1, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$I:$I, 1, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, 1, data!$M:$M, N{formula_row})/M{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$I:$I, 1, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'
                    fnb_pas_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$I:$I, 2, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$I:$I, 2, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, 2, data!$M:$M, N{formula_row})/M{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$I:$I, 2, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'
                    fnb_pro_formula = f'=IFERROR(IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIFS(data!$I:$I, 3, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$I:$I, 3, data!$M:$M, N{formula_row})/M{formula_row}, IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, 3, data!$M:$M, N{formula_row})/M{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$I:$I, 3, data!$M:$M, N{formula_row})/H{formula_row}))), 0)'
                else:
                    bm_count_formula = f'=IF(TOC!$C$2="All Waves", COUNTIFS(data!$H:$H, ">0", data!$M:$M, N{formula_row}), COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, ">0", data!$M:$M, N{formula_row}))'
                    fnb_count_formula = f'=IF(TOC!$C$2="All Waves", COUNTIFS(data!$I:$I, ">0", data!$M:$M, N{formula_row}), COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, ">0", data!$M:$M, N{formula_row}))'
                    
                    bm_det_formula = f'=IFERROR(IF(TOC!$C$2="All Waves", COUNTIFS(data!$H:$H, 1, data!$M:$M, N{formula_row})/H{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, 1, data!$M:$M, N{formula_row})/H{formula_row}), 0)'
                    bm_pas_formula = f'=IFERROR(IF(TOC!$C$2="All Waves", COUNTIFS(data!$H:$H, 2, data!$M:$M, N{formula_row})/H{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, 2, data!$M:$M, N{formula_row})/H{formula_row}), 0)'
                    bm_pro_formula = f'=IFERROR(IF(TOC!$C$2="All Waves", COUNTIFS(data!$H:$H, 3, data!$M:$M, N{formula_row})/H{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$H:$H, 3, data!$M:$M, N{formula_row})/H{formula_row}), 0)'
                    
                    fnb_det_formula = f'=IFERROR(IF(TOC!$C$2="All Waves", COUNTIFS(data!$I:$I, 1, data!$M:$M, N{formula_row})/M{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, 1, data!$M:$M, N{formula_row})/M{formula_row}), 0)'
                    fnb_pas_formula = f'=IFERROR(IF(TOC!$C$2="All Waves", COUNTIFS(data!$I:$I, 2, data!$M:$M, N{formula_row})/M{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, 2, data!$M:$M, N{formula_row})/M{formula_row}), 0)'
                    fnb_pro_formula = f'=IFERROR(IF(TOC!$C$2="All Waves", COUNTIFS(data!$I:$I, 3, data!$M:$M, N{formula_row})/M{formula_row}, COUNTIFS(data!$D:$D, TOC!$C$2, data!$I:$I, 3, data!$M:$M, N{formula_row})/M{formula_row}), 0)'

                ws.append([
                    rm_name, prim_code, col_c_val,
                    f'=IF(H{formula_row}>0, SUM(G{formula_row}-E{formula_row})*100, "")',
                    bm_det_formula, bm_pas_formula, bm_pro_formula, bm_count_formula,
                    f'=IF(M{formula_row}>0, SUM(L{formula_row}-J{formula_row})*100, "")',
                    fnb_det_formula, fnb_pas_formula, fnb_pro_formula, fnb_count_formula,
                    office_code
                ])
                
                ws[f'D{formula_row}'].number_format = '0'
                ws[f'E{formula_row}'].number_format = '0%'
                ws[f'F{formula_row}'].number_format = '0%'
                ws[f'G{formula_row}'].number_format = '0%'
                ws[f'I{formula_row}'].number_format = '0'
                ws[f'J{formula_row}'].number_format = '0%'
                ws[f'K{formula_row}'].number_format = '0%'
                ws[f'L{formula_row}'].number_format = '0%'

                for c in range(1, 15):
                    cell = ws.cell(row=formula_row, column=c)
                    cell.border = THIN_BORDER
                    cell.font = REGULAR_FONT
                    if row_counter % 2 == 0:
                        cell.fill = ZEBRA_FILL

            ws.append([])
            return start_row

        unique_officers_all = df_subset[['OFFICER_NAME', 'PRIM_OFCR_IND', 'SUBREG', 'OFFICE_CODE']].drop_duplicates().sort_values(by='OFFICE_CODE')
        total_start_row = write_block(ws_nps, "TOTAL", unique_officers_all)
        toc_entries.append(("Officer Name and Code by NPS Banner", total_start_row, "TOTAL"))

        if 'SUBREG' in df_subset.columns:
            for subreg in sorted(df_subset['SUBREG'].dropna().unique()):
                df_subreg = df_subset[df_subset['SUBREG'] == subreg]
                unique_officers_subreg = df_subreg[['OFFICER_NAME', 'PRIM_OFCR_IND', 'SUBREG', 'OFFICE_CODE']].drop_duplicates().sort_values(by='OFFICE_CODE')
                subreg_start_row = write_block(ws_nps, str(subreg), unique_officers_subreg)
                toc_entries.append((f"Officer Name and Code by NPS Banner {subreg}", subreg_start_row, str(subreg)))

        for label, nps_row_num, subreg_filter in toc_entries:
            row_idx = ws_toc.max_row + 1
            if is_combined:
                if subreg_filter == "TOTAL":
                    filter_formula = f'=CONCATENATE("Filter: Wave ", TOC!$C$2, ", Type ", TOC!$C$3, ", base n =", IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNT(data!$A:$A), IF(TOC!$C$2="All Waves", COUNTIF(data!$C:$C, TOC!$C$3), IF(TOC!$C$3="All Types", COUNTIF(data!$D:$D, TOC!$C$2), COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3)))))'
                else:
                    filter_formula = f'=CONCATENATE("Filter: Wave ", TOC!$C$2, ", Type ", TOC!$C$3, ", base n =", IF(AND(TOC!$C$2="All Waves", TOC!$C$3="All Types"), COUNTIF(data!$F:$F, "{subreg_filter}"), IF(TOC!$C$2="All Waves", COUNTIFS(data!$C:$C, TOC!$C$3, data!$F:$F, "{subreg_filter}"), IF(TOC!$C$3="All Types", COUNTIFS(data!$D:$D, TOC!$C$2, data!$F:$F, "{subreg_filter}"), COUNTIFS(data!$D:$D, TOC!$C$2, data!$C:$C, TOC!$C$3, data!$F:$F, "{subreg_filter}")))))'
            else:
                if subreg_filter == "TOTAL":
                    filter_formula = f'=CONCATENATE("Filter: Wave ", TOC!$C$2, ", base n =", IF(TOC!$C$2="All Waves", COUNT(data!$A:$A), COUNTIF(data!$D:$D, TOC!$C$2)))'
                else:
                    filter_formula = f'=CONCATENATE("Filter: Wave ", TOC!$C$2, ", base n =", IF(TOC!$C$2="All Waves", COUNTIF(data!$F:$F, "{subreg_filter}"), COUNTIFS(data!$D:$D, TOC!$C$2, data!$F:$F, "{subreg_filter}")))'

            ws_toc.cell(row=row_idx, column=1, value="NPS").font = BOLD_FONT
            ws_toc.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center")
            
            link_cell = ws_toc.cell(row=row_idx, column=2, value=label)
            link_cell.hyperlink = f"#NPS!A{nps_row_num}"
            link_cell.font = Font(name="Calibri", size=11, color="F58220", underline="single")
            
            ws_toc.cell(row=row_idx, column=3, value=filter_formula).font = REGULAR_FONT

        for ws in wb.worksheets:
            if ws.title == 'TOC':
                ws.column_dimensions['A'].width = 3.67
                ws.column_dimensions['B'].width = 78.22
                ws.column_dimensions['C'].width = 43.11
            elif ws.title == 'NPS':
                ws.column_dimensions['A'].width = 27.11
                ws.column_dimensions['B'].width = 10.67
                ws.column_dimensions['C'].width = 20.11
                ws.column_dimensions['D'].width = 16.44
                ws.column_dimensions['I'].width = 24.67
                for col in ['E', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'N']:
                    ws.column_dimensions[col].width = 11

        wb.save(temp_excel)
        with open(temp_excel, "rb") as f:
            excel_bytes = f.read()

        return excel_bytes, excel_name, sav_bytes, sav_name

    if st.button("🚀 Run Processing & Generate Reports", type="primary", key="tab3_run_processing") or st.session_state.reports_ready:
        if uploaded_file_tab3 is None:
            st.error("Please upload a `.sav` file first!")
            st.session_state.reports_ready = False
        else:
            if not st.session_state.reports_ready:
                with st.spinner("Processing data and building formatted reports..."):
                    temp_src_path = "temp_input.sav"
                    with open(temp_src_path, "wb") as f:
                        f.write(uploaded_file_tab3.getbuffer())

                    df_raw, meta = pyreadstat.read_sav(temp_src_path, apply_value_formats=False)
                    df_raw.columns = [str(col).strip().upper() for col in df_raw.columns]

                    df_lbl, _ = pyreadstat.read_sav(temp_src_path, apply_value_formats=True)
                    df_lbl.columns = [str(col).strip().upper() for col in df_lbl.columns]

                    target_columns_mapping = {
                        'UNIQUEID': ['UNIQUEID', 'ID'],
                        'RUID': ['RUID'],
                        'TYPE': ['TYPE'],
                        'WAVE': ['WAVE'],
                        'REGION': ['REGION'],
                        'SUBREG': ['SUBREG'],
                        'SEGMENT': ['SEGMENT'],
                        'RM_NPS1': ['RM_NPS1', 'BM_NPS1', 'BMNPS01'],
                        'FNB_NPS1': ['FNB_NPS1', 'FNBNPS01'],
                        'PRIM_OFCR_IND': ['PRIM_OFCR_IND', 'PRIM_OFC'],
                        'OFFICER_NAME': ['OFFICER_NAME', 'OFFICER_M', 'OFFICER_NAME_']
                    }

                    upper_to_orig = {str(c).strip().upper(): c for c in df_raw.columns}
                    rename_map = {}
                    for target, candidates in target_columns_mapping.items():
                        for cand in candidates:
                            if cand.upper() in upper_to_orig:
                                rename_map[upper_to_orig[cand.upper()]] = target
                                break

                    orig_type_col = [k for k, v in rename_map.items() if v == 'TYPE'][0]
                    raw_type_numeric = pd.to_numeric(df_raw[orig_type_col], errors='coerce')
                    valid_mask = raw_type_numeric.isin([1, 2])

                    df_raw = df_raw[valid_mask].reset_index(drop=True)
                    df_lbl = df_lbl[valid_mask].reset_index(drop=True)
                    raw_type_numeric = raw_type_numeric[valid_mask].reset_index(drop=True)

                    df_base = pd.DataFrame()
                    for orig_col, target in rename_map.items():
                        if target in ['WAVE', 'REGION', 'SUBREG', 'SEGMENT'] and orig_col in df_lbl.columns:
                            df_base[target] = df_lbl[orig_col]
                        else:
                            df_base[target] = df_raw[orig_col]

                    if 'TYPE' in df_lbl.columns:
                        df_base['TYPE'] = df_lbl[orig_type_col]

                    if selected_waves_filter != 'ALL' and 'WAVE' in df_base.columns:
                        df_base = df_base[df_base['WAVE'].isin(selected_waves_filter)].copy()

                    def clean_officer_name(name):
                        if pd.isna(name):
                            return name
                        s = str(name).strip()
                        s = re.sub(r'\s*\(.*?\)', '', s)
                        return ' '.join(s.split())

                    if 'OFFICER_NAME' in df_base.columns:
                        df_base['OFFICER_NAME'] = df_base['OFFICER_NAME'].apply(clean_officer_name)

                    if 'OFFICER_NAME' in df_base.columns and 'PRIM_OFCR_IND' in df_base.columns:
                        df_base['OFFICER_NAME'] = df_base['OFFICER_NAME'].astype(str).str.strip()
                        df_base['PRIM_OFCR_IND'] = df_base['PRIM_OFCR_IND'].astype(str).str.strip()
                        df_base['OFFICER_NAME2'] = df_base['OFFICER_NAME'] + df_base['PRIM_OFCR_IND']
                        df_base = df_base.sort_values(by='OFFICER_NAME2').reset_index(drop=True)
                        unique_names = df_base['OFFICER_NAME2'].unique()
                        name_to_code = {name: idx + 1 for idx, name in enumerate(unique_names)}
                        df_base['OFFICE_CODE'] = df_base['OFFICER_NAME2'].map(name_to_code)

                    runs = []
                    if 'TYPE' in df_raw.columns:
                        subset_type_numeric = pd.to_numeric(df_raw[orig_type_col], errors='coerce')

                        mask_comb = subset_type_numeric.isin([1, 2])
                        mask_grow = subset_type_numeric == 1
                        mask_r10m = subset_type_numeric == 2

                        df_comb = df_base[mask_comb].copy()
                        df_grow = df_base[mask_grow].copy()
                        df_r10m = df_base[mask_r10m].copy()

                        if portfolio_mode == "Generate All (Combined, Growth, and R10M Separately)":
                            runs = [("Combined", df_comb), ("Growth", df_grow), ("R10M", df_r10m)]
                        elif portfolio_mode == "Combined (Growth & R10M)":
                            runs = [("Combined", df_comb)]
                        elif portfolio_mode == "Growth Only":
                            runs = [("Growth", df_grow)]
                        elif portfolio_mode == "R10M Only":
                            runs = [("R10M", df_r10m)]

                    st.session_state.report_files = {}
                    for label, subset_df in runs:
                        ex_bytes, ex_name, sv_bytes, sv_name = generate_report_bytes(subset_df, label)
                        st.session_state.report_files[label] = {
                            "excel_bytes": ex_bytes, "excel_name": ex_name,
                            "sav_bytes": sv_bytes, "sav_name": sv_name
                        }
                    st.session_state.reports_ready = True

            st.success("🎉 Processing complete! Download your report files below:")

            for label, files in st.session_state.report_files.items():
                st.markdown(f"### 📁 {label} Reports")
                col_a, col_b = st.columns(2)
                with col_a:
                    st.download_button(
                        label=f"📥 Download {label} Excel Dashboard",
                        data=files["excel_bytes"],
                        file_name=files["excel_name"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"excel_{label}"
                    )
                with col_b:
                    st.download_button(
                        label=f"📥 Download {label} .sav File",
                        data=files["sav_bytes"],
                        file_name=files["sav_name"],
                        mime="application/octet-stream",
                        key=f"sav_{label}"
                    )

# ==========================================================================
# ==========================================================================
# TAB 4: Q11 RATINGS & REASONS EXTRACTION
# ==========================================================================
# ==========================================================================
with tab4:
    st.markdown("### 📋 Q11 Ratings & Reasons Extraction Control Room")
    st.markdown("Filter records by date window, upload SPSS datasets, and extract Q11 ratings, coded reasons, and open-ended text into structured Excel reports.")
    st.markdown("---")

    st.subheader("⚙ Global Execution Parameters")
    date_mode_q11 = st.radio(
        "Select Date Filtering Mode for Runs:",
        ["Dynamic Past 7 Days (Auto Friday)", "Custom Date Range"],
        horizontal=True,
        key="q11_date_mode",
    )

    today_q11 = datetime.now()
    if date_mode_q11 == "Dynamic Past 7 Days (Auto Friday)":
        current_weekday = today_q11.weekday()
        days_to_subtract = 7 if current_weekday == 4 else (current_weekday - 4) % 7
        if days_to_subtract == 0:
            days_to_subtract = 7
        last_friday_q11 = today_q11 - timedelta(days=days_to_subtract)
        last_friday_q11 = last_friday_q11.replace(hour=0, minute=0, second=0, microsecond=0)
        st.info(f"📅 Target active execution window: **{last_friday_q11.strftime('%Y-%m-%d')}** to **{today_q11.strftime('%Y-%m-%d')}**")
    else:
        col_qd1, col_qd2 = st.columns(2)
        with col_qd1:
            start_date_input_q11 = st.date_input("Start Date", value=today_q11 - timedelta(days=7), key="q11_start")
        with col_qd2:
            end_date_input_q11 = st.date_input("End Date", value=today_q11, key="q11_end")
        last_friday_q11 = datetime.combine(start_date_input_q11, datetime.min.time())
        today_q11 = datetime.combine(end_date_input_q11, datetime.max.time())

    st.markdown("---")
    st.subheader("📁 Upload SPSS Datasets (.sav)")
    
    col_q1, col_q2 = st.columns(2)
    file_q11_r10 = col_q1.file_uploader("Upload R10Mil SPSS File (.sav)", type=["sav"], key="q11_r10_file")
    file_q11_grow = col_q2.file_uploader("Upload Growth SPSS File (.sav)", type=["sav"], key="q11_grow_file")

    if "q11_ready" not in st.session_state: st.session_state.q11_ready = False
    if "q11_bytes" not in st.session_state: st.session_state.q11_bytes = None

    def process_spss_for_q11(uploaded_file, last_friday, today_dt):
        if uploaded_file is None:
            return pd.DataFrame()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
            tmp.write(uploaded_file.getvalue())
            tp = tmp.name
        try:
            df = pd.read_spss(tp, convert_categoricals=False)
            if df.empty or 'STIME' not in df.columns:
                return pd.DataFrame()

            df['STIME_CLEAN'] = pd.to_datetime(df['STIME'].astype(str).str.slice(0, 8), format='%Y%m%d', errors='coerce')
            date_mask = (df['STIME_CLEAN'] >= pd.Timestamp(last_friday)) & (df['STIME_CLEAN'] <= pd.Timestamp(today_dt))
            df = df[date_mask].copy()

            if df.empty:
                return pd.DataFrame()

            df['STIME_STR'] = df['STIME'].astype(str)
            nyear = df['STIME_STR'].str.slice(0, 4)
            nmonth = df['STIME_STR'].str.slice(4, 6)
            nday = df['STIME_STR'].str.slice(6, 8)
            recorded_date = nyear + '-' + nmonth + '-' + nday

            mask = df['INTNR'] > 0 if 'INTNR' in df.columns else np.zeros(len(df), dtype=bool)
            df_out = pd.DataFrame(index=df.index)

            df_out['Interview number'] = df['INTNR'] if 'INTNR' in df.columns else None
            df_out['UCN Number'] = np.where(mask, df['V80116'], None) if 'V80116' in df.columns else None
            df_out['Date'] = np.where(mask, recorded_date, None)

            rating_cols = {
                'Q11.1 RATING - FNB Business Lending products (overdraft, loans, etc.)': 'Q11_1_1',
                'Q11.2 RATING - FNB Business Transactional products (cheque, debit card, credit card etc.)': 'Q11_1_2',
                'Q11.3 RATING - FNB Business Insurance products (business credit protection plan, law-on-call business plan, etc.)': 'Q11_1_3',
                'Q11.4 RATING - FNB Business FNB Business Investment products': 'Q11_1_4',
                'Q11.5 RATING - FNB Business FNB Business FOREX products': 'Q11_1_5'
            }
            for target, src in rating_cols.items():
                if src in df.columns:
                    df_out[target] = np.where(mask, pd.to_numeric(df[src], errors='coerce'), np.nan)
                else:
                    df_out[target] = np.nan

            custom_label_mappings = {
                'Q11 REASONS - Lending_1': {'src': 'Q11A_1_1', 'label': 'Overdraft'},
                'Q11 REASONS - Lending_2': {'src': 'Q11A_1_2', 'label': 'Loans'},
                'Q11 REASONS - Lending_3': {'src': 'Q11A_1_3', 'label': 'Other'},
                'Q11 REASONS - Lending_4': {'src': None, 'label': 'Other'},  
                'Q11 REASONS - Lending_5': {'src': None, 'label': 'Other'},
                
                'Q11 REASONS - Transactional products_1': {'src': 'Q11A_2_1', 'label': 'Cheque card'},
                'Q11 REASONS - Transactional products_2': {'src': 'Q11A_2_2', 'label': 'Debit card'},
                'Q11 REASONS - Transactional products_3': {'src': 'Q11A_2_3', 'label': 'Credit Card'},
                'Q11 REASONS - Transactional products_4': {'src': 'Q11A_2_4', 'label': 'Other'},
                'Q11 REASONS - Transactional products_5': {'src': None, 'label': 'Other'},
                
                'Q11 REASONS - Insurance_1': {'src': 'Q11A_3_1', 'label': 'Business credit protection plan'},
                'Q11 REASONS - Insurance_2': {'src': 'Q11A_3_2', 'label': 'Law-on-call business plan'},
                'Q11 REASONS - Insurance_3': {'src': 'Q11A_3_3', 'label': 'Other'},
                'Q11 REASONS - Insurance_4': {'src': None, 'label': 'Other'},
                'Q11 REASONS - Insurance_5': {'src': None, 'label': 'Other'},
                
                'Q11 REASONS - Investment products_1': {'src': 'Q11A_4_1', 'label': 'Savings'},
                'Q11 REASONS - Investment products_2': {'src': 'Q11A_4_2', 'label': 'Notice deposits'},
                'Q11 REASONS - Investment products_3': {'src': 'Q11A_4_3', 'label': 'Other'},
                'Q11 REASONS - Investment products_4': {'src': None, 'label': 'Other'},
                'Q11 REASONS - Investment products_5': {'src': None, 'label': 'Other'},
            }

            forex_labels = {
                1: 'Foreign Exchange', 2: 'Imports and Exports', 3: 'Structured Trade + Commodity Finance',
                4: 'PayPal', 5: 'Trade (Trade Platform and Transacting)', 6: 'MoneyGram (TM)',
                7: 'Global Payments (business global account)', 8: 'Travel card',
                9: 'Trans-country Interbank Clearing', 10: 'Other'
            }
            for i, label_text in forex_labels.items():
                custom_label_mappings[f'Q11 REASONS - FOREX products_{i}'] = {'src': f'Q11A_5_{i}', 'label': label_text}

            for target_col, config in custom_label_mappings.items():
                src_col = config['src']
                if src_col is None or src_col not in df.columns:
                    df_out[target_col] = None
                    continue
                numeric_src = pd.to_numeric(df[src_col], errors='coerce')
                df_out[target_col] = np.where((mask) & (numeric_src == 1), config['label'], None)

            open_ends = {
                'Q11 REASONS - Lending OTHER': 'TQ11A_1C3',
                'Q11 REASONS OTHER - Transactional products': 'TQ11A_2C4',
                'Q11 REASONS - Insurance OTHER': 'TQ11A_3C3',
                'Q11 REASONS - Investment products OTHER': 'TQ11A_4C3',
                'Q11 REASONS - FOREX products OTHER': 'TQ11A_5C10'
            }
            for target, src in open_ends.items():
                if src in df.columns:
                    string_series = df[src].astype(str).replace(['nan', 'NaN', 'None'], None)
                    df_out[target] = np.where(mask, string_series, None)
                else:
                    df_out[target] = None

            final_column_order = [
                'Interview number', 'UCN Number', 'Date',
                'Q11.1 RATING - FNB Business Lending products (overdraft, loans, etc.)', 
                'Q11 REASONS - Lending_1', 'Q11 REASONS - Lending_2', 'Q11 REASONS - Lending_3', 'Q11 REASONS - Lending_4', 'Q11 REASONS - Lending_5', 
                'Q11 REASONS - Lending OTHER',
                
                'Q11.2 RATING - FNB Business Transactional products (cheque, debit card, credit card etc.)', 
                'Q11 REASONS - Transactional products_1', 'Q11 REASONS - Transactional products_2', 'Q11 REASONS - Transactional products_3', 'Q11 REASONS - Transactional products_4', 'Q11 REASONS - Transactional products_5', 
                'Q11 REASONS OTHER - Transactional products',
                
                'Q11.3 RATING - FNB Business Insurance products (business credit protection plan, law-on-call business plan, etc.)', 
                'Q11 REASONS - Insurance_1', 'Q11 REASONS - Insurance_2', 'Q11 REASONS - Insurance_3', 'Q11 REASONS - Insurance_4', 'Q11 REASONS - Insurance_5', 
                'Q11 REASONS - Insurance OTHER',
                
                'Q11.4 RATING - FNB Business FNB Business Investment products', 
                'Q11 REASONS - Investment products_1', 'Q11 REASONS - Investment products_2', 'Q11 REASONS - Investment products_3', 'Q11 REASONS - Investment products_4', 'Q11 REASONS - Investment products_5', 
                'Q11 REASONS - Investment products OTHER',
                
                'Q11.5 RATING - FNB Business FNB Business FOREX products', 
                'Q11 REASONS - FOREX products_1', 'Q11 REASONS - FOREX products_2', 'Q11 REASONS - FOREX products_3', 'Q11 REASONS - FOREX products_4', 'Q11 REASONS - FOREX products_5', 'Q11 REASONS - FOREX products_6', 'Q11 REASONS - FOREX products_7', 'Q11 REASONS - FOREX products_8', 'Q11 REASONS - FOREX products_9', 'Q11 REASONS - FOREX products_10', 
                'Q11 REASONS - FOREX products OTHER'
            ]
            
            df_final = df_out[final_column_order].copy()
            clean_headers = [header.split('_')[0] for header in df_final.columns]
            df_final.columns = clean_headers
            return df_final
        except Exception as e:
            st.error(f"Error processing SPSS file: {e}")
            return pd.DataFrame()
        finally:
            if os.path.exists(tp):
                os.remove(tp)

    if st.button("▶ Run Q11 Extraction", type="primary", key="run_q11_btn"):
        if not file_q11_r10 and not file_q11_grow:
            st.error("Please upload at least one SPSS dataset (.sav) before running the extraction.")
        else:
            with st.spinner("Processing Q11 extractions with date window filter..."):
                clean_df1 = process_spss_for_q11(file_q11_r10, last_friday_q11, today_q11)
                clean_df2 = process_spss_for_q11(file_q11_grow, last_friday_q11, today_q11)

                out_buf = io.BytesIO()
                with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
                    if not clean_df1.empty:
                        clean_df1.to_excel(writer, sheet_name="Enterprise-R10Mil", index=False)
                    if not clean_df2.empty:
                        clean_df2.to_excel(writer, sheet_name="Business-Growth", index=False)
                    if clean_df1.empty and clean_df2.empty:
                        pd.DataFrame({"Notice": ["No records found for the selected date window."]}).to_excel(writer, sheet_name="No Data", index=False)

                out_buf.seek(0)
                st.session_state.q11_bytes = out_buf.getvalue()
                st.session_state.q11_ready = True
                st.success("✅ Q11 extraction completed successfully using your exact standalone script logic!")

    if st.session_state.q11_ready and st.session_state.q11_bytes:
        st.markdown("---")
        run_date_file = datetime.now().strftime("%Y-%m-%d")
        st.download_button(
            label="📥 Download Q11 Extraction Report (.xlsx)",
            data=st.session_state.q11_bytes,
            file_name=f"Star W22 Q11 extraction {run_date_file}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_q11_xlsx"
        )

# ==========================================================================
# ==========================================================================
# TAB 5: NPS YEARLY DASHBOARD GENERATOR (CROSS-TABULATION GENERATOR)
# ==========================================================================
# ==========================================================================
with tab5:
    st.markdown("### 📊 Project Star: Yearly Cross-Tabulation Generator")
    st.markdown("Upload your latest yearly SPSS file (`.sav`) below to generate and download the cross-tabulation report.")

    uploaded_yearly_file = st.file_uploader("Upload Yearly SPSS File (.sav)", type=["sav"], key="tab5_yearly_sav")

    if uploaded_yearly_file is not None:
        if st.button("Generate Yearly Dashboard Report", type="primary", key="tab5_gen_yearly_btn"):
            with st.spinner("Processing SPSS file and generating cross-tabulation report... Please wait."):
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp_file:
                    tmp_file.write(uploaded_yearly_file.getvalue())
                    tmp_file_path = tmp_file.name

                try:
                    df, meta = pyreadstat.read_sav(tmp_file_path)
                    df = df.loc[:, ~df.columns.duplicated()].copy()

                    rm_driver_cols = [f"Q4_{i:02d}" for i in range(1, 24)]
                    channel_q2_cols = ['Q2_1', 'Q2_2', 'Q2_3', 'Q2_4', 'Q2_5', 'Q2_6', 'Q2_7']
                    pref_channel_col = 'Q2_2_1'
                    personal_banker_col = 'Q4A2'
                    pb_sat_col = 'Q4B2'
                    branch_service_cols = ['Q5_1', 'Q5_2', 'Q5_3', 'Q5_4', 'Q5_21', 'Q5_22', 'Q5_23', 'Q5_24']
                    branch_cols = branch_service_cols + ['Q10_2']
                    cc_cols = ['Q7_1', 'Q7_2', 'Q7_3', 'Q7_4', 'Q7_5', 'Q7_21', 'Q7_22', 'Q10_3']

                    online_cols = ['Q8_1', 'Q8_2', 'Q8_3', 'Q8_4', 'Q8_21', 'Q8_22', 'Q8_23', 'Q8_24', 'Q8_25', 'Q8_26', 'Q8_27']
                    app_cols = ['Q9_1', 'Q9_2', 'Q9_3', 'Q9_4', 'Q9_21', 'Q9_22', 'Q9_23', 'Q9_24', 'Q9_25', 'Q9_26']
                    chan_sat_cols = ['Q10_1', 'Q10_2', 'Q10_3', 'Q10_4', 'Q10_5']
                    product_sat_cols = ['Q11_1_1', 'Q11_1_2', 'Q11_1_3', 'Q11_1_4', 'Q11_1_5']

                    q11a_groups = {
                        'Q11A_1': [('Q11a.1. Lending products - Overdraft', 'Q11A_1_1'), ('Q11a.1. Lending products - Loans', 'Q11A_1_2'), ('Q11a.1. Lending products - Other', 'Q11A_1_3')],
                        'Q11A_2': [('Q11a.2. Transactional products - Business account', 'Q11A_2_1'), ('Q11a.2. Transactional products - Debit card', 'Q11A_2_2'), ('Q11a.2. Transactional products - Credit Card', 'Q11A_2_3'), ('Q11a.2. Transactional products - Other', 'Q11A_2_4')],
                        'Q11A_3': [('Q11a.3. Insurance products - Business credit protection plan', 'Q11A_3_1'), ('Q11a.3. Insurance products - Law-on-call business plan', 'Q11A_3_2'), ('Q11a.3. Insurance products - Other', 'Q11A_3_3')],
                        'Q11A_4': [('Q11a.4. Investment products - Savings', 'Q11A_4_1'), ('Q11A_4. Investment products - Notice deposits', 'Q11A_4_2'), ('Q11A_4. Investment products - Other', 'Q11A_4_3')]
                    }

                    all_q11a_cols = [col for grp in q11a_groups.values() for _, col in grp]
                    expectations_cols = ['Q12_1', 'Q12_2', 'Q12_3']
                    q16_col = 'Q16'

                    consideration_items = [
                        ("Absa", "Q16_1_1_flag"), ("Capitec", "Q16_1_2_flag"), ("Investec", "Q16_1_8_flag"),
                        ("Mercantile", "Q16_1_9_flag"), ("Nedbank", "Q16_1_3_flag"), ("Sasfin", "Q16_1_10_flag"),
                        ("Standard Bank", "Q16_1_4_flag"), ("Some other business banking offering", "Q16_1_5_flag"),
                        ("I would not consider moving from FNB at all", "Q16_1_6_flag")
                    ]
                    raw_consideration_cols = ['Q16_1_1', 'Q16_1_2', 'Q16_1_3', 'Q16_1_4', 'Q16_1_5', 'Q16_1_6', 'Q16_1_8', 'Q16_1_9', 'Q16_1_10']
                    q5a_cols = ['Q5A_1', 'Q5A_3', 'Q5A_4', 'Q5A_5', 'Q5A_6', 'Q5A_7', 'Q5A_8', 'Q5A_9', 'Q5A_10', 'Q5A_11', 'Q5A_12', 'Q5A_13', 'Q5A_14', 'Q5A_2']

                    all_rating_cols = list(set(rm_driver_cols + branch_cols + cc_cols + online_cols + app_cols + chan_sat_cols + product_sat_cols + expectations_cols))
                    if pb_sat_col and pb_sat_col not in all_rating_cols:
                        all_rating_cols.append(pb_sat_col)

                    channel_items = [
                        ("Q2. Business Manager", "Q2_1_flag"), ("Q2. Branch", "Q2_2_flag"), ("Q2. Contact Centre", "Q2_3_flag"),
                        ("Q2. Online Banking", "Q2_4_flag"), ("Q2. FNB Business Banking App", "Q2_5_flag"), ("Q2. Account fulfilment", "Q6_code1_flag"),
                        ("Q2. Product contact centre", "Q6_code2_flag"), ("Q2. Business Desk", "Q6_code3_flag"), ("Q2. Secure chat", "Q2_7_flag"), ("Q2. None of the above", "Q2_6_flag"),
                    ]

                    q5a_items = [
                        ("Q5a. Prefer face-to-face interaction", "Q5A_1_flag"), ("Q5a. Required to submit documents", "Q5A_3_flag"),
                        ("Q5a. Card collection", "Q5A_4_flag"), ("Q5a. Card queries", "Q5A_5_flag"),
                        ("Q5a. Issue could not be resolved digitally/limited options on digital channels", "Q5A_6_flag"),
                        ("Q5a. Difficulty using digital channels", "Q5A_7_flag"), ("Q5a. Not enough information on digital channels", "Q5A_8_flag"),
                        ("Q5a. Directed to branch", "Q5A_9_flag"), ("Q5a. Needed help/assistance", "Q5A_10_flag"),
                        ("Q5a. Cash deposit/withdrawal", "Q5A_11_flag"), ("Q5a. To open account", "Q5A_12_flag"),
                        ("Q5a. To get bank statement", "Q5A_13_flag"), ("Q5a. Update personal information/details", "Q5A_14_flag"),
                        ("Q5a. Others; specify", "Q5A_2_flag"),
                    ]

                    q6_items = [
                        ("FICA, outstanding documents relating to your account", "Q6_code_1"),
                        ("A specific product, e.g. such as Instant Solutions", "Q6_code_2"),
                        ("General enquiries (e.g. account, card & cheque-related)", "Q6_code_3"),
                        ("Don't know / not sure", "Q6_code_4")
                    ]

                    pref_channel_items = [
                        ("Branch", 1), ("Contact Centre", 2), ("Online Banking", 3), ("FNB Business Banking App", 4),
                        ("Business Manager at the branch", 5), ("Business/RM Manager", 6), ("Secure chat Help note", 7), ("None of the above", 8)
                    ]

                    personal_banker_items = [
                        ("Yes", 1), ("No", 2), ("Do not have a personal FNB Account", 3), ("Refused to answer", 4)
                    ]

                    rm_labels = {}
                    for col in rm_driver_cols:
                        if col in meta.column_names_to_labels and meta.column_names_to_labels[col]:
                            rm_labels[col] = re.sub(r'^(Q4[\._\s\d]*)+', '', meta.column_names_to_labels[col], flags=re.IGNORECASE).strip()
                        else:
                            rm_labels[col] = f"Statement {col}"

                    pb_sat_label = "Overall satisfaction with your Personal Banker"
                    if pb_sat_col and pb_sat_col in meta.column_names_to_labels and meta.column_names_to_labels[pb_sat_col]:
                        pb_sat_label = re.sub(r'^(Q4B2[\._\s\d]*)+', '', meta.column_names_to_labels[pb_sat_col], flags=re.IGNORECASE).strip()

                    explicit_branch_labels = {
                        'Q5_1': 'Offering you personalized service', 'Q5_2': 'The manner in which you are welcomed and directed',
                        'Q5_3': 'Staff understanding your business banking needs', 'Q5_4': 'Waiting time for service',
                        'Q5_21': 'Operating hours of the branch', 'Q5_22': 'Staff communicating with you in a clear and easily understandable way',
                        'Q5_23': 'Staff being willing to help', 'Q5_24': 'Consistently delivering on promises made to you', 'Q10_2': 'OVERALL Branch experience'
                    }
                    branch_labels = {col: explicit_branch_labels.get(col, col) for col in branch_cols}

                    explicit_cc_labels = {
                        'Q7_1': 'Knowledge and competency', 'Q7_2': 'Taking ownership of your query', 'Q7_3': 'Offering you personalized service',
                        'Q7_4': 'Processing requests accurately', 'Q7_5': 'Providing the correct advice relating to your query',
                        'Q7_21': 'The agent providing the correct advice relating to your enquiry or transaction',
                        'Q7_22': 'The agent delivering on promises made', 'Q10_3': 'OVERALL Contact Centre experience'
                    }
                    cc_labels = {col: explicit_cc_labels.get(col, col) for col in cc_cols}

                    online_labels = {col: meta.column_names_to_labels[col] if (col in meta.column_names_to_labels and meta.column_names_to_labels[col]) else col for col in online_cols}
                    app_labels = {col: meta.column_names_to_labels[col] if (col in meta.column_names_to_labels and meta.column_names_to_labels[col]) else col for col in app_cols}

                    explicit_chan_sat_items = [
                        ("Q10.1. OVERALL - Business Manager experience", "Q10_1"), ("Q10.2. OVERALL - Branch experience", "Q10_2"),
                        ("Q10.3. OVERALL - Contact Centre experience", "Q10_3"), ("Q10.4. OVERALL - Online Banking experience", "Q10_4"),
                        ("Q10.5. OVERALL - FNB Business Banking App experience?", "Q10_5"),
                    ]

                    explicit_product_sat_items = [
                        ("Q11.1. FNB Business Lending products", "Q11_1_1"), ("Q11.2. FNB Business Transactional products", "Q11_1_2"),
                        ("Q11.3. FNB Business Insurance products", "Q11_1_3"), ("Q11.4. FNB Business Investment products", "Q11_1_4"),
                        ("Q11.5. FNB Business Forex Products", "Q11_1_5")
                    ]

                    expectations_items = [
                        ("Q12.1. Your overall level of satisfaction with the products you received from FNB Business?", "Q12_1"),
                        ("Q12.2. Your overall level of satisfaction with FNB Business?", "Q12_2"),
                        ("Q12.3. Your overall level of satisfaction with your BM over the last 3 months?", "Q12_3")
                    ]

                    cols = ['wave', 'REGION', 'SUBREG', 'SEGMENT', 'Type', 'Q14_1', 'Q14_2', 'Q6', q16_col] + all_rating_cols + channel_q2_cols + q5a_cols + [pref_channel_col, personal_banker_col, pb_sat_col] + raw_consideration_cols + all_q11a_cols
                    cols_present = [c for c in cols if c in df.columns]

                    df_sub = df[cols_present].copy()
                    df_sub = df_sub.loc[:, ~df_sub.columns.duplicated()].copy()

                    df_sub['Q6_raw'] = df_sub['Q6'].copy() if 'Q6' in df_sub.columns else None

                    for col in ['wave', 'REGION', 'SUBREG', 'SEGMENT', 'Type']:
                        if col in meta.variable_value_labels and col in df_sub.columns:
                            df_sub[col] = df_sub[col].map(meta.variable_value_labels[col]).fillna(df_sub[col])

                    df_sub['REGION'] = df_sub['REGION'].fillna("Unspecified")
                    df_sub['SUBREG'] = df_sub['SUBREG'].fillna("Unspecified")
                    df_sub['SEGMENT'] = df_sub['SEGMENT'].fillna("Unspecified")

                    segment_relabel_map = {'MEDIUM TOUCH': 'MEDIUM TOUCH (R10-R60M)', 'HIGH TOUCH': 'HIGH TOUCH (R60-R150M)', 'PREMIUM': 'PREMIUM (R150M+)'}
                    df_sub['SEGMENT'] = df_sub['SEGMENT'].apply(lambda x: segment_relabel_map.get(str(x).strip(), str(x).strip()))

                    def assign_type(segment_val):
                        seg = str(segment_val).strip().upper()
                        if seg in ['R0M-R1M', 'R1M-R5M', 'R5M-R10M', 'MEDIUM TOUCH (R10-R60M)', 'MEDIUM TOUCH']:
                            return 'Growth'
                        elif seg in ['HIGH TOUCH (R60-R150M)', 'HIGH TOUCH', 'PREMIUM (R150M+)', 'PREMIUM']:
                            return 'R10m+'
                        return 'Growth'

                    df_sub['Type'] = df_sub['SEGMENT'].apply(assign_type)

                    def get_wave_number(val):
                        match = re.search(r'\d+', str(val))
                        return int(match.group()) if match else 999

                    df_sub['wave_num'] = df_sub['wave'].apply(get_wave_number)
                    sorted_wave_nums = sorted(df_sub['wave_num'].unique())
                    regions = sorted([str(x) for x in df_sub['REGION'].unique() if x != "Unspecified"])

                    def clean_rating_score(val):
                        try:
                            fval = float(val)
                            if fval == 11 or fval == 11.0: return None
                            return fval if 1 <= fval <= 10 else None
                        except (ValueError, TypeError):
                            return None

                    new_cols = {}
                    for col in all_rating_cols:
                        if col in df_sub.columns:
                            new_cols[f"{col}_clean"] = df_sub[col].apply(clean_rating_score)

                    new_cols['Q14_1_clean'] = df_sub['Q14_1'].apply(lambda x: x if pd.notnull(x) and x in range(0, 11) else None)
                    new_cols['Q14_2_clean'] = df_sub['Q14_2'].apply(lambda x: x if pd.notnull(x) and x in range(0, 11) else None)

                    def is_q6_match(val, code_num, desc_text):
                        if pd.isnull(val): return False
                        sval = str(val).strip().lower()
                        return sval == str(code_num) or sval == f"{code_num}.0" or desc_text.lower() in sval

                    if 'Q6_raw' in df_sub.columns:
                        q10_3_clean = new_cols.get('Q10_3_clean', pd.Series(index=df_sub.index))
                        new_cols['Q6_code1_clean'] = [q10_3_clean[i] if is_q6_match(v, 1, "Business Account Fulfilment") else None for i, v in enumerate(df_sub['Q6_raw'])]
                        new_cols['Q6_code2_clean'] = [q10_3_clean[i] if is_q6_match(v, 2, "Product Contact Centre") else None for i, v in enumerate(df_sub['Q6_raw'])]
                        new_cols['Q6_code3_clean'] = [q10_3_clean[i] if is_q6_match(v, 3, "Business Desk") else None for i, v in enumerate(df_sub['Q6_raw'])]
                        new_cols['Q6_code1_flag'] = [1 if is_q6_match(v, 1, "Business Account Fulfilment") else 0 for v in df_sub['Q6_raw']]
                        new_cols['Q6_code2_flag'] = [1 if is_q6_match(v, 2, "Product Contact Centre") else 0 for v in df_sub['Q6_raw']]
                        new_cols['Q6_code3_flag'] = [1 if is_q6_match(v, 3, "Business Desk") else 0 for v in df_sub['Q6_raw']]
                        for code_val in [1, 2, 3, 4]:
                            new_cols[f"Q6_code_{code_val}"] = df_sub['Q6_raw'].apply(lambda x: 1 if pd.notnull(x) and float(x) == code_val else 0)

                    for q11a_key, sub_items in q11a_groups.items():
                        sub_cols = [c for _, c in sub_items if c in df_sub.columns]
                        if sub_cols: df_sub[f"{q11a_key}_Base"] = df_sub[sub_cols].notnull().any(axis=1).astype(int)
                        for _, col in sub_items:
                            if col in df_sub.columns: new_cols[f"{col}_flag"] = df_sub[col].apply(lambda x: 1 if pd.notnull(x) and float(x) == 1 else 0)

                    if q16_col in df_sub.columns:
                        df_sub['Q16_mapped'] = df_sub[q16_col].apply(lambda x: 'Yes' if str(x).strip().lower() in ['1', '1.0', 'yes'] else ('No' if str(x).strip().lower() in ['2', '2.0', 'no'] else None))
                        new_cols['Q16_Yes'] = df_sub['Q16_mapped'].apply(lambda x: 1 if x == 'Yes' else 0)
                        new_cols['Q16_No'] = df_sub['Q16_mapped'].apply(lambda x: 1 if x == 'No' else 0)
                        new_cols['Q16_Base'] = df_sub['Q16_mapped'].apply(lambda x: 1 if pd.notnull(x) else 0)
                    else:
                        new_cols['Q16_Yes'], new_cols['Q16_No'], new_cols['Q16_Base'] = 0, 0, 0

                    for _, flag_col in consideration_items:
                        raw_col = flag_col.replace('_flag', '')
                        new_cols[flag_col] = df_sub[raw_col].apply(lambda x: 1 if pd.notnull(x) and float(x) == 1 else 0) if raw_col in df_sub.columns else 0

                    bm_clean, fnb_clean = new_cols['Q14_2_clean'], new_cols['Q14_1_clean']
                    new_cols['BM_Det'] = bm_clean.apply(lambda x: 1 if pd.notnull(x) and x <= 6 else 0)
                    new_cols['BM_Pas'] = bm_clean.apply(lambda x: 1 if pd.notnull(x) and 7 <= x <= 8 else 0)
                    new_cols['BM_Pro'] = bm_clean.apply(lambda x: 1 if pd.notnull(x) and x >= 9 else 0)
                    new_cols['BM_Base'] = bm_clean.apply(lambda x: 1 if pd.notnull(x) else 0)

                    new_cols['FNB_Det'] = fnb_clean.apply(lambda x: 1 if pd.notnull(x) and x <= 6 else 0)
                    new_cols['FNB_Pas'] = fnb_clean.apply(lambda x: 1 if pd.notnull(x) and 7 <= x <= 8 else 0)
                    new_cols['FNB_Pro'] = fnb_clean.apply(lambda x: 1 if pd.notnull(x) and x >= 9 else 0)
                    new_cols['FNB_Base'] = fnb_clean.apply(lambda x: 1 if pd.notnull(x) else 0)

                    for qcol in channel_q2_cols:
                        if qcol in df_sub.columns: new_cols[f"{qcol}_flag"] = df_sub[qcol].apply(lambda x: 1 if pd.notnull(x) and float(x) == 1 else 0)
                    for qcol in q5a_cols:
                        if qcol in df_sub.columns: new_cols[f"{qcol}_flag"] = df_sub[qcol].apply(lambda x: 1 if pd.notnull(x) and float(x) == 1 else 0)
                    if pref_channel_col in df_sub.columns:
                        for lbl, code in pref_channel_items: new_cols[f"pref_chan_{code}"] = df_sub[pref_channel_col].apply(lambda x: 1 if pd.notnull(x) and float(x) == code else 0)
                    if personal_banker_col in df_sub.columns:
                        for lbl, code in personal_banker_items: new_cols[f"pb_code_{code}"] = df_sub[personal_banker_col].apply(lambda x: 1 if pd.notnull(x) and float(x) == code else 0)

                    df_sub = pd.concat([df_sub, pd.DataFrame(new_cols, index=df_sub.index)], axis=1)
                    df_sub = df_sub.loc[:, ~df_sub.columns.duplicated()].copy()

                    q16a_flag_cols = [flag_col for _, flag_col in consideration_items if flag_col in df_sub.columns]
                    df_sub['Q16A_Base'] = (df_sub[q16a_flag_cols].sum(axis=1) > 0).astype(int) if q16a_flag_cols else 0
                    df_sub['Channel_Base'] = df_sub[[c for c in channel_q2_cols if c in df_sub.columns]].notnull().any(axis=1).astype(int)
                    df_sub['Q5A_Base'] = (df_sub[[f"{q}_flag" for q in q5a_cols if f"{q}_flag" in df_sub.columns]].sum(axis=1) > 0).astype(int)
                    df_sub['Q6_Base'] = df_sub['Q6_raw'].apply(lambda x: 1 if pd.notnull(x) else 0)
                    df_sub['Pref_Channel_Base'] = df_sub[pref_channel_col].apply(lambda x: 1 if pd.notnull(x) else 0) if pref_channel_col in df_sub.columns else 0
                    df_sub['PB_Base'] = df_sub[personal_banker_col].apply(lambda x: 1 if pd.notnull(x) else 0) if personal_banker_col in df_sub.columns else 0

                    df_xtab = df_sub[df_sub['Type'].isin(['Growth', 'R10m+'])].copy()

                    xtab_base_dict = {
                        'Total_n': ('wave_num', 'count'),
                        'FNB_Base': ('FNB_Base', 'sum'), 'FNB_Det': ('FNB_Det', 'sum'), 'FNB_Pro': ('FNB_Pro', 'sum'),
                        'BM_Base': ('BM_Base', 'sum'), 'BM_Det': ('BM_Det', 'sum'), 'BM_Pro': ('BM_Pro', 'sum'),
                        'Channel_Base': ('Channel_Base', 'sum'), 'Q5A_Base': ('Q5A_Base', 'sum'),
                        'Q6_Base': ('Q6_Base', 'sum'), 'Pref_Channel_Base': ('Pref_Channel_Base', 'sum'), 'PB_Base': ('PB_Base', 'sum'),
                    }

                    for col_name in ['Q16_Base', 'Q16_Yes', 'Q16_No', 'Q16A_Base']:
                        if col_name in df_xtab.columns: xtab_base_dict[col_name] = (col_name, 'sum')
                    for _, flag_col in consideration_items:
                        if flag_col in df_xtab.columns: xtab_base_dict[flag_col] = (flag_col, 'sum')
                    for q11a_key in q11a_groups.keys():
                        if f"{q11a_key}_Base" in df_xtab.columns: xtab_base_dict[f"{q11a_key}_Base"] = (f"{q11a_key}_Base", 'sum')
                    for col in all_rating_cols:
                        clean_name = f"{col}_clean"
                        if clean_name not in df_xtab.columns: df_xtab[clean_name] = None
                        xtab_base_dict[f"{col}_sum"] = (clean_name, 'sum')
                        xtab_base_dict[f"{col}_count"] = (clean_name, 'count')
                    for code_num in [1, 2, 3]:
                        clean_col_name = f"Q6_code{code_num}_clean"
                        if clean_col_name not in df_xtab.columns: df_xtab[clean_col_name] = None
                        xtab_base_dict[f"Q6_code{code_num}_sum"] = (clean_col_name, 'sum')
                        xtab_base_dict[f"Q6_code{code_num}_count"] = (clean_col_name, 'count')
                    for _, flag_col in channel_items:
                        if flag_col in df_xtab.columns: xtab_base_dict[flag_col] = (flag_col, 'sum')
                    for _, flag_col in q5a_items:
                        if flag_col in df_xtab.columns: xtab_base_dict[flag_col] = (flag_col, 'sum')
                    for q11a_key, sub_items in q11a_groups.items():
                        for _, col in sub_items:
                            flag_name = f"{col}_flag"
                            if flag_name in df_xtab.columns: xtab_base_dict[flag_name] = (flag_name, 'sum')
                    for code_val in [1, 2, 3, 4]:
                        col_name = f"Q6_code_{code_val}"
                        if col_name in df_xtab.columns: xtab_base_dict[col_name] = (col_name, 'sum')
                    for _, code in pref_channel_items:
                        flag_name = f"pref_chan_{code}"
                        if flag_name in df_xtab.columns: xtab_base_dict[flag_name] = (flag_name, 'sum')
                    for _, code in personal_banker_items:
                        flag_name = f"pb_code_{code}"
                        if flag_name in df_xtab.columns: xtab_base_dict[flag_name] = (flag_name, 'sum')

                    xtab_agg_dict = {k: v for k, v in xtab_base_dict.items() if v[0] in df_xtab.columns}
                    summary_xtab_wave_type = df_xtab.groupby(['wave_num', 'Type'], as_index=False).agg(**xtab_agg_dict)
                    summary_xtab_wave_region = df_xtab.groupby(['wave_num', 'REGION', 'Type'], as_index=False).agg(**xtab_agg_dict)

                    wb_xtab = openpyxl.Workbook()
                    ws_xtab = wb_xtab.active
                    ws_xtab.title = "Cross Tabulation"
                    ws_xtab.views.sheetView[0].showGridLines = True

                    ws_xtab_wt_data = wb_xtab.create_sheet(title="_CrossTab_WaveType_Data")
                    ws_xtab_wt_data.sheet_state = 'hidden'
                    ws_xtab_wt_data.append(list(summary_xtab_wave_type.columns))
                    for row in summary_xtab_wave_type.itertuples(index=False): ws_xtab_wt_data.append(list(row))

                    ws_xtab_wr_data = wb_xtab.create_sheet(title="_CrossTab_WaveReg_Data")
                    ws_xtab_wr_data.sheet_state = 'hidden'
                    ws_xtab_wr_data.append(list(summary_xtab_wave_region.columns))
                    for row in summary_xtab_wave_region.itertuples(index=False): ws_xtab_wr_data.append(list(row))

                    FNB_TEAL, FNB_AMBER, LIGHT_TEAL, LIGHT_AMBER, WHITE = "009B9E", "EA8B24", "E5F5F5", "FDF3E7", "FFFFFF"
                    font_title_main = Font(name="Calibri", size=14, bold=True, color=WHITE)
                    font_title_sub = Font(name="Calibri", size=10, italic=True, color=WHITE)
                    font_header = Font(name="Calibri", size=9, bold=True, color=WHITE)
                    font_bold = Font(name="Calibri", size=10, bold=True)
                    font_regular = Font(name="Calibri", size=10)
                    font_sub_subheader = Font(name="Calibri", size=10, italic=True, bold=True, color="105B5C")
                    font_filter_lbl = Font(name="Calibri", size=10, bold=True, color=WHITE)
                    font_filter_val = Font(name="Calibri", size=11, bold=True, color="105B5C")

                    fill_teal_header = PatternFill(start_color=FNB_TEAL, end_color=FNB_TEAL, fill_type="solid")
                    fill_amber_header = PatternFill(start_color=FNB_AMBER, end_color=FNB_AMBER, fill_type="solid")
                    fill_light_teal = PatternFill(start_color=LIGHT_TEAL, end_color=LIGHT_TEAL, fill_type="solid
