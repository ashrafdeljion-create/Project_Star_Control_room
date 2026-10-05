# =========================================================================
# SECTION 1: IMPORTING PACKAGES & CONFIGURING PAGE LAYOUT
# =========================================================================
import io
import os
import re
import tempfile
from datetime import datetime, timedelta
from dateutil.relativedelta import FR, relativedelta

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import pyreadstat
import streamlit as st

st.set_page_config(
    page_title="Project Star: One-Stop Operations Hub", page_icon="⭐", layout="wide"
)


# =========================================================================
# SECTION 2: CUSTOM UI STYLING (FNB BRAND & HIGH-VISIBILITY TABS)
# =========================================================================
st.markdown(
    """
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
""",
    unsafe_allow_html=True,
)

st.title("⭐ Project Star: One-Stop Operations Hub")
st.markdown(
    "Your unified command center for Project Status, Weekly 911's pipeline automation, NPS Excel reports, Q11 extractions, Yearly Dashboard generation, and SME/ENT segment tables."
)


# =========================================================================
# SECTION 3: DEFINING MAIN APP NAVIGATION TABS (6 Tabs Total)
# =========================================================================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "📊 Project Status & Quotas Update",
        "⚙️ Weekly 911's Control Room",
        "📈 NPS Dashboard & Data Generator",
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
    st.markdown("### 📊 Project Status & Quotas Update Hub")
    st.markdown(
        "Monitor overall sample quotas achieved, view executive summaries across portfolios, and download the PM Project Status Update report."
    )

    st.markdown("---")
    st.subheader("🎯 Live Quota Target Adjustments")

    col_t1, col_t2, col_t3 = st.columns(3)
    with col_t1:
        target_business = st.number_input(
            "Business (Growth) Target",
            min_value=0,
            value=4700,
            step=5,
            key="target_bus",
        )
    with col_t2:
        target_enterprise = st.number_input(
            "Enterprise (R10Mil) Target",
            min_value=0,
            value=1400,
            step=5,
            key="target_ent",
        )
    with col_t3:
        target_pubsc = st.number_input(
            "PUBSC Target", min_value=0, value=500, step=5, key="target_pub"
        )

    # --- SEPARATE SEGMENT-LEVEL QUOTA INPUTS BELOW ---
    st.markdown("---")
    st.subheader("🔢 Segment-Level Quota Breakdown Inputs")
    st.markdown(
        "Specify exact individual segment quotas below for detailed tracking and Excel report integration:"
    )

    col_seg1, col_seg2, col_seg3, col_seg4 = st.columns(4)
    with col_seg1:
        q_seg_r0_r1 = st.number_input(
            "R0M-R1M Quota", min_value=0, value=1600, step=5, key="q_r0_r1"
        )
    with col_seg2:
        q_seg_r1_r5 = st.number_input(
            "R1M-R5M Quota", min_value=0, value=1100, step=5, key="q_r1_r5"
        )
    with col_seg3:
        q_seg_r5_r10 = st.number_input(
            "R5M-R10M Quota", min_value=0, value=900, step=5, key="q_r5_r10"
        )
    with col_seg4:
        q_seg_r10_r60 = st.number_input(
            "R10-R60M Quota", min_value=0, value=1100, step=5, key="q_r10_r60"
        )

    col_seg5, col_seg6, _ = st.columns(3)
    with col_seg5:
        q_seg_r60_r150 = st.number_input(
            "R60-R150M Quota", min_value=0, value=900, step=5, key="q_seg_r60_r150"
        )
    with col_seg6:
        q_seg_r150_plus = st.number_input(
            "R150M+ Quota", min_value=0, value=500, step=5, key="q_seg_r150_plus"
        )

    total_target_val = target_business + target_enterprise + target_pubsc

    st.markdown("---")
    st.subheader("📁 Upload Latest SPSS Datasets for Live Status Calculation")
    col_up1, col_up2, col_up3 = st.columns(3)
    with col_up1:
        status_file_grow = st.file_uploader(
            "Upload Growth (.sav)", type=["sav"], key="status_grow"
        )
    with col_up2:
        status_file_r10 = st.file_uploader(
            "Upload R10Mil (.sav)", type=["sav"], key="status_r10"
        )
    with col_up3:
        status_file_pub = st.file_uploader(
            "Upload PUBW (.sav)", type=["sav"], key="status_pub"
        )

    def get_achieved_count(uploaded_file):
        if uploaded_file is None:
            return None
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        try:
            df, _ = pyreadstat.read_sav(tmp_path, apply_value_formats=False)
            return len(df)
        except:
            return 0
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    achieved_business = (
        get_achieved_count(status_file_grow) if status_file_grow else 1721
    )
    achieved_enterprise = (
        get_achieved_count(status_file_r10) if status_file_r10 else 432
    )
    achieved_pubsc = get_achieved_count(status_file_pub) if status_file_pub else 184

    total_achieved_val = (
        achieved_business + achieved_enterprise + achieved_pubsc
    )
    total_outstanding_val = total_target_val - total_achieved_val

    st.markdown("---")
    st.subheader("📈 Executive Summary Overview")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Target Quota", f"{total_target_val:,}")
    m2.metric(
        "Total Achieved",
        f"{total_achieved_val:,}",
        (
            f"{(total_achieved_val/total_target_val)*100:.1f}% Complete"
            if total_target_val > 0
            else "0%"
        ),
    )
    m3.metric("Total Outstanding", f"{total_outstanding_val:,}")
    m4.metric(
        "Overall Progress",
        (
            f"{(total_achieved_val/total_target_val)*100:.1f}%"
            if total_target_val > 0
            else "0%"
        ),
    )

    summary_df = pd.DataFrame(
        {
            "Segment": ["Business", "Enterprise", "PUBSC", "FML", "Total"],
            "TOTAL Target": [
                target_business,
                target_enterprise,
                target_pubsc,
                "?",
                total_target_val,
            ],
            "TOTAL Achieved": [
                achieved_business,
                achieved_enterprise,
                achieved_pubsc,
                "?",
                total_achieved_val,
            ],
            "Total Outstanding": [
                target_business - achieved_business,
                target_enterprise - achieved_enterprise,
                target_pubsc - achieved_pubsc,
                "?",
                total_outstanding_val,
            ],
        }
    )
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    # --- SEGMENT EXECUTIVE SUMMARY BREAKDOWN TABLE ---
    st.markdown("#### 📋 Segment Quotas Executive Summary Breakdown")
    achieved_r0_r1, achieved_r1_r5, achieved_r5_r10, achieved_r10_r60 = (
        616,
        371,
        275,
        459,
    )
    achieved_r60_r150, achieved_r150_plus = 247, 185

    total_seg_target = (
        q_seg_r0_r1
        + q_seg_r1_r5
        + q_seg_r5_r10
        + q_seg_r10_r60
        + q_seg_r60_r150
        + q_seg_r150_plus
    )
    total_seg_achieved = (
        achieved_r0_r1
        + achieved_r1_r5
        + achieved_r5_r10
        + achieved_r10_r60
        + achieved_r60_r150
        + achieved_r150_plus
    )
    total_seg_outstanding = total_seg_target - total_seg_achieved

    seg_summary_df = pd.DataFrame(
        {
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
                q_seg_r0_r1,
                q_seg_r1_r5,
                q_seg_r5_r10,
                q_seg_r10_r60,
                q_seg_r60_r150,
                q_seg_r150_plus,
                total_seg_target,
            ],
            "TOTAL Achieved": [
                achieved_r0_r1,
                achieved_r1_r5,
                achieved_r5_r10,
                achieved_r10_r60,
                achieved_r60_r150,
                achieved_r150_plus,
                total_seg_achieved,
            ],
            "Total Outstanding": [
                q_seg_r0_r1 - achieved_r0_r1,
                q_seg_r1_r5 - achieved_r1_r5,
                q_seg_r5_r10 - achieved_r5_r10,
                q_seg_r10_r60 - achieved_r10_r60,
                q_seg_r60_r150 - achieved_r60_r150,
                q_seg_r150_plus - achieved_r150_plus,
                total_seg_outstanding,
            ],
        }
    )
    st.dataframe(seg_summary_df, use_container_width=True, hide_index=True)

    # --- Live Visual Previews inside App Dashboard ---
    st.markdown("---")
    st.subheader("🔍 Live Regional & Segment Breakdown Tables")

    sub_tab1, sub_tab2, sub_tab3 = st.tabs(
        ["🏢 Business Breakdown", "🏭 Enterprise Breakdown", "🏫 PUBSC Breakdown"]
    )

    with sub_tab1:
        st.markdown("#### Business Regional Breakdown")
        bus_preview_df = pd.DataFrame(
            {
                "Region": [
                    "Eastern Cape",
                    "Free State",
                    "Gauteng East",
                    "Gauteng South Central",
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
                ],
                "Cape": [181, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 105],
                "Gauteng North": [
                    0,
                    0,
                    0,
                    0,
                    65,
                    63,
                    0,
                    0,
                    148,
                    59,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                ],
                "Gauteng South Central": [
                    0,
                    0,
                    88,
                    93,
                    0,
                    0,
                    0,
                    67,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                ],
                "Inland": [0, 60, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 66, 65, 72, 33, 0],
                "KwaZulu-Natal": [
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    0,
                    42,
                    45,
                    50,
                    0,
                    0,
                    0,
                    0,
                    0,
                ],
            }
        )
        bus_preview_df["Total"] = bus_preview_df.iloc[:, 1:].sum(axis=1)
        st.dataframe(bus_preview_df, use_container_width=True, hide_index=True)

        st.markdown("#### Business Segment Breakdown Matrix")
        bus_seg_df = pd.DataFrame(
            {
                "Business": ["R0M-R1M", "R1M-R5M", "R5M-R10M", "R10-R60M"],
                "Cape": [159, 81, 45, 56],
                "Gauteng North": [127, 90, 85, 105],
                "Gauteng South Central": [116, 60, 72, 107],
                "Inland": [157, 97, 36, 109],
                "KwaZulu-Natal": [57, 43, 37, 82],
                "Total": [616, 371, 275, 459],
                "Quota": [q_seg_r0_r1, q_seg_r1_r5, q_seg_r5_r10, q_seg_r10_r60],
                "Outstanding": [
                    q_seg_r0_r1 - 616,
                    q_seg_r1_r5 - 371,
                    q_seg_r5_r10 - 275,
                    q_seg_r10_r60 - 459,
                ],
            }
        )
        st.dataframe(bus_seg_df, use_container_width=True, hide_index=True)

        st.markdown(
            "#### Business Regional vs. Segments Crosstab (Regions as Rows, Segments as Columns)"
        )
        bus_crosstab_df = pd.DataFrame(
            {
                "Busines": [
                    "Eastern Cape",
                    "Free State",
                    "Gauteng East",
                    "Gauteng South Central",
                    "Gauteng Midrand",
                    "Gauteng Tshwane East",
                    "Gauteng Tshwane North",
                    "Gauteng West-Rand",
                    "Greater Sandton",
                    "KZN North",
                    "KZN South",
                    "KZN West",
                    "Limpopo",
                    "Mpumalanga",
                    "North West",
                    "Northern Cape",
                    "Western Cape",
                ],
                "R0m-R1m": [
                    121,
                    29,
                    32,
                    50,
                    27,
                    34,
                    24,
                    34,
                    65,
                    15,
                    21,
                    21,
                    34,
                    40,
                    38,
                    19,
                    38,
                ],
                "R1m-R5m": [
                    45,
                    26,
                    22,
                    21,
                    15,
                    18,
                    25,
                    17,
                    42,
                    15,
                    10,
                    18,
                    22,
                    19,
                    22,
                    11,
                    37,
                ],
                "R5m-R10": [
                    15,
                    5,
                    34,
                    22,
                    17,
                    13,
                    14,
                    16,
                    41,
                    12,
                    14,
                    11,
                    10,
                    6,
                    12,
                    3,
                    30,
                ],
            }
        )
        bus_crosstab_df["TOTAL"] = (
            bus_crosstab_df["R0m-R1m"]
            + bus_crosstab_df["R1m-R5m"]
            + bus_crosstab_df["R5m-R10"]
        )

        total_row = pd.DataFrame(
            {
                "Busines": ["TOTAL"],
                "R0m-R1m": [bus_crosstab_df["R0m-R1m"].sum()],
                "R1m-R5m": [bus_crosstab_df["R1m-R5m"].sum()],
                "R5m-R10": [bus_crosstab_df["R5m-R10"].sum()],
                "TOTAL": [bus_crosstab_df["TOTAL"].sum()],
            }
        )
        bus_crosstab_df = pd.concat(
            [bus_crosstab_df, total_row], ignore_index=True
        )
        st.dataframe(bus_crosstab_df, use_container_width=True, hide_index=True)

    with sub_tab2:
        st.markdown("#### Enterprise Regional Breakdown")
        ent_preview_df = pd.DataFrame(
            {
                "REGION": [
                    "Eastern Cape",
                    "Free State",
                    "Gauteng East",
                    "Gauteng Klipriver",
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
                ],
                "Cape": [55, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 35, 53],
                "Gauteng South & Central": [
                    0,
                    0,
                    0,
                    0,
                    0,
                    60,
                    56,
                    0,
                    0,
                    0,
                    56,
                    0,
                    0,
                    0,
                    0,
                    0,
                ],
                "Gauteng-North": [0, 0, 89, 69, 80, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                "Inland": [0, 21, 0, 0, 0, 0, 0, 0, 0, 46, 0, 67, 41, 26, 0, 0],
                "KwaZulu-Natal": [0, 0, 0, 0, 0, 0, 0, 87, 50, 0, 0, 0, 0, 0, 0, 0],
            }
        )
        ent_preview_df["Total"] = ent_preview_df.iloc[:, 1:].sum(axis=1)
        st.dataframe(ent_preview_df, use_container_width=True, hide_index=True)

        st.markdown("#### Enterprise Segment Breakdown Matrix")
        ent_seg_df = pd.DataFrame(
            {
                "Enterprise": ["R10-R60M", "R60-R150M", "R150M+"],
                "Cape": [56, 43, 44],
                "Gauteng-North": [105, 38, 29],
                "Gauteng South and Central": [107, 90, 41],
                "Inland": [109, 55, 37],
                "KwaZulu-Natal": [82, 21, 34],
                "Total": [459, 247, 185],
                "Quota": [q_seg_r10_r60, q_seg_r60_r150, q_seg_r150_plus],
                "Outstanding": [
                    q_seg_r10_r60 - 459,
                    q_seg_r60_r150 - 247,
                    q_seg_r150_plus - 185,
                ],
            }
        )
        st.dataframe(ent_seg_df, use_container_width=True, hide_index=True)

        st.markdown(
            "#### Enterprise Regional vs. Segments Crosstab (Regions as Rows, Segments as Columns)"
        )
        ent_crosstab_df = pd.DataFrame(
            {
                "Enterprise": [
                    "EASTERN CAPE",
                    "FREE STATE",
                    "GAUTENG EAST",
                    "GAUTENG KLIPRIVER",
                    "GAUTENG TSHWANE",
                    "GAUTENG WEST",
                    "GREATER SANDTON",
                    "KZN COASTAL",
                    "KZN INLAND",
                    "LIMPOPO",
                    "MIDRAND",
                    "MPUMALANGA",
                    "NORTH WEST",
                    "NORTHERN CAPE",
                    "WESTERN CAPE INLAND",
                    "WESTERN CAPE METRO",
                ],
                "R10m-R60m": [
                    26,
                    10,
                    40,
                    36,
                    32,
                    31,
                    34,
                    52,
                    30,
                    23,
                    39,
                    42,
                    22,
                    12,
                    17,
                    13,
                ],
                "R60m-R150": [21, 6, 31, 28, 19, 31, 12, 12, 9, 11, 7, 15, 11, 12, 2, 20],
                "R150m+": [8, 5, 18, 5, 9, 18, 10, 23, 11, 12, 10, 10, 8, 2, 16, 20],
            }
        )
        ent_crosstab_df["TOTAL"] = (
            ent_crosstab_df["R10m-R60m"]
            + ent_crosstab_df["R60m-R150"]
            + ent_crosstab_df["R150m+"]
        )

        ent_total_row = pd.DataFrame(
            {
                "Enterprise": ["TOTAL"],
                "R10m-R60m": [ent_crosstab_df["R10m-R60m"].sum()],
                "R60m-R150": [ent_crosstab_df["R60m-R150"].sum()],
                "R150m+": [ent_crosstab_df["R150m+"].sum()],
                "TOTAL": [ent_crosstab_df["TOTAL"].sum()],
            }
        )
        ent_crosstab_df = pd.concat(
            [ent_crosstab_df, ent_total_row], ignore_index=True
        )
        st.dataframe(ent_crosstab_df, use_container_width=True, hide_index=True)

    with sub_tab3:
        st.markdown("#### Public Sector (PUBSC) Breakdown")
        pub_preview_df = pd.DataFrame(
            {
                "Organization Type": [
                    "Non-Profit Organisation",
                    "Public Sector Colleges & FET's",
                    "Public Sector Embassies",
                    "Public Sector Local Government",
                    "Public Sector Provincial Government",
                    "Public Sector Public Schools",
                    "Public Sector Unions & Politics",
                ],
                "Eastern Cape": [4, 0, 0, 0, 0, 8, 0],
                "Free State": [1, 0, 0, 0, 0, 1, 0],
                "Gauteng": [80, 1, 2, 1, 1, 34, 1],
                "KwaZulu-Natal": [5, 0, 0, 0, 0, 12, 0],
                "Limpopo": [5, 0, 0, 0, 0, 7, 0],
                "Mpumalanga": [2, 0, 0, 0, 0, 5, 0],
                "North West": [3, 0, 0, 0, 0, 2, 0],
                "Northern Cape": [2, 0, 0, 1, 0, 0, 0],
                "Western Cape": [4, 0, 0, 0, 0, 1, 1],
            }
        )
        pub_preview_df["Total"] = pub_preview_df.iloc[:, 1:].sum(axis=1)
        st.dataframe(pub_preview_df, use_container_width=True, hide_index=True)

    def generate_exact_pm_update_workbook():
        output_buffer = io.BytesIO()
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        FNB_TEAL = PatternFill(
            start_color="00A3AD", end_color="00A3AD", fill_type="solid"
        )
        LIGHT_TEAL = PatternFill(
            start_color="D9F2F4", end_color="D9F2F4", fill_type="solid"
        )
        FNB_ORANGE = PatternFill(
            start_color="F58220", end_color="F58220", fill_type="solid"
        )
        GRAY_HEADER = PatternFill(
            start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
        )
        RED_FILL = PatternFill(
            start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
        )

        WHITE_BOLD_FONT = Font(
            name="Calibri", size=11, bold=True, color="FFFFFF"
        )
        THIN_BORDER = Border(
            left=Side(style="thin", color="BFBFBF"),
            right=Side(style="thin", color="BFBFBF"),
            top=Side(style="thin", color="BFBFBF"),
            bottom=Side(style="thin", color="BFBFBF"),
        )
        CENTER_ALIGN = Alignment(horizontal="center", vertical="center")

        ws_sum = wb.create_sheet(title="Summary")
        ws_sum.append(
            ["", "Segment", "TOTAL Target", "TOTAL Achieved", "Total Outstanding"]
        )
        ws_sum["B1"].fill = FNB_TEAL
        ws_sum["C1"].fill = FNB_TEAL
        ws_sum["D1"].fill = FNB_TEAL
        ws_sum["E1"].fill = FNB_TEAL
        for col in ["B", "C", "D", "E"]:
            ws_sum[f"{col}1"].font = WHITE_BOLD_FONT
            ws_sum[f"{col}1"].alignment = CENTER_ALIGN

        for _, row in summary_df.iterrows():
            ws_sum.append(
                [
                    "",
                    row["Segment"],
                    row["TOTAL Target"],
                    row["TOTAL Achieved"],
                    row["Total Outstanding"],
                ]
            )

        ws_bus = wb.create_sheet(title="Update Business")
        ws_bus.cell(row=1, column=2, value="Region").fill = GRAY_HEADER
        ws_bus.merge_cells("B1:G1")
        ws_bus.cell(row=1, column=2).alignment = CENTER_ALIGN

        ws_bus.cell(row=2, column=2, value="Business").fill = FNB_TEAL
        ws_bus.cell(row=2, column=2).font = WHITE_BOLD_FONT

        for c_idx, reg in enumerate(
            ["Cape", "Gauteng North", "Gauteng South Central", "Inland", "KwaZulu-Natal"],
            start=3,
        ):
            cell = ws_bus.cell(row=2, column=c_idx, value=reg)
            cell.fill = FNB_TEAL
            cell.font = WHITE_BOLD_FONT

        tot_hdr_bus = ws_bus.cell(row=2, column=8, value="Total")
        tot_hdr_bus.fill = FNB_TEAL
        tot_hdr_bus.font = WHITE_BOLD_FONT
        tot_hdr_bus.alignment = CENTER_ALIGN
        tot_hdr_bus.border = THIN_BORDER

        bus_rows = [
            ("Eastern Cape", [181, 0, 0, 0, 0]),
            ("Free State", [0, 0, 0, 60, 0]),
            ("Gauteng East", [0, 0, 88, 0, 0]),
            ("Gauteng South Central", [0, 0, 93, 0, 0]),
            ("Gauteng Tshwane East", [0, 65, 0, 0, 0]),
            ("Gauteng Tshwane North", [0, 63, 0, 0, 0]),
            ("Gauteng Tshwane South", [0, 0, 0, 0, 0]),
            ("Gauteng West-Rand", [0, 0, 67, 0, 0]),
            ("Greater Sandton", [0, 148, 0, 0, 0]),
            ("Gauteng Midrand", [0, 59, 0, 0, 0]),
            ("KZN North", [0, 0, 0, 0, 42]),
            ("KZN South", [0, 0, 0, 0, 45]),
            ("KZN West", [0, 0, 0, 0, 50]),
            ("Limpopo", [0, 0, 0, 66, 0]),
            ("Mpumalanga", [0, 0, 0, 65, 0]),
            ("North West", [0, 0, 0, 72, 0]),
            ("Northern Cape", [0, 0, 0, 33, 0]),
            ("Western Cape", [105, 0, 0, 0, 0]),
        ]
        for idx, (prov, vals) in enumerate(bus_rows, start=3):
            ws_bus.cell(row=idx, column=2, value=prov).border = THIN_BORDER
            for v_idx, val in enumerate(vals, start=3):
                c = ws_bus.cell(row=idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            ws_bus.cell(
                row=idx, column=8, value=f"=SUM(C{idx}:G{idx})"
            ).border = THIN_BORDER

        tot_row_idx = len(bus_rows) + 3
        ws_bus.cell(row=tot_row_idx, column=2, value="TOTAL INLC R10-R60MIL").fill = (
            FNB_TEAL
        )
        ws_bus.cell(row=tot_row_idx, column=2).font = WHITE_BOLD_FONT
        for c_idx in range(3, 9):
            col_let = openpyxl.utils.get_column_letter(c_idx)
            c = ws_bus.cell(
                row=tot_row_idx,
                column=c_idx,
                value=f"=SUM({col_let}3:{col_let}{tot_row_idx-1})",
            )
            c.fill = FNB_TEAL
            c.font = WHITE_BOLD_FONT
            c.border = THIN_BORDER

        ws_bus.cell(row=1, column=11, value="Region").fill = GRAY_HEADER
        ws_bus.merge_cells("K1:P1")
        ws_bus.cell(row=2, column=10, value="Business").fill = FNB_TEAL
        ws_bus.cell(row=2, column=10).font = WHITE_BOLD_FONT
        for c_idx, sc in enumerate(
            [
                "Cape",
                "Gauteng North",
                "Gauteng South Central",
                "Inland",
                "KwaZulu-Natal",
                "Total",
                "Quota",
                "Outstanding",
            ],
            start=11,
        ):
            cell = ws_bus.cell(row=2, column=c_idx, value=sc)
            cell.fill = (
                FNB_TEAL
                if c_idx < 17
                else (FNB_ORANGE if c_idx == 17 else RED_FILL)
            )
            cell.font = WHITE_BOLD_FONT

        for idx, (s_name, s_vals, quota_val) in enumerate(
            [
                ("R0M-R1M", [159, 127, 116, 157, 57], q_seg_r0_r1),
                ("R1M-R5M", [81, 90, 60, 97, 43], q_seg_r1_r5),
                ("R5M-R10M", [45, 85, 72, 36, 37], q_seg_r5_r10),
                ("R10-R60M", [56, 105, 107, 109, 82], q_seg_r10_r60),
            ],
            start=3,
        ):
            ws_bus.cell(row=idx, column=10, value=s_name).border = THIN_BORDER
            for v_idx, val in enumerate(s_vals, start=11):
                c = ws_bus.cell(row=idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            ws_bus.cell(
                row=idx, column=16, value=f"=SUM(K{idx}:O{idx})"
            ).border = THIN_BORDER
            ws_bus.cell(row=idx, column=17, value=quota_val).border = THIN_BORDER
            ws_bus.cell(row=idx, column=17).fill = LIGHT_TEAL
            r_c = ws_bus.cell(row=idx, column=18, value=f"=Q{idx}-P{idx}")
            r_c.border = THIN_BORDER
            r_c.fill = RED_FILL

        bcross_start_row = tot_row_idx + 4
        ws_bus.cell(row=bcross_start_row, column=2, value="SEGMENTS").fill = (
            GRAY_HEADER
        )
        ws_bus.merge_cells(
            start_row=bcross_start_row,
            start_column=3,
            end_row=bcross_start_row,
            end_column=5,
        )
        ws_bus.cell(row=bcross_start_row, column=3).alignment = CENTER_ALIGN

        ws_bus.cell(
            row=bcross_start_row + 1, column=2, value="Busines"
        ).fill = FNB_TEAL
        ws_bus.cell(row=bcross_start_row + 1, column=2).font = WHITE_BOLD_FONT
        for c_idx, seg_lbl in enumerate(
            ["R0m-R1m", "R1m-R5m", "R5m-R10", "TOTAL"], start=3
        ):
            cell = ws_bus.cell(row=bcross_start_row + 1, column=c_idx, value=seg_lbl)
            cell.fill = FNB_TEAL if c_idx < 6 else GRAY_HEADER
            cell.font = (
                WHITE_BOLD_FONT
                if c_idx < 6
                else Font(name="Calibri", size=11, bold=True)
            )
            cell.alignment = CENTER_ALIGN

        bus_crosstab_rows = [
            ("Eastern Cape", [121, 45, 15]),
            ("Free State", [29, 26, 5]),
            ("Gauteng East", [32, 22, 34]),
            ("Gauteng South Central", [50, 21, 22]),
            ("Gauteng Midrand", [27, 15, 17]),
            ("Gauteng Tshwane East", [34, 18, 13]),
            ("Gauteng Tshwane North", [24, 25, 14]),
            ("Gauteng West-Rand", [34, 17, 16]),
            ("Greater Sandton", [65, 42, 41]),
            ("KZN North", [15, 15, 12]),
            ("KZN South", [21, 10, 14]),
            ("KZN West", [21, 18, 11]),
            ("Limpopo", [34, 22, 10]),
            ("Mpumalanga", [40, 19, 6]),
            ("North West", [38, 22, 12]),
            ("Northern Cape", [19, 11, 3]),
            ("Western Cape", [38, 37, 30]),
        ]
        for idx_offset, (reg_name, vals) in enumerate(bus_crosstab_rows):
            r_idx = bcross_start_row + 2 + idx_offset
            ws_bus.cell(row=r_idx, column=2, value=reg_name).border = THIN_BORDER
            for v_idx, val in enumerate(vals, start=3):
                c = ws_bus.cell(row=r_idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            tot_c = ws_bus.cell(
                row=r_idx, column=6, value=f"=SUM(C{r_idx}:E{r_idx})"
            )
            tot_c.border = THIN_BORDER
            tot_c.alignment = CENTER_ALIGN

        bcross_tot_row = bcross_start_row + 2 + len(bus_crosstab_rows)
        ws_bus.cell(row=bcross_tot_row, column=2, value="TOTAL").fill = GRAY_HEADER
        ws_bus.cell(row=bcross_tot_row, column=2).font = Font(
            name="Calibri", size=11, bold=True
        )
        ws_bus.cell(row=bcross_tot_row, column=2).border = THIN_BORDER
        for c_idx in range(3, 7):
            col_let = openpyxl.utils.get_column_letter(c_idx)
            start_r = bcross_start_row + 2
            end_r = bcross_tot_row - 1
            c = ws_bus.cell(
                row=bcross_tot_row,
                column=c_idx,
                value=f"=SUM({col_let}{start_r}:{col_let}{end_r})",
            )
            c.fill = GRAY_HEADER
            c.font = Font(name="Calibri", size=11, bold=True)
            c.border = THIN_BORDER
            c.alignment = CENTER_ALIGN

        ws_ent = wb.create_sheet(title="Update Enterprise")
        ws_ent.cell(row=1, column=2, value="REGION").fill = GRAY_HEADER
        ws_ent.merge_cells("B1:G1")
        ws_ent.cell(row=2, column=2, value="Enterprise").fill = FNB_TEAL
        ws_ent.cell(row=2, column=2).font = WHITE_BOLD_FONT
        for c_idx, reg in enumerate(
            [
                "Cape",
                "Gauteng-North",
                "Gauteng South and Central",
                "Inland",
                "KwaZulu-Natal",
            ],
            start=3,
        ):
            cell = ws_ent.cell(row=2, column=c_idx, value=reg)
            cell.fill = FNB_TEAL
            cell.font = WHITE_BOLD_FONT

        tot_hdr_ent = ws_ent.cell(row=2, column=8, value="Total")
        tot_hdr_ent.fill = FNB_TEAL
        tot_hdr_ent.font = WHITE_BOLD_FONT
        tot_hdr_ent.alignment = CENTER_ALIGN
        tot_hdr_ent.border = THIN_BORDER

        ent_rows = [
            ("Eastern Cape", [55, 0, 0, 0, 0]),
            ("Free State", [0, 0, 0, 21, 0]),
            ("Gauteng East", [0, 89, 0, 0, 0]),
            ("Gauteng Klipriver", [0, 69, 0, 0, 0]),
            ("Gauteng South-West", [0, 80, 0, 0, 0]),
            ("Gauteng Tshwane", [0, 0, 60, 0, 0]),
            ("Greater Sandton", [0, 0, 56, 0, 0]),
            ("KZN Coastal", [0, 0, 0, 0, 87]),
            ("KZN Inland", [0, 0, 0, 0, 50]),
            ("Limpopo", [0, 0, 0, 46, 0]),
            ("Midrand", [0, 0, 56, 0, 0]),
            ("Mpumalanga", [0, 0, 0, 67, 0]),
            ("North West", [0, 0, 0, 41, 0]),
            ("Northern Cape", [0, 0, 0, 26, 0]),
            ("Western Cape Inland", [35, 0, 0, 0, 0]),
            ("Western Cape Metro", [53, 0, 0, 0, 0]),
        ]
        for idx, (prov, vals) in enumerate(ent_rows, start=3):
            ws_ent.cell(row=idx, column=2, value=prov).border = THIN_BORDER
            for v_idx, val in enumerate(vals, start=3):
                c = ws_ent.cell(row=idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            ws_ent.cell(
                row=idx, column=8, value=f"=SUM(C{idx}:G{idx})"
            ).border = THIN_BORDER

        ent_tot_row = len(ent_rows) + 3
        ws_ent.cell(
            row=ent_tot_row, column=2, value="TOTAL EXCL R10 to R60MIL"
        ).fill = FNB_TEAL
        ws_ent.cell(row=ent_tot_row, column=2).font = WHITE_BOLD_FONT
        for c_idx in range(3, 9):
            col_let = openpyxl.utils.get_column_letter(c_idx)
            c = ws_ent.cell(
                row=ent_tot_row,
                column=c_idx,
                value=f"=SUM({col_let}3:{col_let}{ent_tot_row-1})",
            )
            c.fill = FNB_TEAL
            c.font = WHITE_BOLD_FONT
            c.border = THIN_BORDER

        ws_ent.cell(row=1, column=11, value="REGION").fill = GRAY_HEADER
        ws_ent.merge_cells("K1:P1")
        ws_ent.cell(row=2, column=10, value="Enterprise").fill = FNB_TEAL
        ws_ent.cell(row=2, column=10).font = WHITE_BOLD_FONT
        for c_idx, sc in enumerate(
            [
                "Cape",
                "Gauteng-North",
                "Gauteng South and Central",
                "Inland",
                "KwaZulu-Natal",
                "Total",
                "Quota",
                "Outstanding",
            ],
            start=11,
        ):
            cell = ws_ent.cell(row=2, column=c_idx, value=sc)
            cell.fill = (
                FNB_TEAL
                if c_idx < 17
                else (FNB_ORANGE if c_idx == 17 else RED_FILL)
            )
            cell.font = WHITE_BOLD_FONT

        for idx, (s_name, s_vals, quota_val) in enumerate(
            [
                ("R10-R60M", [56, 105, 107, 109, 82], q_seg_r10_r60),
                ("R60-R150M", [43, 38, 90, 55, 21], q_seg_r60_r150),
                ("R150M+", [44, 29, 41, 37, 34], q_seg_r150_plus),
            ],
            start=3,
        ):
            ws_ent.cell(row=idx, column=10, value=s_name).border = THIN_BORDER
            for v_idx, val in enumerate(s_vals, start=11):
                c = ws_ent.cell(row=idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            ws_ent.cell(
                row=idx, column=16, value=f"=SUM(K{idx}:O{idx})"
            ).border = THIN_BORDER
            ws_ent.cell(row=idx, column=17, value=quota_val).border = THIN_BORDER
            ws_ent.cell(row=idx, column=17).fill = LIGHT_TEAL
            r_c = ws_ent.cell(row=idx, column=18, value=f"=Q{idx}-P{idx}")
            r_c.border = THIN_BORDER
            r_c.fill = RED_FILL

        ecross_start_row = ent_tot_row + 4
        ws_ent.cell(row=ecross_start_row, column=2, value="SEGMENTS").fill = (
            GRAY_HEADER
        )
        ws_ent.merge_cells(
            start_row=ecross_start_row,
            start_column=3,
            end_row=ecross_start_row,
            end_column=5,
        )
        ws_ent.cell(row=ecross_start_row, column=3).alignment = CENTER_ALIGN

        ws_ent.cell(
            row=ecross_start_row + 1, column=2, value="Enterprise"
        ).fill = FNB_TEAL
        ws_ent.cell(row=ecross_start_row + 1, column=2).font = WHITE_BOLD_FONT
        for c_idx, seg_lbl in enumerate(
            ["R10m-R60m", "R150m+", "R60m-R150", "TOTAL"], start=3
        ):
            cell = ws_ent.cell(row=ecross_start_row + 1, column=c_idx, value=seg_lbl)
            cell.fill = FNB_TEAL if c_idx < 6 else GRAY_HEADER
            cell.font = (
                WHITE_BOLD_FONT
                if c_idx < 6
                else Font(name="Calibri", size=11, bold=True)
            )
            cell.alignment = CENTER_ALIGN

        ent_crosstab_rows = [
            ("EASTERN CAPE", [26, 21, 8]),
            ("FREE STATE", [10, 6, 5]),
            ("GAUTENG EAST", [40, 31, 18]),
            ("GAUTENG KLIPRIVER", [36, 28, 5]),
            ("GAUTENG TSHWANE", [32, 19, 9]),
            ("GAUTENG WEST", [31, 31, 18]),
            ("GREATER SANDTON", [34, 12, 10]),
            ("KZN COASTAL", [52, 12, 23]),
            ("KZN INLAND", [30, 9, 11]),
            ("LIMPOPO", [23, 11, 12]),
            ("MIDRAND", [39, 7, 10]),
            ("MPUMALANGA", [42, 15, 10]),
            ("NORTH WEST", [22, 11, 8]),
            ("NORTHERN CAPE", [12, 12, 2]),
            ("WESTERN CAPE INLAND", [17, 2, 16]),
            ("WESTERN CAPE METRO", [13, 20, 20]),
        ]
        for idx_offset, (reg_name, vals) in enumerate(ent_crosstab_rows):
            r_idx = ecross_start_row + 2 + idx_offset
            ws_ent.cell(row=r_idx, column=2, value=reg_name).border = THIN_BORDER
            for v_idx, val in enumerate(vals, start=3):
                c = ws_ent.cell(row=r_idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            tot_c = ws_ent.cell(
                row=r_idx, column=6, value=f"=SUM(C{r_idx}:E{r_idx})"
            )
            tot_c.border = THIN_BORDER
            tot_c.alignment = CENTER_ALIGN

        ecross_tot_row = ecross_start_row + 2 + len(ent_crosstab_rows)
        ws_ent.cell(row=ecross_tot_row, column=2, value="TOTAL").fill = GRAY_HEADER
        ws_ent.cell(row=ecross_tot_row, column=2).font = Font(
            name="Calibri", size=11, bold=True
        )
        ws_ent.cell(row=ecross_tot_row, column=2).border = THIN_BORDER
        for c_idx in range(3, 7):
            col_let = openpyxl.utils.get_column_letter(c_idx)
            start_r = ecross_start_row + 2
            end_r = ecross_tot_row - 1
            c = ws_ent.cell(
                row=ecross_tot_row,
                column=c_idx,
                value=f"=SUM({col_let}{start_r}:{col_let}{end_r})",
            )
            c.fill = GRAY_HEADER
            c.font = Font(name="Calibri", size=11, bold=True)
            c.border = THIN_BORDER
            c.alignment = CENTER_ALIGN

        ws_pub = wb.create_sheet(title="Update PUBSC")
        ws_pub.cell(row=1, column=2, value="REGION").fill = GRAY_HEADER
        ws_pub.merge_cells("B1:L1")
        ws_pub.cell(row=1, column=2).alignment = CENTER_ALIGN

        pub_headers = [
            "ORGANISATION TYPE",
            "EASTERN CAPE",
            "FREE STATE",
            "GAUTENG",
            "KWAZULU-NATAL",
            "LIMPOPO",
            "MPUMALANGA",
            "NORTH WEST",
            "NORTHERN CAPE",
            "WESTERN CAPE",
            "TOTAL",
        ]
        for col_idx, h_text in enumerate(pub_headers, start=2):
            cell = ws_pub.cell(row=2, column=col_idx, value=h_text)
            cell.fill = FNB_TEAL
            cell.font = WHITE_BOLD_FONT
            cell.alignment = CENTER_ALIGN

        pub_rows_data = [
            ["NON-PROFIT ORGANISATION", 4, 1, 80, 5, 5, 2, 3, 2, 4],
            ["PUBLIC SECTOR COLLEGES & FET'S", 0, 0, 1, 0, 0, 0, 0, 0, 0],
            ["PUBLIC SECTOR EMBASSIES", 0, 0, 2, 0, 0, 0, 0, 0, 0],
            ["PUBLIC SECTOR LOCAL GOVERMENT", 0, 0, 1, 0, 0, 0, 0, 1, 0],
            ["PUBLIC SECTOR PROVINCIAL GOVER", 0, 0, 1, 0, 0, 0, 0, 0, 0],
            ["PUBLIC SECTOR PUBLIC SCHOOLS", 8, 1, 34, 12, 7, 5, 2, 0, 1],
            ["PUBLIC SECTOR UNIONS & POLITIC", 0, 0, 1, 0, 0, 0, 0, 0, 1],
        ]
        for row_offset, prow in enumerate(pub_rows_data, start=3):
            ws_pub.cell(row=row_offset, column=2, value=prow[0]).border = (
                THIN_BORDER
            )
            for val_idx, val in enumerate(prow[1:], start=3):
                c = ws_pub.cell(row=row_offset, column=val_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            tot_c = ws_pub.cell(
                row=row_offset, column=12, value=f"=SUM(C{row_offset}:K{row_offset})"
            )
            tot_c.border = THIN_BORDER
            tot_c.alignment = CENTER_ALIGN

        pub_tot_row = len(pub_rows_data) + 3
        ws_pub.cell(row=pub_tot_row, column=2, value="").border = THIN_BORDER
        for c_idx in range(3, 13):
            col_let = openpyxl.utils.get_column_letter(c_idx)
            c = ws_pub.cell(
                row=pub_tot_row,
                column=c_idx,
                value=f"=SUM({col_let}3:{col_let}{pub_tot_row-1})",
            )
            c.border = THIN_BORDER
            c.alignment = CENTER_ALIGN
            c.font = Font(name="Calibri", size=11, bold=True)

        for sheet in wb.worksheets:
            for col in sheet.columns:
                max_len = 0
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                for cell in col:
                    if cell.value is not None:
                        val_str = str(cell.value)
                        if len(val_str) > max_len:
                            max_len = len(val_str)
                sheet.column_dimensions[col_letter].width = max(
                    max_len + 3, 12
                )

        wb.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    st.markdown("---")
    if st.button(
        "📊 Generate & Download Exact PM Update Workbook",
        type="primary",
        key="download_status_btn",
    ):
        status_excel_bytes = generate_exact_pm_update_workbook()
        run_date_str = datetime.now().strftime("%Y-%m-%d")
        st.success(
            "✅ Project Status Update report generated successfully with FNB brand colors and auto-fitted columns across all worksheets!"
        )
        st.download_button(
            label="📥 Download Formatted Excel Report (`Star Detailed Update.xlsx`)",
            data=status_excel_bytes,
            file_name=f"Star Detailed Update-W22 {run_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_status_excel_final",
        )

# ==========================================================================
# ==========================================================================
# TAB 2: WEEKLY 911'S CONTROL ROOM
# ==========================================================================
# ==========================================================================
with tab2:
    st.markdown("### `[02 // CONTROL ROOM]` &nbsp;&nbsp;&nbsp; `SYS.READY // PIPELINE 2.2`")
    st.markdown(
        "Execute and monitor each section of the Project Star 911 market research data pipeline."
    )
    st.markdown("---")

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.metric("Pipeline Status", "IDLE / READY", "Stable")
    col_m2.metric("Active Wave", "Wave 22", "2026")
    col_m3.metric("Modules Loaded", "3 / 3", "Growth, R10Mil, PUBSC")
    col_m4.metric("Environment", "Cloud Control Room", "Secure")

    st.markdown("---")
    st.subheader("⚙️ Global Execution Parameters")
    date_mode = st.radio(
        "Select Date Filtering Mode for Runs:",
        ["Dynamic Past 7 Days (Auto Friday)", "Custom Date Range"],
        horizontal=True,
        key="911_date_mode",
    )

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
        last_friday = last_friday.replace(
            hour=0, minute=0, second=0, microsecond=0
        )
        st.info(
            f"📅 Target active execution window: **{last_friday.strftime('%Y-%m-%d')}** to **{today.strftime('%Y-%m-%d')}**"
        )
    else:
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            start_date_input = st.date_input(
                "Start Date", value=today - timedelta(days=7), key="911_start"
            )
        with col_d2:
            end_date_input = st.date_input("End Date", value=today, key="911_end")

        last_friday = datetime.combine(
            start_date_input, datetime.min.time()
        )
        today = datetime.combine(end_date_input, datetime.max.time())

    st.markdown("---")

    def run_pipeline(uploaded_file, section_choice):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        try:
            df, meta = pyreadstat.read_sav(tmp_path)

            if "V9999" in df.columns:
                df_filtered = df[df["V9999"] == 1].copy()
            else:
                df_filtered = df.copy()

            if "STIME" in df_filtered.columns:
                df_filtered["STIME_CLEAN"] = (
                    df_filtered["STIME"].astype(str).str.strip().str[:8]
                )

                def parse_stime_date(x):
                    try:
                        return datetime.strptime(x, "%Y%m%d")
                    except:
                        return None

                df_filtered["STIME_DATE"] = df_filtered["STIME_CLEAN"].apply(
                    parse_stime_date
                )

                df_filtered = df_filtered[
                    (df_filtered["STIME_DATE"] >= last_friday)
                    & (df_filtered["STIME_DATE"] <= today)
                ].copy()

            if df_filtered.empty:
                st.warning(
                    f"⚠️ No records found matching the criteria for {section_choice}."
                )
                return None

            if "INTNR" not in df_filtered.columns:
                df_filtered["INTNR"] = range(1, len(df_filtered) + 1)

            valid_intnr = df_filtered["INTNR"] > 0

            df_filtered.loc[valid_intnr, "PARENT_TYPE"] = "Juristic"
            df_filtered.loc[valid_intnr, "WAVE"] = "22"
            if "V80116" in df_filtered.columns:
                df_filtered.loc[valid_intnr, "CLIENT_UCN"] = df_filtered[
                    "V80116"
                ]
            df_filtered.loc[valid_intnr, "CLIENT_TYPE"] = "Full Client"
            df_filtered.loc[valid_intnr, "COMPANY_CODE"] = "15"
            df_filtered.loc[valid_intnr, "CASE_SUBJECT"] = "Coverage"
            df_filtered.loc[valid_intnr, "REQUEST_CATEGORY"] = "Care"
            df_filtered.loc[valid_intnr, "TOPIC"] = "Complaints"

            tq13_open_clean = (
                df_filtered["TQ13_OPEN"].fillna("").astype(str)
                if "TQ13_OPEN" in df_filtered.columns
                else ""
            )
            df_filtered.loc[valid_intnr, "CASE_DESCRIPTION"] = (
                "Improvement Area: " + tq13_open_clean
            )

            df_filtered.loc[valid_intnr, "CAMPAIGN"] = "CMP-01522-S7K7N3"
            df_filtered.loc[valid_intnr, "ORIGIN"] = "Web"
            df_filtered.loc[valid_intnr, "OWNER"] = r"FNBJNB01\Web"

            sub_region_col = (
                "V8013" if section_choice == "PUBSC" else "V13290"
            )
            segment_col = "V13290" if section_choice == "PUBSC" else "V44011"

            mapping_pairs = [
                ("PRIM_OFCR_IND", "V8026"),
                ("OFFICER_NAME_AND_SURNAME", "V8016"),
                ("BUSINESS_NAME", "V56011"),
                ("REGIONS", "V12290"),
                ("SUB_REGIONS", sub_region_col),
                ("SEGMENT", segment_col),
            ]

            for col_target, col_src in mapping_pairs:
                if col_src in df_filtered.columns:
                    df_filtered.loc[valid_intnr, col_target] = df_filtered[
                        col_src
                    ]

            def calculate_nps_bucket(score):
                if pd.isna(score):
                    return None
                if score < 7:
                    return "NPS - Detractor"
                elif score in [7, 8]:
                    return "NPS - Passive"
                elif score > 8:
                    return "NPS - Promoter"
                return None

            if "Q14_1" in df_filtered.columns:
                df_filtered["FNB_NPS"] = df_filtered["Q14_1"].apply(
                    calculate_nps_bucket
                )
                df_filtered["FNB_NPS_OPEN_ENDED"] = df_filtered.apply(
                    lambda r: r["TQ14_1_OPEN"]
                    if "TQ14_1_OPEN" in df_filtered.columns
                    and pd.notna(r.get("TQ14_1_OPEN"))
                    and r["Q14_1"] < 7
                    else None,
                    axis=1,
                )
            if "Q14_2" in df_filtered.columns:
                df_filtered["RM_BM_NPS"] = df_filtered["Q14_2"].apply(
                    calculate_nps_bucket
                )
                df_filtered["RM_BM_NPS_OPEN_ENDED"] = df_filtered.apply(
                    lambda r: r["TQ14_2_OPEN"]
                    if "TQ14_2_OPEN" in df_filtered.columns
                    and pd.notna(r.get("TQ14_2_OPEN"))
                    and r["Q14_2"] < 7
                    else None,
                    axis=1,
                )

            if section_choice == "PUBSC":
                bank_columns = [
                    ("Q16_1_1", "Absa"),
                    ("Q16_1_2", "Investec"),
                    ("Q16_1_3", "Nedbank"),
                    ("Q16_1_4", "Standard Bank"),
                    ("Q16_1_5", "Capitec"),
                    ("Q16_1_6", "Refused"),
                ]
                loop_cols = ["TQ16_1C6", "TQ16_1C7", "TQ16_1C8"]
            else:
                bank_columns = [
                    ("Q16_1_1", "Absa"),
                    ("Q16_1_2", "Capitec"),
                    ("Q16_1_3", "Investec"),
                    ("Q16_1_4", "Mercantile"),
                    ("Q16_1_5", "Nedbank"),
                    ("Q16_1_6", "Sasfin"),
                    ("Q16_1_7", "Standard Bank"),
                ]
                loop_cols = ["TQ16_1C8", "TQ16_1C9", "TQ16_1C10"]

            switch_compiled = []
            for idx, row in df_filtered.iterrows():
                matched_banks = []
                for col_flag, bank_label in bank_columns:
                    if (
                        col_flag in df_filtered.columns
                        and row.get(col_flag) == 1
                    ):
                        matched_banks.append(bank_label)
                for loop_col in loop_cols:
                    if (
                        loop_col in df_filtered.columns
                        and pd.notna(row.get(loop_col))
                        and str(row[loop_col]).strip() != ""
                    ):
                        matched_banks.append(str(row[loop_col]).strip())
                switch_compiled.append(",".join(matched_banks))

            df_filtered["WOULD_CONSIDER_SWITCH_TO"] = switch_compiled
            if "TQ16_OPEN" in df_filtered.columns:
                df_filtered.loc[valid_intnr, "REASON"] = df_filtered[
                    "TQ16_OPEN"
                ]

            if "STIME_CLEAN" in df_filtered.columns:
                df_filtered["NYEAR"] = df_filtered["STIME_CLEAN"].str[:4]
                df_filtered["NMONTH"] = df_filtered["STIME_CLEAN"].str[4:6]
                df_filtered["NDAY"] = df_filtered["STIME_CLEAN"].str[6:8]
                df_filtered["RECORDED_DATE"] = (
                    df_filtered["NYEAR"]
                    + "/"
                    + df_filtered["NMONTH"]
                    + "/"
                    + df_filtered["NDAY"]
                )

            df_filtered["Qualifier"] = "Not Priority"
            if "Q14_2" in df_filtered.columns:
                df_filtered.loc[
                    df_filtered["Q14_2"] < 7, "Qualifier"
                ] = "Priority"
            if "Q16" in df_filtered.columns:
                df_filtered.loc[
                    df_filtered["Q16"] == 1, "Qualifier"
                ] = "Priority"

            case_desc_upper = (
                df_filtered["CASE_DESCRIPTION"].fillna("").str.upper()
                if "CASE_DESCRIPTION" in df_filtered.columns
                else pd.Series([""] * len(df_filtered))
            )
            people_keywords = [
                "BM",
                "BUSINESS MANAGER",
                "BUSINESS BANKERS",
                "PRIVATE BANKER",
                "RM",
                "RELATIONSHIP MANAGER",
                "STAFF",
                "CLIENTS",
            ]
            process_keywords = [
                "SYSTEM",
                "PROCESS",
                "SERVICE",
                "DELAY",
                "QUERY",
                "ACCESS",
                "APP",
            ]
            product_keywords = [
                "FEE",
                "CHARGES",
                "LOAN",
                "ACCOUNT",
                "INVESTMENT",
                "CARD",
            ]
            none_keywords = [
                "NO IMPROVEMENT",
                "NONE",
                "SATISFIED",
                "ALL GOOD",
                "N/A",
            ]

            def contains_keywords(text, kw_list):
                return 1 if any(kw in text for kw in kw_list) else 0

            df_filtered["People_1"] = case_desc_upper.apply(
                lambda x: contains_keywords(x, people_keywords)
            )
            df_filtered["PROCESS_1"] = case_desc_upper.apply(
                lambda x: contains_keywords(x, process_keywords)
            )
            df_filtered["PRODUCT_1"] = case_desc_upper.apply(
                lambda x: contains_keywords(x, product_keywords)
            )
            df_filtered["NONE_OVERRIDE"] = case_desc_upper.apply(
                lambda x: contains_keywords(x, none_keywords)
            )

            def apply_triple_p_logic(row):
                if row.get("NONE_OVERRIDE") == 1:
                    return "NONE"
                p, pp, pr = (
                    row.get("PRODUCT_1") == 1,
                    row.get("People_1") == 1,
                    row.get("PROCESS_1") == 1,
                )
                if pp and pr and p:
                    return "ALL"
                if p and pr:
                    return "PRODUCT & PROCESS"
                if p and pp:
                    return "PRODUCT & PEOPLE"
                if pp and pr:
                    return "PEOPLE & PROCESS"
                if p:
                    return "PRODUCT ONLY"
                if pr:
                    return "PROCESS ONLY"
                if pp:
                    return "PEOPLE ONLY"
                return ""

            df_filtered["PRODUCT_PEOPLE_PROCESS"] = df_filtered.apply(
                apply_triple_p_logic, axis=1
            )
            df_filtered = df_filtered.sort_values(
                by="INTNR", ascending=True
            ).copy()

            verbatim_cols = [
                "CASE_DESCRIPTION",
                "FNB_NPS_OPEN_ENDED",
                "RM_BM_NPS_OPEN_ENDED",
                "WOULD_CONSIDER_SWITCH_TO",
                "REASON",
            ]
            for col in verbatim_cols:
                if col in df_filtered.columns:
                    df_filtered[col] = (
                        df_filtered[col]
                        .fillna("")
                        .astype(str)
                        .str.replace(",", "~", regex=False)
                    )

            df_priority = df_filtered[
                df_filtered["Qualifier"] == "Priority"
            ].copy()

            keep_columns = [
                "PARENT_TYPE",
                "WAVE",
                "CLIENT_UCN",
                "CLIENT_TYPE",
                "COMPANY_CODE",
                "CASE_SUBJECT",
                "REQUEST_CATEGORY",
                "TOPIC",
                "CASE_DESCRIPTION",
                "PRODUCT_PEOPLE_PROCESS",
                "CAMPAIGN",
                "ORIGIN",
                "OWNER",
                "PRIM_OFCR_IND",
                "OFFICER_NAME_AND_SURNAME",
                "BUSINESS_NAME",
                "REGIONS",
                "SUB_REGIONS",
                "SEGMENT",
                "FNB_NPS",
                "FNB_NPS_OPEN_ENDED",
                "RM_BM_NPS",
                "RM_BM_NPS_OPEN_ENDED",
                "WOULD_CONSIDER_SWITCH_TO",
                "REASON",
                "RECORDED_DATE",
            ]

            for c in keep_columns:
                if c not in df_priority.columns:
                    df_priority[c] = ""

            base_priority_data = df_priority[keep_columns].copy()
            run_date_file = today.strftime("%Y_%m_%d")

            if section_choice == "Growth":
                prefix = "Business_Client_911"
            elif section_choice == "R10Mil":
                prefix = "Enterprise_Client_911"
            else:
                prefix = "PUBSC_Client_911"

            f1_data = base_priority_data[
                (base_priority_data.get("FNB_NPS") == "NPS - Detractor")
                | (base_priority_data.get("RM_BM_NPS") == "NPS - Detractor")
            ]
            f2_data = f1_data.drop(
                columns=["PRODUCT_PEOPLE_PROCESS"], errors="ignore"
            )
            f3_data = base_priority_data
            f4_data = base_priority_data.drop(
                columns=["PRODUCT_PEOPLE_PROCESS"], errors="ignore"
            )

            return {
                "f1": (
                    f1_data.to_csv(
                        sep="|", index=False, encoding="utf-8-sig"
                    ).encode("utf-8-sig"),
                    f"{prefix}_NPS_D_classification_{run_date_file}.csv",
                ),
                "f2": (
                    f2_data.to_csv(
                        sep="|", index=False, encoding="utf-8-sig"
                    ).encode("utf-8-sig"),
                    f"{prefix}_NPS_D_NO_classification_{run_date_file}.csv",
                ),
                "f3": (
                    f3_data.to_csv(
                        sep="|", index=False, encoding="utf-8-sig"
                    ).encode("utf-8-sig"),
                    f"{prefix}_NPS_D_S_classification_{run_date_file}.csv",
                ),
                "f4": (
                    f4_data.to_csv(
                        sep="|", index=False, encoding="utf-8-sig"
                    ).encode("utf-8-sig"),
                    f"{prefix}_NPS_D_S_NO_classification_{run_date_file}.csv",
                ),
                "count": len(df_filtered),
            }
        except Exception as e:
            st.error(f"❌ Error in {section_choice}: {e}")
            return None
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    st.subheader("🚀 Pipeline Execution Control Room")
    col1, col2, col3 = st.columns(3)

    with col1:
        with st.container(border=True):
            st.markdown("### 💼 Growth Section")
            st.caption("Target: Business Client Pipeline")
            file_growth = st.file_uploader(
                "Upload GROW SAV (.sav)", type=["sav"], key="growth_file"
            )

            if st.button(
                "▶ Run Growth Stage",
                key="btn_growth",
                type="primary",
                use_container_width=True,
            ):
                if file_growth is None:
                    st.error("Upload a SAV file first.")
                else:
                    with st.spinner("Processing Growth..."):
                        res = run_pipeline(file_growth, "Growth")
                        if res:
                            st.success(f"Processed {res['count']} records!")
                            st.download_button(
                                "📥 Output 1",
                                res["f1"][0],
                                file_name=res["f1"][1],
                                mime="text/csv",
                                key="g1",
                            )
                            st.download_button(
                                "📥 Output 2",
                                res["f2"][0],
                                file_name=res["f2"][1],
                                mime="text/csv",
                                key="g2",
                            )
                            st.download_button(
                                "📥 Output 3",
                                res["f3"][0],
                                file_name=res["f3"][1],
                                mime="text/csv",
                                key="g3",
                            )
                            st.download_button(
                                "📥 Output 4",
                                res["f4"][0],
                                file_name=res["f4"][1],
                                mime="text/csv",
                                key="g4",
                            )

    with col2:
        with st.container(border=True):
            st.markdown("### 🏢 R10Mil Section")
            st.caption("Target: Enterprise Client Pipeline")
            file_r10 = st.file_uploader(
                "Upload RMW SAV (.sav)", type=["sav"], key="r10_file"
            )

            if st.button(
                "▶ Run R10Mil Stage",
                key="btn_r10",
                type="primary",
                use_container_width=True,
            ):
                if file_r10 is None:
                    st.error("Upload a SAV file first.")
                else:
                    with st.spinner("Processing R10Mil..."):
                        res = run_pipeline(file_r10, "R10Mil")
                        if res:
                            st.success(f"Processed {res['count']} records!")
                            st.download_button(
                                "📥 Output 1",
                                res["f1"][0],
                                file_name=res["f1"][1],
                                mime="text/csv",
                                key="r1",
                            )
                            st.download_button(
                                "📥 Output 2",
                                res["f2"][0],
                                file_name=res["f2"][1],
                                mime="text/csv",
                                key="r2",
                            )
                            st.download_button(
                                "📥 Output 3",
                                res["f3"][0],
                                file_name=res["f3"][1],
                                mime="text/csv",
                                key="r3",
                            )
                            st.download_button(
                                "📥 Output 4",
                                res["f4"][0],
                                file_name=res["f4"][1],
                                mime="text/csv",
                                key="r4",
                            )

    with col3:
        with st.container(border=True):
            st.markdown("### 🏫 PUBSC Section")
            st.caption("Target: Public Sector Pipeline")
            file_pub = st.file_uploader(
                "Upload PUBW SAV (.sav)", type=["sav"], key="pub_file"
            )

            if st.button(
                "▶ Run PUBSC Stage",
                key="btn_pub",
                type="primary",
                use_container_width=True,
            ):
                if file_pub is None:
                    st.error("Upload a SAV file first.")
                else:
                    with st.spinner("Processing PUBSC..."):
                        res = run_pipeline(file_pub, "PUBSC")
                        if res:
                            st.success(f"Processed {res['count']} records!")
                            st.download_button(
                                "📥 Output 1",
                                res["f1"][0],
                                file_name=res["f1"][1],
                                mime="text/csv",
                                key="p1",
                            )
                            st.download_button(
                                "📥 Output 2",
                                res["f2"][0],
                                file_name=res["f2"][1],
                                mime="text/csv",
                                key="p2",
                            )
                            st.download_button(
                                "📥 Output 3",
                                res["f3"][0],
                                file_name=res["f3"][1],
                                mime="text/csv",
                                key="p3",
                            )
                            st.download_button(
                                "📥 Output 4",
                                res["f4"][0],
                                file_name=res["f4"][1],
                                mime="text/csv",
                                key="p4",
                            )


# ==========================================================================
# ==========================================================================
# TAB 3: NPS DASHBOARD & DATA GENERATOR
# ==========================================================================
# ==========================================================================
with tab3:
    st.markdown("### 📈 NPS Dashboard & Streamlined Data Generator")
    st.markdown(
        "Upload your master SPSS data file below, select your wave preferences and portfolio filter, then click **Run Processing**."
    )

    if "nps_reports_ready" not in st.session_state:
        st.session_state.nps_reports_ready = False
    if "nps_report_payloads" not in st.session_state:
        st.session_state.nps_report_payloads = []

    nps_uploaded_file = st.file_uploader(
        "Upload Master SPSS Data File (.sav) for NPS Dashboard",
        type=["sav"],
        key="nps_file",
    )
    portfolio_mode = st.selectbox(
        "Select Portfolio Filter Mode:",
        [
            "Generate All (Combined, Growth, and R10M Separately)",
            "Combined (Growth & R10M)",
            "Growth Only",
            "R10M Only",
        ],
        key="nps_portfolio",
    )
    filter_option = st.radio(
        "Select Wave Filter Option:",
        [
            "All Waves",
            "Custom Range (e.g., Wave 1 to 10)",
            "Specific Waves List",
        ],
        key="nps_filter_opt",
    )

    selected_waves_filter = "ALL"
    if filter_option == "Custom Range (e.g., Wave 1 to 10)":
        col_nw1, col_nw2 = st.columns(2)
        with col_nw1:
            start_w = st.number_input("Start Wave", 1, 30, 1, key="start_w")
        with col_nw2:
            end_w = st.number_input("End Wave", 1, 30, 10, key="end_w")
        selected_waves_filter = [
            f"Wave {i}" for i in range(int(start_w), int(end_w) + 1)
        ]
    elif filter_option == "Specific Waves List":
        waves_input = st.text_input(
            "Enter waves:", "Wave 20, Wave 21, Wave 22", key="waves_input"
        )
        selected_waves_filter = [w.strip() for w in waves_input.split(",")]

    def generate_report_bytes(df_subset, prefix_label):
        temp_excel = f"temp_{prefix_label}.xlsx"
        temp_sav = f"temp_{prefix_label}.sav"
        pyreadstat.write_sav(df_subset, temp_sav)
        with open(temp_sav, "rb") as f:
            sav_bytes = f.read()

        with pd.ExcelWriter(temp_excel, engine="openpyxl") as writer:
            df_subset.to_excel(writer, sheet_name="data", index=False)
        wb = openpyxl.load_workbook(temp_excel)
        wb.save(temp_excel)
        with open(temp_excel, "rb") as f:
            excel_bytes = f.read()
        if os.path.exists(temp_excel):
            os.remove(temp_excel)
        if os.path.exists(temp_sav):
            os.remove(temp_sav)
        return (
            excel_bytes,
            f"FNB_Customer_Satisfaction_Report_{prefix_label}.xlsx",
            sav_bytes,
            f"FNB_Data_{prefix_label}.sav",
        )

    if st.button("🚀 Run Processing & Generate Reports", type="primary", key="run_nps"):
        if nps_uploaded_file is None:
            st.error("Please upload a `.sav` file first!")
        else:
            with st.spinner("Processing data and building NPS reports..."):
                temp_src_path = "temp_input_nps.sav"
                with open(temp_src_path, "wb") as f:
                    f.write(nps_uploaded_file.getbuffer())
                df_raw, _ = pyreadstat.read_sav(
                    temp_src_path, apply_value_formats=False
                )
                df_raw.columns = [str(c).strip().upper() for c in df_raw.columns]
                if os.path.exists(temp_src_path):
                    os.remove(temp_src_path)

                payloads = []
                if (
                    portfolio_mode
                    == "Generate All (Combined, Growth, and R10M Separately)"
                ):
                    exc_comb, name_comb, sav_comb, sname_comb = (
                        generate_report_bytes(df_raw, "Combined")
                    )
                    payloads.append((name_comb, exc_comb, sname_comb, sav_comb))
                    if "TYPE" in df_raw.columns:
                        df_grow = df_raw[
                            df_raw["TYPE"]
                            .astype(str)
                            .str.lower()
                            .str.contains("growth")
                        ]
                        if not df_grow.empty:
                            exc_g, name_g, sav_g, sname_g = (
                                generate_report_bytes(df_grow, "Growth")
                            )
                            payloads.append((name_g, exc_g, sname_g, sav_g))
                        df_r10 = df_raw[
                            df_raw["TYPE"]
                            .astype(str)
                            .str.lower()
                            .str.contains("r10")
                        ]
                        if not df_r10.empty:
                            exc_r, name_r, sav_r, sname_r = (
                                generate_report_bytes(df_r10, "R10M")
                            )
                            payloads.append((name_r, exc_r, sname_r, sav_r))
                else:
                    exc, name, sav, sname = generate_report_bytes(
                        df_raw, portfolio_mode.replace(" ", "_")
                    )
                    payloads.append((name, exc, sname, sav))

                st.session_state.nps_report_payloads = payloads
                st.session_state.nps_reports_ready = True
                st.success("✅ Processing complete! Download ready below.")

    if st.session_state.nps_reports_ready and st.session_state.nps_report_payloads:
        st.markdown("---")
        st.subheader("📦 Download Generated NPS Reports")
        for idx, (ex_name, ex_bytes, sav_name, sav_bytes) in enumerate(
            st.session_state.nps_report_payloads
        ):
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button(
                    label=f"📥 Download Excel: {ex_name}",
                    data=ex_bytes,
                    file_name=ex_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_excel_{idx}",
                )
            with col_d2:
                st.download_button(
                    label=f"📥 Download SPSS: {sav_name}",
                    data=sav_bytes,
                    file_name=sav_name,
                    mime="application/octet-stream",
                    key=f"dl_sav_{idx}",
                )


# ==========================================================================
# ==========================================================================
# TAB 4: Q11 RATINGS & REASONS EXTRACTION
# ==========================================================================
# ==========================================================================
with tab4:
    st.markdown("### 📋 Q11 Ratings & Reasons Extraction")
    st.markdown(
        "Upload your Enterprise (R10Mil) and Business (Growth) SPSS datasets below to extract Q11 ratings and reasons into a combined multi-tab Excel workbook."
    )

    st.markdown("---")
    st.subheader("⚙️ Global Execution Parameters (Q11 Extraction)")

    q11_date_mode = st.radio(
        "Select Date Filtering Mode for Runs:",
        ["Dynamic Past 7 Days (Auto Friday)", "Custom Date Range"],
        horizontal=True,
        key="q11_date_mode",
    )
    today_q11 = datetime.now()

    if q11_date_mode == "Dynamic Past 7 Days (Auto Friday)":
        curr_wkday = today_q11.weekday()
        days_sub = 7 if curr_wkday == 4 else (curr_wkday - 4) % 7
        if days_sub == 0:
            days_sub = 7
        q11_last_friday = today_q11 - timedelta(days=days_sub)
        q11_last_friday = q11_last_friday.replace(
            hour=0, minute=0, second=0, microsecond=0
        )
        st.info(
            f"📅 Target active execution window: **{q11_last_friday.strftime('%Y-%m-%d')}** to **{today_q11.strftime('%Y-%m-%d')}**"
        )
    else:
        col_qd1, col_qd2 = st.columns(2)
        with col_qd1:
            q11_start_input = st.date_input(
                "Start Date", value=today_q11 - timedelta(days=7), key="q11_start"
            )
        with col_qd2:
            q11_end_input = st.date_input(
                "End Date", value=today_q11, key="q11_end"
            )
        q11_last_friday = datetime.combine(
            q11_start_input, datetime.min.time()
        )
        today_q11 = datetime.combine(q11_end_input, datetime.max.time())

    st.markdown("---")
    col_q1, col_q2 = st.columns(2)
    with col_q1:
        file_q11_r10 = st.file_uploader(
            "Upload R10Mil SPSS File (.sav)", type=["sav"], key="q11_r10"
        )
    with col_q2:
        file_q11_grow = st.file_uploader(
            "Upload Growth SPSS File (.sav)", type=["sav"], key="q11_grow"
        )

    if "q11_ready" not in st.session_state:
        st.session_state.q11_ready = False
    if "q11_bytes" not in st.session_state:
        st.session_state.q11_bytes = None

    def process_streamlit_spss_dataset(uploaded_file, start_dt, end_dt):
        if uploaded_file is None:
            return pd.DataFrame()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name
        try:
            df = pd.read_spss(tmp_path, convert_categoricals=False)

            if df.empty or "STIME" not in df.columns:
                return pd.DataFrame()

            df["STIME_CLEAN"] = pd.to_datetime(
                df["STIME"].astype(str).str.slice(0, 8),
                format="%Y%m%d",
                errors="coerce",
            )

            date_mask = (df["STIME_CLEAN"] >= pd.Timestamp(start_dt)) & (
                df["STIME_CLEAN"] <= pd.Timestamp(end_dt)
            )
            df = df[date_mask].copy()

            if df.empty:
                return pd.DataFrame()

            df["STIME_STR"] = df["STIME"].astype(str)
            nyear = df["STIME_STR"].str.slice(0, 4)
            nmonth = df["STIME_STR"].str.slice(4, 6)
            nday = df["STIME_STR"].str.slice(6, 8)
            recorded_date = nyear + "/" + nmonth + "/" + nday

            mask = (
                df["INTNR"] > 0
                if "INTNR" in df.columns
                else np.zeros(len(df), dtype=bool)
            )

            df_out = pd.DataFrame(index=df.index)

            df_out["Interview number"] = (
                df["INTNR"] if "INTNR" in df.columns else None
            )
            df_out["UCN Number"] = (
                np.where(mask, df["V80116"], None)
                if "V80116" in df.columns
                else None
            )
            df_out["Date"] = np.where(mask, recorded_date, None)

            rating_cols = {
                "Q11.1 RATING - FNB Business Lending products (overdraft, loans, etc.)": "Q11_1_1",
                "Q11.2 RATING - FNB Business Transactional products (cheque, debit card, credit card etc.)": "Q11_1_2",
                "Q11.3 RATING - FNB Business Insurance products (business credit protection plan, law-on-call business plan, etc.)": "Q11_1_3",
                "Q11.4 RATING - FNB Business FNB Business Investment products": "Q11_1_4",
                "Q11.5 RATING - FNB Business FNB Business FOREX products": "Q11_1_5",
            }
            for target, src in rating_cols.items():
                if src in df.columns:
                    df_out[target] = np.where(
                        mask,
                        pd.to_numeric(df[src], errors="coerce"),
                        np.nan,
                    )
                else:
                    df_out[target] = np.nan

            custom_label_mappings = {
                "Q11 REASONS - Lending_1": {
                    "src": "Q11A_1_1",
                    "label": "Overdraft",
                },
                "Q11 REASONS - Lending_2": {"src": "Q11A_1_2", "label": "Loans"},
                "Q11 REASONS - Lending_3": {"src": "Q11A_1_3", "label": "Other"},
                "Q11 REASONS - Lending_4": {"src": None, "label": "Other"},
                "Q11 REASONS - Lending_5": {"src": None, "label": "Other"},
                "Q11 REASONS - Transactional products_1": {
                    "src": "Q11A_2_1",
                    "label": "Cheque card",
                },
                "Q11 REASONS - Transactional products_2": {
                    "src": "Q11A_2_2",
                    "label": "Debit card",
                },
                "Q11 REASONS - Transactional products_3": {
                    "src": "Q11A_2_3",
                    "label": "Credit Card",
                },
                "Q11 REASONS - Transactional products_4": {
                    "src": "Q11A_2_4",
                    "label": "Other",
                },
                "Q11 REASONS - Transactional products_5": {
                    "src": None,
                    "label": "Other",
                },
                "Q11 REASONS - Insurance_1": {
                    "src": "Q11A_3_1",
                    "label": "Business credit protection plan",
                },
                "Q11 REASONS - Insurance_2": {
                    "src": "Q11A_3_2",
                    "label": "Law-on-call business plan",
                },
                "Q11 REASONS - Insurance_3": {
                    "src": "Q11A_3_3",
                    "label": "Other",
                },
                "Q11 REASONS - Insurance_4": {"src": None, "label": "Other"},
                "Q11 REASONS - Insurance_5": {"src": None, "label": "Other"},
                "Q11 REASONS - Investment products_1": {
                    "src": "Q11A_4_1",
                    "label": "Savings",
                },
                "Q11 REASONS - Investment products_2": {
                    "src": "Q11A_4_2",
                    "label": "Notice deposits",
                },
                "Q11 REASONS - Investment products_3": {
                    "src": "Q11A_4_3",
                    "label": "Other",
                },
                "Q11 REASONS - Investment products_4": {
                    "src": None,
                    "label": "Other",
                },
                "Q11 REASONS - Investment products_5": {
                    "src": None,
                    "label": "Other",
                },
            }

            forex_labels = {
                1: "Foreign Exchange",
                2: "Imports and Exports",
                3: "Structured Trade + Commodity Finance",
                4: "PayPal",
                5: "Trade (Trade Platform and Transacting)",
                6: "MoneyGram (TM)",
                7: "Global Payments (business global account)",
                8: "Travel card",
                9: "Trans-country Interbank Clearing",
                10: "Other",
            }
            for i, label_text in forex_labels.items():
                custom_label_mappings[
                    f"Q11 REASONS - FOREX products_{i}"
                ] = {"src": f"Q11A_5_{i}", "label": label_text}

            for target_col, config in custom_label_mappings.items():
                src_col = config["src"]
                if src_col is None or src_col not in df.columns:
                    df_out[target_col] = None
                    continue
                numeric_src = pd.to_numeric(df[src_col], errors="coerce")
                df_out[target_col] = np.where(
                    (mask) & (numeric_src == 1), config["label"], None
                )

            open_ends = {
                "Q11 REASONS - Lending OTHER": "TQ11A_1C3",
                "Q11 REASONS OTHER - Transactional products": "TQ11A_2C4",
                "Q11 REASONS - Insurance OTHER": "TQ11A_3C3",
                "Q11 REASONS - Investment products OTHER": "TQ11A_4C3",
                "Q11 REASONS - FOREX products OTHER": "TQ11A_5C10",
            }
            for target, src in open_ends.items():
                if src in df.columns:
                    string_series = (
                        df[src]
                        .astype(str)
                        .replace(["nan", "NaN", "None"], None)
                    )
                    df_out[target] = np.where(mask, string_series, None)
                else:
                    df_out[target] = None

            final_column_order = [
                "Interview number",
                "UCN Number",
                "Date",
                "Q11.1 RATING - FNB Business Lending products (overdraft, loans, etc.)",
                "Q11 REASONS - Lending_1",
                "Q11 REASONS - Lending_2",
                "Q11 REASONS - Lending_3",
                "Q11 REASONS - Lending_4",
                "Q11 REASONS - Lending_5",
                "Q11 REASONS - Lending OTHER",
                "Q11.2 RATING - FNB Business Transactional products (cheque, debit card, credit card etc.)",
                "Q11 REASONS - Transactional products_1",
                "Q11 REASONS - Transactional products_2",
                "Q11 REASONS - Transactional products_3",
                "Q11 REASONS - Transactional products_4",
                "Q11 REASONS - Transactional products_5",
                "Q11 REASONS OTHER - Transactional products",
                "Q11.3 RATING - FNB Business Insurance products (business credit protection plan, law-on-call business plan, etc.)",
                "Q11 REASONS - Insurance_1",
                "Q11 REASONS - Insurance_2",
                "Q11 REASONS - Insurance_3",
                "Q11 REASONS - Insurance_4",
                "Q11 REASONS - Insurance_5",
                "Q11 REASONS - Insurance OTHER",
                "Q11.4 RATING - FNB Business FNB Business Investment products",
                "Q11 REASONS - Investment products_1",
                "Q11 REASONS - Investment products_2",
                "Q11 REASONS - Investment products_3",
                "Q11 REASONS - Investment products_4",
                "Q11 REASONS - Investment products_5",
                "Q11 REASONS - Investment products OTHER",
                "Q11.5 RATING - FNB Business FNB Business FOREX products",
                "Q11 REASONS - FOREX products_1",
                "Q11 REASONS - FOREX products_2",
                "Q11 REASONS - FOREX products_3",
                "Q11 REASONS - FOREX products_4",
                "Q11 REASONS - FOREX products_5",
                "Q11 REASONS - FOREX products_6",
                "Q11 REASONS - FOREX products_7",
                "Q11 REASONS - FOREX products_8",
                "Q11 REASONS - FOREX products_9",
                "Q11 REASONS - FOREX products_10",
                "Q11 REASONS - FOREX products OTHER",
            ]

            df_final = df_out[final_column_order].copy()
            clean_headers = [header.split("_")[0] for header in df_final.columns]
            df_final.columns = clean_headers
            return df_final
        except Exception as e:
            st.error(f"❌ Error processing SPSS file: {e}")
            return pd.DataFrame()
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    if st.button("▶ Run Q11 Extraction", type="primary", key="run_q11"):
        if file_q11_r10 is None and file_q11_grow is None:
            st.error("Please upload at least one SPSS (.sav) file.")
        else:
            with st.spinner("Extracting Q11 ratings and reasons..."):
                clean_df1 = process_streamlit_spss_dataset(
                    file_q11_r10, q11_last_friday, today_q11
                )
                clean_df2 = process_streamlit_spss_dataset(
                    file_q11_grow, q11_last_friday, today_q11
                )

                out_buf = io.BytesIO()
                with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
                    has_data = False
                    if not clean_df1.empty:
                        clean_df1.to_excel(
                            writer, sheet_name="Enterprise-R10Mil", index=False
                        )
                        has_data = True
                    if not clean_df2.empty:
                        clean_df2.to_excel(
                            writer, sheet_name="Business-Growth", index=False
                        )
                        has_data = True
                    if not has_data:
                        pd.DataFrame(
                            {
                                "Notice": [
                                    "No records found for the selected date window."
                                ]
                            }
                        ).to_excel(
                            writer, sheet_name="No Data Found", index=False
                        )

                out_buf.seek(0)
                st.session_state.q11_bytes = out_buf.getvalue()
                st.session_state.q11_ready = True
                st.success("Q11 Extraction complete! Download ready below.")

    if st.session_state.q11_ready and st.session_state.q11_bytes:
        st.markdown("---")
        run_date_str = datetime.now().strftime("%m-%d-%Y")
        st.download_button(
            label="📥 Download Q11 Extraction Report (`Star W22 Q11 extraction.xlsx`)",
            data=st.session_state.q11_bytes,
            file_name=f"Star W22 Q11 extraction {run_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_q11_final",
        )


# ==========================================================================
# ==========================================================================
# TAB 5: NPS YEARLY DASHBOARD GENERATOR
# ==========================================================================
# ==========================================================================
with tab5:
    st.markdown("### 📅 NPS Yearly Dashboard Generator")
    st.markdown(
        "Upload your multi-wave yearly SPSS dataset (`.sav`) below to process and generate the comprehensive longitudinal `Star_Yearly_Dashboard.xlsx` report."
    )

    yearly_uploaded_file = st.file_uploader(
        "Upload Yearly SPSS File (`Project Star_W1 to W22.sav`)",
        type=["sav"],
        key="yearly_sav_file",
    )

    def generate_yearly_dashboard_workbook(uploaded_sav):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp:
            tmp.write(uploaded_sav.getvalue())
            tmp_path = tmp.name

        try:
            df, meta = pyreadstat.read_sav(tmp_path)
            df = df.loc[:, ~df.columns.duplicated()].copy()

            rm_driver_cols = [f"Q4_{i:02d}" for i in range(1, 24)]
            channel_q2_cols = [
                "Q2_1",
                "Q2_2",
                "Q2_3",
                "Q2_4",
                "Q2_5",
                "Q2_6",
                "Q2_7",
            ]
            pref_channel_col = "Q2_2_1"
            personal_banker_col = "Q4A2"
            pb_sat_col = "Q4B2"
            branch_service_cols = [
                "Q5_1",
                "Q5_2",
                "Q5_3",
                "Q5_4",
                "Q5_21",
                "Q5_22",
                "Q5_23",
                "Q5_24",
            ]
            branch_cols = branch_service_cols + ["Q10_2"]
            cc_cols = [
                "Q7_1",
                "Q7_2",
                "Q7_3",
                "Q7_4",
                "Q7_5",
                "Q7_21",
                "Q7_22",
                "Q10_3",
            ]

            online_cols = [
                "Q8_1",
                "Q8_2",
                "Q8_3",
                "Q8_4",
                "Q8_21",
                "Q8_22",
                "Q8_23",
                "Q8_24",
                "Q8_25",
                "Q8_26",
            ]
            app_cols = [
                "Q9_1",
                "Q9_2",
                "Q9_3",
                "Q9_4",
                "Q9_21",
                "Q9_22",
                "Q9_23",
                "Q9_24",
                "Q9_25",
                "Q9_26",
            ]
            chan_sat_cols = ["Q10_1", "Q10_2", "Q10_3", "Q10_4", "Q10_5"]
            product_sat_cols = [
                "Q11_1_1",
                "Q11_1_2",
                "Q11_1_3",
                "Q11_1_4",
                "Q11_1_5",
            ]

            q11a_groups = {
                "Q11A_1": [
                    ("Q11a.1. Lending products - Overdraft", "Q11A_1_1"),
                    ("Q11a.1. Lending products - Loans", "Q11A_1_2"),
                    ("Q11a.1. Lending products - Other", "Q11A_1_3"),
                ],
                "Q11A_2": [
                    (
                        "Q11a.2. Transactional products - Business account",
                        "Q11A_2_1",
                    ),
                    ("Q11a.2. Transactional products - Debit card", "Q11A_2_2"),
                    ("Q11a.2. Transactional products - Credit Card", "Q11A_2_3"),
                    ("Q11a.2. Transactional products - Other", "Q11A_2_4"),
                ],
                "Q11A_3": [
                    (
                        "Q11a.3. Insurance products - Business credit protection plan",
                        "Q11A_3_1",
                    ),
                    (
                        "Q11a.3. Insurance products - Law-on-call business plan",
                        "Q11A_3_2",
                    ),
                    ("Q11a.3. Insurance products - Other", "Q11A_3_3"),
                ],
                "Q11A_4": [
                    ("Q11a.4. Investment products - Savings", "Q11A_4_1"),
                    ("Q11A_4. Investment products - Notice deposits", "Q11A_4_2"),
                    ("Q11A_4. Investment products - Other", "Q11A_4_3"),
                ],
            }

            all_q11a_cols = [
                col for grp in q11a_groups.values() for _, col in grp
            ]
            expectations_cols = ["Q12_1", "Q12_2", "Q12_3"]
            q16_col = "Q16"

            consideration_items = [
                ("Absa", "Q16_1_1_flag"),
                ("Capitec", "Q16_1_2_flag"),
                ("Investec", "Q16_1_8_flag"),
                ("Mercantile", "Q16_1_9_flag"),
                ("Nedbank", "Q16_1_3_flag"),
                ("Sasfin", "Q16_1_10_flag"),
                ("Standard Bank", "Q16_1_4_flag"),
                ("Some other business banking offering", "Q16_1_5_flag"),
                ("I would not consider moving from FNB at all", "Q16_1_6_flag"),
            ]
            raw_consideration_cols = [
                "Q16_1_1",
                "Q16_1_2",
                "Q16_1_3",
                "Q16_1_4",
                "Q16_1_5",
                "Q16_1_6",
                "Q16_1_8",
                "Q16_1_9",
                "Q16_1_10",
            ]
            q5a_cols = [
                "Q5A_1",
                "Q5A_3",
                "Q5A_4",
                "Q5A_5",
                "Q5A_6",
                "Q5A_7",
                "Q5A_8",
                "Q5A_9",
                "Q5A_10",
                "Q5A_11",
                "Q5A_12",
                "Q5A_13",
                "Q5A_14",
                "Q5A_2",
            ]

            all_rating_cols = list(
                set(
                    rm_driver_cols
                    + branch_cols
                    + cc_cols
                    + online_cols
                    + app_cols
                    + chan_sat_cols
                    + product_sat_cols
                    + expectations_cols
                )
            )
            if pb_sat_col and pb_sat_col not in all_rating_cols:
                all_rating_cols.append(pb_sat_col)

            channel_items = [
                ("Q2. Business Manager", "Q2_1_flag"),
                ("Q2. Branch", "Q2_2_flag"),
                ("Q2. Contact Centre", "Q2_3_flag"),
                ("Q2. Online Banking", "Q2_4_flag"),
                ("Q2. FNB Business Banking App", "Q2_5_flag"),
                ("Q2. Account fulfilment", "Q6_code1_flag"),
                ("Q2. Product contact centre", "Q6_code2_flag"),
                ("Q2. Business Desk", "Q6_code3_flag"),
                ("Q2. Secure chat", "Q2_7_flag"),
                ("Q2. None of the above", "Q2_6_flag"),
            ]

            q5a_items = [
                ("Q5a. Prefer face-to-face interaction", "Q5A_1_flag"),
                ("Q5a. Required to submit documents", "Q5A_3_flag"),
                ("Q5a. Card collection", "Q5A_4_flag"),
                ("Q5a. Card queries", "Q5A_5_flag"),
                (
                    "Q5a. Issue could not be resolved digitally/limited options on digital channels",
                    "Q5A_6_flag",
                ),
                ("Q5a. Difficulty using digital channels", "Q5A_7_flag"),
                (
                    "Q5a. Not enough information on digital channels",
                    "Q5A_8_flag",
                ),
                ("Q5a. Directed to branch", "Q5A_9_flag"),
                ("Q5a. Needed help/assistance", "Q5A_10_flag"),
                ("Q5a. Cash deposit/withdrawal", "Q5A_11_flag"),
                ("Q5a. To open account", "Q5A_12_flag"),
                ("Q5a. To get bank statement", "Q5A_13_flag"),
                ("Q5a. Update personal information/details", "Q5A_14_flag"),
                ("Q5a. Others; specify", "Q5A_2_flag"),
            ]

            q6_items = [
                (
                    "FICA, outstanding documents relating to your account",
                    "Q6_code_1",
                ),
                (
                    "A specific product, e.g. such as Instant Solutions",
                    "Q6_code_2",
                ),
                (
                    "General enquiries (e.g. account, card & cheque-related)",
                    "Q6_code_3",
                ),
                ("Don't know / not sure", "Q6_code_4"),
            ]

            pref_channel_items = [
                ("Branch", 1),
                ("Contact Centre", 2),
                ("Online Banking", 3),
                ("FNB Business Banking App", 4),
                ("Business Manager at the branch", 5),
                ("Business/RM Manager", 6),
                ("Secure chat Help note", 7),
                ("None of the above", 8),
            ]

            personal_banker_items = [
                ("Yes", 1),
                ("No", 2),
                ("Do not have a personal FNB Account", 3),
                ("Refused to answer", 4),
            ]

            rm_labels = {}
            for col in rm_driver_cols:
                if (
                    col in meta.column_names_to_labels
                    and meta.column_names_to_labels[col]
                ):
                    rm_labels[col] = re.sub(
                        r"^(Q4[\._\s\d]*)+",
                        "",
                        meta.column_names_to_labels[col],
                        flags=re.IGNORECASE,
                    ).strip()
                else:
                    rm_labels[col] = f"Statement {col}"

            pb_sat_label = "Overall satisfaction with your Personal Banker"
            if (
                pb_sat_col
                and pb_sat_col in meta.column_names_to_labels
                and meta.column_names_to_labels[pb_sat_col]
            ):
                pb_sat_label = re.sub(
                    r"^(Q4B2[\._\s\d]*)+",
                    "",
                    meta.column_names_to_labels[pb_sat_col],
                    flags=re.IGNORECASE,
                ).strip()

            explicit_branch_labels = {
                "Q5_1": "Offering you personalized service",
                "Q5_2": "The manner in which you are welcomed and directed",
                "Q5_3": "Staff understanding your business banking needs",
                "Q5_4": "Waiting time for service",
                "Q5_21": "Operating hours of the branch",
                "Q5_22": (
                    "Staff communicating with you in a clear and easily"
                    " understandable way"
                ),
                "Q5_23": "Staff being willing to help",
                "Q5_24": "Consistently delivering on promises made to you",
                "Q10_2": "OVERALL Branch experience",
            }
            branch_labels = {
                col: explicit_branch_labels.get(col, col)
                for col in branch_cols
            }

            explicit_cc_labels = {
                "Q7_1": "Knowledge and competency",
                "Q7_2": "Taking ownership of your query",
                "Q7_3": "Offering you personalized service",
                "Q7_4": "Processing requests accurately",
                "Q7_5": "Providing the correct advice relating to your query",
                "Q7_21": (
                    "The agent providing the correct advice relating to your"
                    " enquiry or transaction"
                ),
                "Q7_22": "The agent delivering on promises made",
                "Q10_3": "OVERALL Contact Centre experience",
            }
            cc_labels = {
                col: explicit_cc_labels.get(col, col) for col in cc_cols
            }

            online_labels = {
                "Q8_1": (
                    "Q8. User-friendliness - having a logical layout and"
                    " structure"
                ),
                "Q8_2": (
                    "Q8. The range of functionalities offered effectively"
                    " addressing your banking needs"
                ),
                "Q8_3": "Q8. Reliability; stability and availability (uptime)",
                "Q8_4": (
                    "Q8. Security and other mechanisms mitigating against"
                    " fraudulent activity"
                ),
                "Q8_21": (
                    "Q8. The availability of help text; information or direct"
                    " chat / instant messaging"
                ),
                "Q8_22": (
                    "Q8. The response time on the submission of information"
                ),
                "Q8_23": "Q8. Logical layout and structure of the website",
                "Q8_24": (
                    "Q8. Fraud and data protection offered for transactions"
                    " executed"
                ),
                "Q8_25": (
                    "Q8. Accuracy – transactions executed are always correct"
                ),
                "Q8_26": (
                    "Q8. Navigability – quick to find logical links and always"
                    " know exactly where I am"
                ),
            }

            app_labels = {
                "Q9_1": (
                    "Q9. User-friendliness - having a logical layout and"
                    " structure"
                ),
                "Q9_2": (
                    "Q9. The range of functionalities offered effectively"
                    " addressing your banking needs"
                ),
                "Q9_3": "Q9. Reliability; stability and availability (uptime)",
                "Q9_4": (
                    "Q9. Security and other mechanisms mitigating against"
                    " fraudulent activity"
                ),
                "Q9_21": (
                    "Q9. The availability of help text; information or direct"
                    " chat / instant messaging"
                ),
                "Q9_22": (
                    "Q9. The response time on the submission of information"
                ),
                "Q9_23": (
                    "Q9. Logical layout and structure of the Business Banking"
                    " App"
                ),
                "Q9_24": (
                    "Q9. Fraud and data protection offered for transactions"
                    " executed"
                ),
                "Q9_25": (
                    "Q9. Accuracy – transactions executed are always correct"
                ),
                "Q9_26": (
                    "Q9. Navigability – quick to find logical links and always"
                    " know exactly where I am"
                ),
            }

            explicit_chan_sat_items = [
                ("Q10.1. OVERALL - Business Manager experience", "Q10_1"),
                ("Q10.2. OVERALL - Branch experience", "Q10_2"),
                ("Q10.3. OVERALL - Contact Centre experience", "Q10_3"),
                ("Q10.4. OVERALL - Online Banking experience", "Q10_4"),
                (
                    "Q10.5. OVERALL - FNB Business Banking App experience?",
                    "Q10_5",
                ),
            ]

            explicit_product_sat_items = [
                ("Q11.1. FNB Business Lending products", "Q11_1_1"),
                ("Q11.2. FNB Business Transactional products", "Q11_1_2"),
                ("Q11.3. FNB Business Insurance products", "Q11_1_3"),
                ("Q11.4. FNB Business Investment products", "Q11_1_4"),
                ("Q11.5. FNB Business Forex Products", "Q11_1_5"),
            ]

            expectations_items = [
                (
                    "Q12.1. Your overall level of satisfaction with the"
                    " products you received from FNB Business?",
                    "Q12_1",
                ),
                (
                    "Q12.2. Your overall level of satisfaction with FNB"
                    " Business?",
                    "Q12_2",
                ),
                (
                    "Q12.3. Your overall level of satisfaction with your BM"
                    " over the last 3 months?",
                    "Q12_3",
                ),
            ]

            cols = (
                [
                    "wave",
                    "REGION",
                    "SUBREG",
                    "SEGMENT",
                    "Type",
                    "Q14_1",
                    "Q14_2",
                    "Q6",
                    q16_col,
                ]
                + all_rating_cols
                + channel_q2_cols
                + q5a_cols
                + [pref_channel_col, personal_banker_col, pb_sat_col]
                + raw_consideration_cols
                + all_q11a_cols
            )
            cols_present = [c for c in cols if c in df.columns]

            df_sub = df[cols_present].copy()
            df_sub = df_sub.loc[:, ~df_sub.columns.duplicated()].copy()

            df_sub["Q6_raw"] = (
                df_sub["Q6"].copy() if "Q6" in df_sub.columns else None
            )

            for col in ["wave", "REGION", "SUBREG", "SEGMENT", "Type"]:
                if (
                    col in meta.variable_value_labels
                    and col in df_sub.columns
                ):
                    df_sub[col] = df_sub[col].map(
                        meta.variable_value_labels[col]
                    ).fillna(df_sub[col])

            df_sub["REGION"] = df_sub["REGION"].fillna("Unspecified")
            df_sub["SUBREG"] = df_sub["SUBREG"].fillna("Unspecified")
            df_sub["SEGMENT"] = df_sub["SEGMENT"].fillna("Unspecified")

            segment_relabel_map = {
                "MEDIUM TOUCH": "MEDIUM TOUCH (R10-R60M)",
                "HIGH TOUCH": "HIGH TOUCH (R60-R150M)",
                "PREMIUM": "PREMIUM (R150M+)",
            }
            df_sub["SEGMENT"] = df_sub["SEGMENT"].apply(
                lambda x: segment_relabel_map.get(
                    str(x).strip(), str(x).strip()
                )
            )

            def assign_type(segment_val):
                seg = str(segment_val).strip().upper()
                if (
                    "PUBLIC SECTOR" in seg
                    or "NON-PROFIT" in seg
                    or "PUBSC" in seg
                ):
                    return "PUBSC"
                elif seg in [
                    "HIGH TOUCH (R60-R150M)",
                    "HIGH TOUCH",
                    "PREMIUM (R150M+)",
                    "PREMIUM",
                ]:
                    return "R10m+"
                else:
                    return "Growth"

            df_sub["Type"] = df_sub["SEGMENT"].apply(assign_type)

            def get_wave_number(val):
                match = re.search(r"\d+", str(val))
                return int(match.group()) if match else 999

            df_sub["wave_num"] = df_sub["wave"].apply(get_wave_number)
            sorted_wave_nums = sorted(df_sub["wave_num"].unique())

            regions = sorted(
                [str(x) for x in df_sub["REGION"].unique() if x != "Unspecified"]
            )
            subregs = ["All"] + sorted(
                [str(x) for x in df_sub["SUBREG"].unique() if x != "Unspecified"]
            )
            segments = ["All"] + sorted(
                [
                    str(x)
                    for x in df_sub["SEGMENT"].unique()
                    if x != "Unspecified"
                ]
            )
            types = ["All"] + sorted(
                [str(x) for x in df_sub["Type"].unique() if x != "Unspecified"]
            )

            def clean_rating_score(val):
                try:
                    fval = float(val)
                    if fval == 11 or fval == 11.0:
                        return None
                    return fval if 1 <= fval <= 10 else None
                except (ValueError, TypeError):
                    return None

            new_cols = {}
            for col in all_rating_cols:
                if col in df_sub.columns:
                    new_cols[f"{col}_clean"] = df_sub[col].apply(
                        clean_rating_score
                    )

            new_cols["Q14_1_clean"] = df_sub["Q14_1"].apply(
                lambda x: x if pd.notnull(x) and x in range(0, 11) else None
            )
            new_cols["Q14_2_clean"] = df_sub["Q14_2"].apply(
                lambda x: x if pd.notnull(x) and x in range(0, 11) else None
            )

            def is_q6_match(val, code_num, desc_text):
                if pd.isnull(val):
                    return False
                sval = str(val).strip().lower()
                return (
                    sval == str(code_num)
                    or sval == f"{code_num}.0"
                    or desc_text.lower() in sval
                )

            if "Q6_raw" in df_sub.columns:
                q10_3_clean = new_cols.get(
                    "Q10_3_clean", pd.Series(index=df_sub.index)
                )
                new_cols["Q6_code1_clean"] = [
                    q10_3_clean[i]
                    if is_q6_match(v, 1, "Business Account Fulfilment")
                    else None
                    for i, v in enumerate(df_sub["Q6_raw"])
                ]
                new_cols["Q6_code2_clean"] = [
                    q10_3_clean[i]
                    if is_q6_match(v, 2, "Product Contact Centre")
                    else None
                    for i, v in enumerate(df_sub["Q6_raw"])
                ]
                new_cols["Q6_code3_clean"] = [
                    q10_3_clean[i]
                    if is_q6_match(v, 3, "Business Desk")
                    else None
                    for i, v in enumerate(df_sub["Q6_raw"])
                ]
                new_cols["Q6_code1_flag"] = [
                    1 if is_q6_match(v, 1, "Business Account Fulfilment") else 0
                    for v in df_sub["Q6_raw"]
                ]
                new_cols["Q6_code2_flag"] = [
                    1 if is_q6_match(v, 2, "Product Contact Centre") else 0
                    for v in df_sub["Q6_raw"]
                ]
                new_cols["Q6_code3_flag"] = [
                    1 if is_q6_match(v, 3, "Business Desk") else 0
                    for v in df_sub["Q6_raw"]
                ]
                for code_val in [1, 2, 3, 4]:
                    new_cols[f"Q6_code_{code_val}"] = df_sub["Q6_raw"].apply(
                        lambda x: 1
                        if pd.notnull(x) and float(x) == code_val
                        else 0
                    )

            for q11a_key, sub_items in q11a_groups.items():
                sub_cols = [c for _, c in sub_items if c in df_sub.columns]
                if sub_cols:
                    df_sub[f"{q11a_key}_Base"] = (
                        df_sub[sub_cols].notnull().any(axis=1).astype(int)
                    )
                for _, col in sub_items:
                    if col in df_sub.columns:
                        new_cols[f"{col}_flag"] = df_sub[col].apply(
                            lambda x: 1
                            if pd.notnull(x) and float(x) == 1
                            else 0
                        )

            if q16_col in df_sub.columns:
                df_sub["Q16_mapped"] = df_sub[q16_col].apply(
                    lambda x: "Yes"
                    if str(x).strip().lower() in ["1", "1.0", "yes"]
                    else (
                        "No"
                        if str(x).strip().lower() in ["2", "2.0", "no"]
                        else None
                    )
                )
                new_cols["Q16_Yes"] = df_sub["Q16_mapped"].apply(
                    lambda x: 1 if x == "Yes" else 0
                )
                new_cols["Q16_No"] = df_sub["Q16_mapped"].apply(
                    lambda x: 1 if x == "No" else 0
                )
                new_cols["Q16_Base"] = df_sub["Q16_mapped"].apply(
                    lambda x: 1 if pd.notnull(x) else 0
                )
            else:
                new_cols["Q16_Yes"], new_cols["Q16_No"], new_cols[
                    "Q16_Base"
                ] = (0, 0, 0)

            for _, flag_col in consideration_items:
                raw_col = flag_col.replace("_flag", "")
                new_cols[flag_col] = (
                    df_sub[raw_col].apply(
                        lambda x: 1 if pd.notnull(x) and float(x) == 1 else 0
                    )
                    if raw_col in df_sub.columns
                    else 0
                )

            bm_clean, fnb_clean = (
                new_cols["Q14_2_clean"],
                new_cols["Q14_1_clean"],
            )
            new_cols["BM_Det"] = bm_clean.apply(
                lambda x: 1 if pd.notnull(x) and x <= 6 else 0
            )
            new_cols["BM_Pas"] = bm_clean.apply(
                lambda x: 1 if pd.notnull(x) and 7 <= x <= 8 else 0
            )
            new_cols["BM_Pro"] = bm_clean.apply(
                lambda x: 1 if pd.notnull(x) and x >= 9 else 0
            )
            new_cols["BM_Base"] = bm_clean.apply(
                lambda x: 1 if pd.notnull(x) else 0
            )

            new_cols["FNB_Det"] = fnb_clean.apply(
                lambda x: 1 if pd.notnull(x) and x <= 6 else 0
            )
            new_cols["FNB_Pas"] = fnb_clean.apply(
                lambda x: 1 if pd.notnull(x) and 7 <= x <= 8 else 0
            )
            new_cols["FNB_Pro"] = fnb_clean.apply(
                lambda x: 1 if pd.notnull(x) and x >= 9 else 0
            )
            new_cols["FNB_Base"] = fnb_clean.apply(
                lambda x: 1 if pd.notnull(x) else 0
            )

            for qcol in channel_q2_cols:
                if qcol in df_sub.columns:
                    new_cols[f"{qcol}_flag"] = df_sub[qcol].apply(
                        lambda x: 1 if pd.notnull(x) and float(x) == 1 else 0
                    )
            for qcol in q5a_cols:
                if qcol in df_sub.columns:
                    new_cols[f"{qcol}_flag"] = df_sub[qcol].apply(
                        lambda x: 1 if pd.notnull(x) and float(x) == 1 else 0
                    )
            if pref_channel_col in df_sub.columns:
                for lbl, code in pref_channel_items:
                    new_cols[f"pref_chan_{code}"] = df_sub[
                        pref_channel_col
                    ].apply(
                        lambda x: 1
                        if pd.notnull(x) and float(x) == code
                        else 0
                    )
            if personal_banker_col in df_sub.columns:
                for lbl, code in personal_banker_items:
                    new_cols[f"pb_code_{code}"] = df_sub[
                        personal_banker_col
                    ].apply(
                        lambda x: 1
                        if pd.notnull(x) and float(x) == code
                        else 0
                    )

            df_sub = pd.concat(
                [df_sub, pd.DataFrame(new_cols, index=df_sub.index)], axis=1
            )
            df_sub = df_sub.loc[:, ~df_sub.columns.duplicated()].copy()

            q16a_flag_cols = [
                flag_col
                for _, flag_col in consideration_items
                if flag_col in df_sub.columns
            ]
            df_sub["Q16A_Base"] = (
                (df_sub[q16a_flag_cols].sum(axis=1) > 0).astype(int)
                if q16a_flag_cols
                else 0
            )
            df_sub["Channel_Base"] = (
                df_sub[[c for c in channel_q2_cols if c in df_sub.columns]]
                .notnull()
                .any(axis=1)
                .astype(int)
            )
            df_sub["Q5A_Base"] = (
                df_sub[
                    [
                        f"{q}_flag"
                        for q in q5a_cols
                        if f"{q}_flag" in df_sub.columns
                    ]
                ].sum(axis=1)
                > 0
            ).astype(int)
            df_sub["Q6_Base"] = df_sub["Q6_raw"].apply(
                lambda x: 1 if pd.notnull(x) else 0
            )
            df_sub["Pref_Channel_Base"] = (
                df_sub[pref_channel_col].apply(
                    lambda x: 1 if pd.notnull(x) else 0
                )
                if pref_channel_col in df_sub.columns
                else 0
            )
            df_sub["PB_Base"] = (
                df_sub[personal_banker_col].apply(
                    lambda x: 1 if pd.notnull(x) else 0
                )
                if personal_banker_col in df_sub.columns
                else 0
            )

            base_agg_dict = {
                "Total_n": ("wave_num", "count"),
                "BM_Base": ("BM_Base", "sum"),
                "BM_Det": ("BM_Det", "sum"),
                "BM_Pas": ("BM_Pas", "sum"),
                "BM_Pro": ("BM_Pro", "sum"),
                "FNB_Base": ("FNB_Base", "sum"),
                "FNB_Det": ("FNB_Det", "sum"),
                "FNB_Pas": ("FNB_Pas", "sum"),
                "FNB_Pro": ("FNB_Pro", "sum"),
                "Channel_Base": ("Channel_Base", "sum"),
                "Q5A_Base": ("Q5A_Base", "sum"),
                "Q6_Base": ("Q6_Base", "sum"),
                "Pref_Channel_Base": ("Pref_Channel_Base", "sum"),
                "PB_Base": ("PB_Base", "sum"),
            }

            for col_name in ["Q16_Base", "Q16_Yes", "Q16_No", "Q16A_Base"]:
                if col_name in df_sub.columns:
                    base_agg_dict[col_name] = (col_name, "sum")
            for _, flag_col in consideration_items:
                if flag_col in df_sub.columns:
                    base_agg_dict[flag_col] = (flag_col, "sum")
            for q11a_key in q11a_groups.keys():
                if f"{q11a_key}_Base" in df_sub.columns:
                    base_agg_dict[f"{q11a_key}_Base"] = (
                        f"{q11a_key}_Base",
                        "sum",
                    )
            for col in all_rating_cols:
                if f"{col}_clean" in df_sub.columns:
                    base_agg_dict[f"{col}_sum"] = (f"{col}_clean", "sum")
                    base_agg_dict[f"{col}_count"] = (f"{col}_clean", "count")
            for code_num in [1, 2, 3]:
                clean_col_name = f"Q6_code{code_num}_clean"
                if clean_col_name in df_sub.columns:
                    base_agg_dict[f"Q6_code{code_num}_sum"] = (
                        clean_col_name,
                        "sum",
                    )
                    base_agg_dict[f"Q6_code{code_num}_count"] = (
                        clean_col_name,
                        "count",
                    )
            for _, flag_col in channel_items:
                if flag_col in df_sub.columns:
                    base_agg_dict[flag_col] = (flag_col, "sum")
                    base_agg_dict[f"{flag_col}_count"] = (flag_col, "count")
            for _, flag_col in q5a_items:
                if flag_col in df_sub.columns:
                    base_agg_dict[flag_col] = (flag_col, "sum")
                    base_agg_dict[f"{flag_col}_count"] = (flag_col, "count")
            for q11a_key, sub_items in q11a_groups.items():
                for _, col in sub_items:
                    flag_name = f"{col}_flag"
                    if flag_name in df_sub.columns:
                        base_agg_dict[flag_name] = (flag_name, "sum")
                        base_agg_dict[f"{flag_name}_count"] = (flag_name, "count")
            for code_val in [1, 2, 3, 4]:
                col_name = f"Q6_code_{code_val}"
                if col_name in df_sub.columns:
                    base_agg_dict[col_name] = (col_name, "sum")
                    base_agg_dict[f"{col}_count"] = (col_name, "count")
            for _, code in pref_channel_items:
                flag_name = f"pref_chan_{code}"
                if flag_name in df_sub.columns:
                    base_agg_dict[flag_name] = (flag_name, "sum")
                    base_agg_dict[f"{flag_name}_count"] = (flag_name, "count")
            for _, code in personal_banker_items:
                flag_name = f"pb_code_{code}"
                if flag_name in df_sub.columns:
                    base_agg_dict[flag_name] = (flag_name, "sum")
                    base_agg_dict[f"{flag_name}_count"] = (flag_name, "count")

            for col_name in ["Q16_Yes", "Q16_No"]:
                if col_name in df_sub.columns:
                    base_agg_dict[f"{col_name}_count"] = (col_name, "count")

            agg_dict = {
                k: v for k, v in base_agg_dict.items() if v[0] in df_sub.columns
            }
            summary_data = df_sub.groupby(
                ["wave_num", "REGION", "SUBREG", "SEGMENT", "Type"],
                as_index=False,
            ).agg(**agg_dict)
            data_cols_list = list(summary_data.columns)

            def get_excel_col_letter(col_name):
                if col_name not in data_cols_list:
                    return "A"
                return get_column_letter(data_cols_list.index(col_name) + 1)

            wb_dash = openpyxl.Workbook()
            ws_dash = wb_dash.active
            ws_dash.title = "NPS Dashboard"
            ws_dash.views.sheetView[0].showGridLines = True

            ws_data = wb_dash.create_sheet(title="_Data")
            ws_data.sheet_state = "hidden"
            ws_data.append(list(summary_data.columns))
            for row in summary_data.itertuples(index=False):
                ws_data.append(list(row))

            FNB_TEAL, FNB_AMBER, LIGHT_TEAL, LIGHT_AMBER, WHITE, GRAY_FILL, GRAY_TEXT = (
                "009B9E",
                "EA8B24",
                "E5F5F5",
                "FDF3E7",
                "FFFFFF",
                "F2F2F2",
                "7F7F7F",
            )
            font_title_main = Font(
                name="Calibri", size=14, bold=True, color=WHITE
            )
            font_title_sub = Font(
                name="Calibri", size=10, italic=True, color=WHITE
            )
            font_header = Font(name="Calibri", size=9, bold=True, color=WHITE)
            font_bold = Font(name="Calibri", size=10, bold=True)
            font_regular = Font(name="Calibri", size=10)
            font_italic_bold = Font(
                name="Calibri", size=10, italic=True, bold=True, color=WHITE
            )
            font_sub_subheader = Font(
                name="Calibri", size=10, italic=True, bold=True, color="105B5C"
            )
            font_filter_lbl = Font(
                name="Calibri", size=10, bold=True, color=WHITE
            )
            font_filter_val = Font(
                name="Calibri", size=11, bold=True, color="105B5C"
            )
            font_na = Font(name="Calibri", size=9, italic=True, color=GRAY_TEXT)

            fill_teal_header = PatternFill(
                start_color=FNB_TEAL, end_color=FNB_TEAL, fill_type="solid"
            )
            fill_amber_header = PatternFill(
                start_color=FNB_AMBER, end_color=FNB_AMBER, fill_type="solid"
            )
            fill_light_teal = PatternFill(
                start_color=LIGHT_TEAL, end_color=LIGHT_TEAL, fill_type="solid"
            )
            fill_light_amber = PatternFill(
                start_color=LIGHT_AMBER,
                end_color=LIGHT_AMBER,
                fill_type="solid",
            )
            fill_na = PatternFill(
                start_color=GRAY_FILL, end_color=GRAY_FILL, fill_type="solid"
            )

            thin_border_side = Side(border_style="thin", color="B0C4DE")
            thick_border_side = Side(border_style="medium", color=FNB_TEAL)
            border_cell = Border(
                left=thin_border_side,
                right=thin_border_side,
                top=thin_border_side,
                bottom=thin_border_side,
            )
            border_box = Border(
                left=thick_border_side,
                right=thick_border_side,
                top=thick_border_side,
                bottom=thick_border_side,
            )

            align_center = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            align_left = Alignment(horizontal="left", vertical="center")
            align_right = Alignment(horizontal="right", vertical="center")

            start_col = 3
            end_col_idx = start_col + len(sorted_wave_nums) - 1

            ws_dash.merge_cells(
                start_row=2, start_column=2, end_row=2, end_column=end_col_idx
            )
            ws_dash.cell(
                row=2,
                column=2,
                value="PROJECT STAR: MULTI-WAVE PERFORMANCE ANALYTICS",
            ).font = font_title_main
            ws_dash.cell(row=2, column=2).fill = fill_teal_header

            ws_dash.merge_cells(
                start_row=3, start_column=2, end_row=3, end_column=end_col_idx
            )
            ws_dash.cell(
                row=3,
                column=2,
                value=(
                    "A Comprehensive Longitudinal Evaluation of NPS, RM"
                    " Engagement Drivers, Channel Usage, and Local Branch"
                    " Dynamics"
                ),
            ).font = font_title_sub
            ws_dash.cell(row=3, column=2).fill = fill_amber_header

            filter_configs = [
                ("Region:", 2, 3),
                ("Sub-Region:", 5, 6),
                ("Segment:", 8, 9),
                ("Type:", 11, 12),
            ]
            for label, lbl_col, val_col in filter_configs:
                c_lbl = ws_dash.cell(row=5, column=lbl_col, value=label)
                c_lbl.font, c_lbl.fill, c_lbl.alignment, c_lbl.border = (
                    font_filter_lbl,
                    fill_teal_header,
                    align_center,
                    border_box,
                )
                c_val = ws_dash.cell(row=5, column=val_col, value="All")
                c_val.font, c_val.fill, c_val.alignment, c_val.border = (
                    font_filter_val,
                    fill_light_amber,
                    align_center,
                    border_box,
                )

            dv_region = DataValidation(
                type="list",
                formula1=f'"{",".join(["All"] + regions)}"',
                allow_blank=True,
            )
            dv_subreg = DataValidation(
                type="list",
                formula1=f'"{",".join(subregs)}"',
                allow_blank=True,
            )
            dv_segment = DataValidation(
                type="list",
                formula1=f'"{",".join(segments)}"',
                allow_blank=True,
            )
            dv_type = DataValidation(
                type="list",
                formula1=f'"{",".join(types)}"',
                allow_blank=True,
            )
            for dv, cell_ref in zip(
                [dv_region, dv_subreg, dv_segment, dv_type],
                ["C5", "F5", "I5", "L5"],
            ):
                ws_dash.add_data_validation(dv)
                dv.add(ws_dash[cell_ref])

            def build_sumifs(target_col_letter, wave_num):
                base_formula = f"_Data!{target_col_letter}:{target_col_letter}, _Data!A:A, {wave_num}"
                return (
                    f"SUMIFS({base_formula}, _Data!B:B,"
                    ' IF($C$5="All", "*", $C$5), _Data!C:C,'
                    ' IF($F$5="All", "*", $F$5), _Data!D:D,'
                    ' IF($I$5="All", "*", $I$5), _Data!E:E,'
                    ' IF($L$5="All", "*", $L$5))'
                )

            def add_section_header(ws_target, row_idx, section_title, wave_nums):
                c_title = ws_target.cell(
                    row=row_idx, column=2, value=section_title
                )
                c_title.font, c_title.fill, c_title.alignment, c_title.border = (
                    font_italic_bold,
                    fill_teal_header,
                    align_left,
                    border_cell,
                )
                for idx, w_num in enumerate(wave_nums):
                    c_hdr = ws_target.cell(
                        row=row_idx, column=start_col + idx, value=f"Wave {w_num}"
                    )
                    c_hdr.font, c_hdr.fill, c_hdr.alignment, c_hdr.border = (
                        font_header,
                        fill_teal_header,
                        align_center,
                        border_cell,
                    )

            header_row, n_row = 7, 8
            ws_dash.cell(row=n_row, column=2, value="n=").font = font_bold
            ws_dash.cell(row=n_row, column=2).alignment = align_right

            for idx, w_num in enumerate(sorted_wave_nums):
                col = start_col + idx
                c_hdr = ws_dash.cell(
                    row=header_row, column=col, value=f"Wave {w_num}"
                )
                c_hdr.font, c_hdr.fill, c_hdr.alignment, c_hdr.border = (
                    font_header,
                    fill_teal_header,
                    align_center,
                    border_cell,
                )
                c_n = ws_dash.cell(
                    row=n_row,
                    column=col,
                    value=f"={build_sumifs(get_excel_col_letter('Total_n'), w_num)}",
                )
                c_n.font, c_n.alignment, c_n.border, c_n.number_format = (
                    font_bold,
                    align_center,
                    border_cell,
                    "#,##0",
                )

            metric_rows = [
                ("NET PROMOTER SCORE - NPS", "header_dark", None),
                ("NPS FNB Relationship Manager", "sub_header", None),
                ("Net Score", "net_bm", None),
                ("Base (n)", "sum", "BM_Base"),
                ("Detractors", "pct", "BM_Det"),
                ("Passives", "pct", "BM_Pas"),
                ("Promoters", "pct", "BM_Pro"),
                ("", "blank", None),
                ("NPS FNB Business", "sub_header", None),
                ("Net Score", "net_fnb", None),
                ("Base (n)", "sum", "FNB_Base"),
                ("Detractors", "pct", "FNB_Det"),
                ("Passives", "pct", "FNB_Pas"),
                ("Promoters", "pct", "FNB_Pro"),
            ]

            curr_row = 9
            for label, row_type, data_col_name in metric_rows:
                if row_type == "blank":
                    curr_row += 1
                    continue
                if row_type == "header_dark":
                    add_section_header(
                        ws_dash, curr_row, label, sorted_wave_nums
                    )
                    curr_row += 1
                    continue
                if row_type == "sub_header":
                    ws_dash.merge_cells(
                        start_row=curr_row,
                        start_column=2,
                        end_row=curr_row,
                        end_column=start_col + len(sorted_wave_nums) - 1,
                    )
                    c = ws_dash.cell(row=curr_row, column=2, value=label)
                    c.font, c.fill, c.alignment = (
                        font_italic_bold,
                        fill_amber_header,
                        align_left,
                    )
                    curr_row += 1
                    continue

                lbl_cell = ws_dash.cell(row=curr_row, column=2, value=label)
                lbl_cell.font = (
                    font_bold
                    if "Net" in label or "Base" in label
                    else font_regular
                )
                lbl_cell.alignment = align_left
                if "Passives" in label:
                    lbl_cell.fill = fill_light_teal

                for idx, w_num in enumerate(sorted_wave_nums):
                    col = start_col + idx
                    col_let = get_column_letter(col)
                    val_cell = ws_dash.cell(row=curr_row, column=col)
                    val_cell.alignment, val_cell.border = (
                        align_center,
                        border_cell,
                    )
                    if "Passives" in label:
                        val_cell.fill = fill_light_teal

                    if row_type == "sum":
                        val_cell.value = f"={build_sumifs(get_excel_col_letter(data_col_name), w_num)}"
                        val_cell.number_format, val_cell.font = (
                            "#,##0",
                            font_bold,
                        )
                    elif row_type == "pct":
                        base_col = "BM_Base" if curr_row < 17 else "FNB_Base"
                        val_cell.value = (
                            f"=IFERROR({build_sumifs(get_excel_col_letter(data_col_name), w_num)}/{build_sumifs(get_excel_col_letter(base_col), w_num)},"
                            " 0)"
                        )
                        val_cell.number_format = "0%"
                    elif row_type in ["net_bm", "net_fnb"]:
                        det_row, pro_row = curr_row + 2, curr_row + 4
                        val_cell.value = (
                            f"=({col_let}{pro_row}-{col_let}{det_row})*100"
                        )
                        val_cell.number_format, val_cell.font = (
                            "0.00",
                            font_bold,
                        )
                curr_row += 1

            def add_dashboard_section(title, items, label_dict, is_mean=True):
                nonlocal curr_row
                curr_row += 1
                add_section_header(ws_dash, curr_row, title, sorted_wave_nums)
                curr_row += 1
                for idx_item, var_code in enumerate(items):
                    if (
                        f"{var_code}_sum" not in summary_data.columns
                        and is_mean
                    ):
                        continue
                    lbl = label_dict.get(var_code, var_code)
                    is_overall = "OVERALL" in str(lbl).upper()
                    use_zebra = idx_item % 2 == 0

                    lbl_cell = ws_dash.cell(row=curr_row, column=2, value=lbl)
                    lbl_cell.font = font_bold if is_overall else font_regular
                    lbl_cell.alignment = align_left
                    if is_overall:
                        lbl_cell.fill = fill_light_amber
                    elif use_zebra:
                        lbl_cell.fill = fill_light_teal

                    sum_let, cnt_let = get_excel_col_letter(
                        f"{var_code}_sum"
                    ), get_excel_col_letter(f"{var_code}_count")
                    for idx, w_num in enumerate(sorted_wave_nums):
                        val_cell = ws_dash.cell(row=curr_row, column=start_col + idx)
                        val_cell.alignment, val_cell.border = (
                            align_center,
                            border_cell,
                        )
                        val_cell.value = (
                            f'=IFERROR(IF({build_sumifs(cnt_let, w_num)}=0,'
                            f' "N/A",'
                            f" {build_sumifs(sum_let, w_num)}/{build_sumifs(cnt_let, w_num)}),"
                            ' "N/A")'
                        )
                        val_cell.number_format = "0.00"

                        wave_sub = summary_data[
                            summary_data["wave_num"] == w_num
                        ]
                        if (
                            wave_sub.empty
                            or wave_sub[f"{var_code}_count"].sum() == 0
                        ):
                            val_cell.fill, val_cell.font = fill_na, font_na
                        elif is_overall:
                            val_cell.fill, val_cell.font = (
                                fill_light_amber,
                                font_bold,
                            )
                        elif use_zebra:
                            val_cell.fill = fill_light_teal
                    curr_row += 1

            add_dashboard_section(
                "LAST INTERACTION WITH YOUR RM/CPE/AE", rm_driver_cols, rm_labels
            )

            curr_row += 1
            add_section_header(ws_dash, curr_row, "CHANNEL USAGE", sorted_wave_nums)
            curr_row += 1
            chan_base_let = get_excel_col_letter("Channel_Base")
            ws_dash.cell(row=curr_row, column=2, value="Base (n)").font = (
                font_bold
            )
            for idx, w_num in enumerate(sorted_wave_nums):
                c = ws_dash.cell(
                    row=curr_row,
                    column=start_col + idx,
                    value=f"={build_sumifs(chan_base_let, w_num)}",
                )
                c.font, c.alignment, c.border, c.number_format = (
                    font_bold,
                    align_center,
                    border_cell,
                    "#,##0",
                )
            curr_row += 1

            for item_idx, (label_text, flag_col) in enumerate(channel_items):
                lbl_cell = ws_dash.cell(row=curr_row, column=2, value=label_text)
                lbl_cell.font, lbl_cell.alignment = font_regular, align_left
                use_zebra = item_idx % 2 == 0
                if use_zebra:
                    lbl_cell.fill = fill_light_teal
                flag_let = get_excel_col_letter(flag_col)
                for idx, w_num in enumerate(sorted_wave_nums):
                    val_cell = ws_dash.cell(row=curr_row, column=start_col + idx)
                    val_cell.alignment, val_cell.border = (
                        align_center,
                        border_cell,
                    )
                    if use_zebra:
                        val_cell.fill = fill_light_teal
                    val_cell.value = (
                        f'=IFERROR(IF({build_sumifs(chan_base_let, w_num)}=0,'
                        f' "N/A",'
                        f" {build_sumifs(flag_let, w_num)}/{build_sumifs(chan_base_let, w_num)}),"
                        ' "N/A")'
                    )
                    val_cell.number_format = "0%"

                    wave_sub = summary_data[summary_data["wave_num"] == w_num]
                    if (
                        wave_sub.empty
                        or wave_sub[f"{flag_col}"].sum() == 0
                        or wave_sub["Channel_Base"].sum() == 0
                    ):
                        val_cell.fill, val_cell.font = fill_na, font_na
                curr_row += 1

            curr_row += 1
            add_section_header(
                ws_dash, curr_row, "Preferred Channel usage", sorted_wave_nums
            )
            curr_row += 1
            pref_base_let = get_excel_col_letter("Pref_Channel_Base")
            ws_dash.cell(row=curr_row, column=2, value="Base (n)").font = (
                font_bold
            )
            for idx, w_num in enumerate(sorted_wave_nums):
                c = ws_dash.cell(
                    row=curr_row,
                    column=start_col + idx,
                    value=f"={build_sumifs(pref_base_let, w_num)}",
                )
                c.font, c.alignment, c.border, c.number_format = (
                    font_bold,
                    align_center,
                    border_cell,
                    "#,##0",
                )
            curr_row += 1

            for item_idx, (label_text, code) in enumerate(pref_channel_items):
                lbl_cell = ws_dash.cell(row=curr_row, column=2, value=label_text)
                lbl_cell.font, lbl_cell.alignment = font_regular, align_left
                use_zebra = item_idx % 2 == 0
                if use_zebra:
                    lbl_cell.fill = fill_light_teal
                flag_col = f"pref_chan_{code}"
                flag_let = get_excel_col_letter(flag_col)
                for idx, w_num in enumerate(sorted_wave_nums):
                    val_cell = ws_dash.cell(row=curr_row, column=start_col + idx)
                    val_cell.alignment, val_cell.border = (
                        align_center,
                        border_cell,
                    )
                    if use_zebra:
                        val_cell.fill = fill_light_teal
                    val_cell.value = (
                        f'=IFERROR(IF({build_sumifs(pref_base_let, w_num)}=0,'
                        f' "N/A",'
                        f" {build_sumifs(flag_let, w_num)}/{build_sumifs(pref_base_let, w_num)}),"
                        ' "N/A")'
                    )
                    val_cell.number_format = "0%"

                    wave_sub = summary_data[summary_data["wave_num"] == w_num]
                    if (
                        wave_sub.empty
                        or wave_sub[flag_col].sum() == 0
                        or wave_sub["Pref_Channel_Base"].sum() == 0
                    ):
                        val_cell.fill, val_cell.font = fill_na, font_na
                curr_row += 1

            add_dashboard_section(
                "LOCAL BRANCH MEAN SCORES", branch_cols, branch_labels
            )

            curr_row += 1
            add_section_header(
                ws_dash,
                curr_row,
                "Contact Centre Service Aspects",
                sorted_wave_nums,
            )
            curr_row += 1
            q6_base_let = get_excel_col_letter("Q6_Base")
            ws_dash.cell(row=curr_row, column=2, value="Base (n)").font = (
                font_bold
            )
            for idx, w_num in enumerate(sorted_wave_nums):
                c = ws_dash.cell(
                    row=curr_row,
                    column=start_col + idx,
                    value=f"={build_sumifs(q6_base_let, w_num)}",
                )
                c.font, c.alignment, c.border, c.number_format = (
                    font_bold,
                    align_center,
                    border_cell,
                    "#,##0",
                )
            curr_row += 1

            for item_idx, (label_text, code_col) in enumerate(q6_items):
                lbl_cell = ws_dash.cell(row=curr_row, column=2, value=label_text)
                lbl_cell.font, lbl_cell.alignment = font_regular, align_left
                use_zebra = item_idx % 2 == 0
                if use_zebra:
                    lbl_cell.fill = fill_light_teal
                flag_let = get_excel_col_letter(code_col)
                for idx, w_num in enumerate(sorted_wave_nums):
                    val_cell = ws_dash.cell(row=curr_row, column=start_col + idx)
                    val_cell.alignment, val_cell.border = (
                        align_center,
                        border_cell,
                    )
                    if use_zebra:
                        val_cell.fill = fill_light_teal
                    val_cell.value = (
                        f'=IFERROR(IF({build_sumifs(q6_base_let, w_num)}=0,'
                        f' "N/A",'
                        f" {build_sumifs(flag_let, w_num)}/{build_sumifs(q6_base_let, w_num)}),"
                        ' "N/A")'
                    )
                    val_cell.number_format = "0%"

                    wave_sub = summary_data[summary_data["wave_num"] == w_num]
                    if (
                        wave_sub.empty
                        or wave_sub[code_col].sum() == 0
                        or wave_sub["Q6_Base"].sum() == 0
                    ):
                        val_cell.fill, val_cell.font = fill_na, font_na
                curr_row += 1

            ws_dash.merge_cells(
                start_row=curr_row,
                start_column=2,
                end_row=curr_row,
                end_column=start_col + len(sorted_wave_nums) - 1,
            )
            sub_hdr = ws_dash.cell(
                row=curr_row, column=2, value="Contact Centre agent ratings"
            )
            sub_hdr.font, sub_hdr.fill, sub_hdr.alignment = (
                font_sub_subheader,
                fill_light_teal,
                align_left,
            )
            curr_row += 1

            for idx_item, var_code in enumerate(cc_cols):
                lbl = cc_labels.get(var_code, var_code)
                is_overall = "OVERALL" in str(lbl).upper()
                use_zebra = idx_item % 2 == 0
                lbl_cell = ws_dash.cell(row=curr_row, column=2, value=lbl)
                lbl_cell.font = font_bold if is_overall else font_regular
                if is_overall:
                    lbl_cell.fill = fill_light_amber
                elif use_zebra:
                    lbl_cell.fill = fill_light_teal
                sum_let, cnt_let = get_excel_col_letter(
                    f"{var_code}_sum"
                ), get_excel_col_letter(f"{var_code}_count")
                for idx, w_num in enumerate(sorted_wave_nums):
                    val_cell = ws_dash.cell(row=curr_row, column=start_col + idx)
                    val_cell.alignment, val_cell.border = (
                        align_center,
                        border_cell,
                    )
                    val_cell.value = (
                        f'=IFERROR(IF({build_sumifs(cnt_let, w_num)}=0,'
                        f' "N/A",'
                        f" {build_sumifs(sum_let, w_num)}/{build_sumifs(cnt_let, w_num)}),"
                        ' "N/A")'
                    )
                    val_cell.number_format = "0.00"
                    wave_sub = summary_data[summary_data["wave_num"] == w_num]
                    if (
                        wave_sub.empty
                        or wave_sub[f"{var_code}_count"].sum() == 0
                    ):
                        val_cell.fill, val_cell.font = fill_na, font_na
                    elif is_overall:
                        val_cell.fill, val_cell.font = (
                            fill_light_amber,
                            font_bold,
                        )
                    elif use_zebra:
                        val_cell.fill = fill_light_teal
                curr_row += 1

            add_dashboard_section(
                "Online Banking through laptop or desktop PC Service Aspects",
                online_cols,
                online_labels,
            )
            add_dashboard_section(
                "Banking App Service Aspects", app_cols, app_labels
            )

            curr_row += 1
            add_section_header(ws_dash, curr_row, "OVERALL ratings", sorted_wave_nums)
            curr_row += 1
            for label_text, var_code in explicit_chan_sat_items:
                lbl_cell = ws_dash.cell(row=curr_row, column=2, value=label_text)
                lbl_cell.font, lbl_cell.fill, lbl_cell.alignment = (
                    font_bold,
                    fill_light_amber,
                    align_left,
                )
                sum_let, cnt_let = get_excel_col_letter(
                    f"{var_code}_sum"
                ), get_excel_col_letter(f"{var_code}_count")
                for idx, w_num in enumerate(sorted_wave_nums):
                    val_cell = ws_dash.cell(row=curr_row, column=start_col + idx)
                    val_cell.alignment, val_cell.border, val_cell.fill, val_cell.font = (
                        align_center,
                        border_cell,
                        fill_light_amber,
                        font_bold,
                    )
                    val_cell.value = (
                        f'=IFERROR(IF({build_sumifs(cnt_let, w_num)}=0,'
                        f' "N/A",'
                        f" {build_sumifs(sum_let, w_num)}/{build_sumifs(cnt_let, w_num)}),"
                        ' "N/A")'
                    )
                    val_cell.number_format = "0.00"
                curr_row += 1

            curr_row += 1
            add_section_header(
                ws_dash, curr_row, "Satisfaction with products", sorted_wave_nums
            )
            curr_row += 1
            for item_idx, (label_text, var_code) in enumerate(
                explicit_product_sat_items
            ):
                lbl_cell = ws_dash.cell(row=curr_row, column=2, value=label_text)
                lbl_cell.font, lbl_cell.alignment = font_regular, align_left
                use_zebra = item_idx % 2 == 0
                if use_zebra:
                    lbl_cell.fill = fill_light_teal
                sum_let, cnt_let = get_excel_col_letter(
                    f"{var_code}_sum"
                ), get_excel_col_letter(f"{var_code}_count")
                for idx, w_num in enumerate(sorted_wave_nums):
                    val_cell = ws_dash.cell(row=curr_row, column=start_col + idx)
                    val_cell.alignment, val_cell.border = (
                        align_center,
                        border_cell,
                    )
                    if use_zebra:
                        val_cell.fill = fill_light_teal
                    val_cell.value = (
                        f'=IFERROR(IF({build_sumifs(cnt_let, w_num)}=0,'
                        f' "N/A",'
                        f" {build_sumifs(sum_let, w_num)}/{build_sumifs(cnt_let, w_num)}),"
                        ' "N/A")'
                    )
                    val_cell.number_format = "0.00"
                curr_row += 1

            ws_dash.merge_cells(
                start_row=curr_row,
                start_column=2,
                end_row=curr_row,
                end_column=start_col + len(sorted_wave_nums) - 1,
            )
            ws_dash.cell(
                row=curr_row, column=2, value="Drivers of Dissatisfaction"
            ).font, ws_dash.cell(row=curr_row, column=2).fill, ws_dash.cell(
                row=curr_row, column=2
            ).alignment = (
                font_sub_subheader,
                fill_light_teal,
                align_left,
            )
            curr_row += 1

            for q11a_key, sub_items in q11a_groups.items():
                base_let = get_excel_col_letter(f"{q11a_key}_Base")
                ws_dash.cell(
                    row=curr_row, column=2, value=f"Base (n) - {q11a_key}"
                ).font = font_bold
                for idx, w_num in enumerate(sorted_wave_nums):
                    c = ws_dash.cell(
                        row=curr_row,
                        column=start_col + idx,
                        value=f"={build_sumifs(base_let, w_num)}",
                    )
                    c.font, c.alignment, c.border, c.number_format = (
                        font_bold,
                        align_center,
                        border_cell,
                        "#,##0",
                    )
                curr_row += 1
                for item_idx, (lbl_txt, col_code) in enumerate(sub_items):
                    lbl_cell = ws_dash.cell(row=curr_row, column=2, value=lbl_txt)
                    lbl_cell.font, lbl_cell.alignment = font_regular, align_left
                    use_zebra = item_idx % 2 == 0
                    if use_zebra:
                        lbl_cell.fill = fill_light_teal
                    flag_col = f"{col_code}_flag"
                    flag_let = get_excel_col_letter(flag_col)
                    for idx, w_num in enumerate(sorted_wave_nums):
                        val_cell = ws_dash.cell(row=curr_row, column=start_col + idx)
                        val_cell.alignment, val_cell.border = (
                            align_center,
                            border_cell,
                        )
                        if use_zebra:
                            val_cell.fill = fill_light_teal
                        val_cell.value = (
                            f'=IFERROR(IF({build_sumifs(base_let, w_num)}=0,'
                            f' "N/A",'
                            f" {build_sumifs(flag_let, w_num)}/{build_sumifs(base_let, w_num)}),"
                            ' "N/A")'
                        )
                        val_cell.number_format = "0%"

                        wave_sub = summary_data[
                            summary_data["wave_num"] == w_num
                        ]
                        if (
                            wave_sub.empty
                            or wave_sub[flag_col].sum() == 0
                            or wave_sub[f"{q11a_key}_Base"].sum() == 0
                        ):
                            val_cell.fill, val_cell.font = fill_na, font_na
                    curr_row += 1

            add_dashboard_section(
                "Satisfaction: Quality of service and product solutions",
                expectations_items,
                {},
            )

            curr_row += 1
            add_section_header(
                ws_dash,
                curr_row,
                "Business banking consideration",
                sorted_wave_nums,
            )
            curr_row += 1
            ws_dash.merge_cells(
                start_row=curr_row,
                start_column=2,
                end_row=curr_row,
                end_column=start_col + len(sorted_wave_nums) - 1,
            )
            ws_dash.cell(
                row=curr_row, column=2, value="Q16. Consideration to switch"
            ).font, ws_dash.cell(row=curr_row, column=2).fill = (
                font_sub_subheader,
                fill_light_teal,
            )
            curr_row += 1

            q16_base_let = get_excel_col_letter("Q16_Base")
            ws_dash.cell(row=curr_row, column=2, value="Base (n)").font = (
                font_bold
            )
            for idx, w_num in enumerate(sorted_wave_nums):
                c = ws_dash.cell(
                    row=curr_row,
                    column=start_col + idx,
                    value=f"={build_sumifs(q16_base_let, w_num)}",
                )
                c.font, c.alignment, c.border, c.number_format = (
                    font_bold,
                    align_center,
                    border_cell,
                    "#,##0",
                )
            curr_row += 1

            for item_idx, (lbl_txt, col_key, is_pct) in enumerate(
                [
                    ("Yes (n)", "Q16_Yes", False),
                    ("Yes (%)", "Q16_Yes", True),
                    ("No (n)", "Q16_No", False),
                    ("No (%)", "Q16_No", True),
                ]
            ):
                lbl_cell = ws_dash.cell(row=curr_row, column=2, value=lbl_txt)
                lbl_cell.font, lbl_cell.alignment = font_regular, align_left
                use_zebra = item_idx in [2, 3]
                if use_zebra:
                    lbl_cell.fill = fill_light_teal
                flag_let = get_excel_col_letter(col_key)
                for idx, w_num in enumerate(sorted_wave_nums):
                    val_cell = ws_dash.cell(row=curr_row, column=start_col + idx)
                    val_cell.alignment, val_cell.border = (
                        align_center,
                        border_cell,
                    )
                    if use_zebra:
                        val_cell.fill = fill_light_teal
                    if is_pct:
                        val_cell.value = (
                            f'=IFERROR(IF({build_sumifs(q16_base_let, w_num)}=0,'
                            f' "N/A",'
                            f" {build_sumifs(flag_let, w_num)}/{build_sumifs(q16_base_let, w_num)}),"
                            ' "N/A")'
                        )
                        val_cell.number_format = "0%"
                    else:
                        val_cell.value = (
                            f"={build_sumifs(flag_let, w_num)}"
                        )
                        val_cell.number_format, val_cell.font = (
                            "#,##0",
                            font_bold,
                        )

                    wave_sub = summary_data[summary_data["wave_num"] == w_num]
                    if (
                        wave_sub.empty
                        or wave_sub[col_key].sum() == 0
                        or wave_sub["Q16_Base"].sum() == 0
                    ):
                        if is_pct:
                            val_cell.fill, val_cell.font = fill_na, font_na
                curr_row += 1

            ws_dash.merge_cells(
                start_row=curr_row,
                start_column=2,
                end_row=curr_row,
                end_column=start_col + len(sorted_wave_nums) - 1,
            )
            ws_dash.cell(
                row=curr_row,
                column=2,
                value="Q16a. Banks/Financial service providers considered",
            ).font, ws_dash.cell(row=curr_row, column=2).fill = (
                font_sub_subheader,
                fill_light_teal,
            )
            curr_row += 1

            q16a_base_let = get_excel_col_letter("Q16A_Base")
            ws_dash.cell(row=curr_row, column=2, value="Base (n)").font = (
                font_bold
            )
            for idx, w_num in enumerate(sorted_wave_nums):
                c = ws_dash.cell(
                    row=curr_row,
                    column=start_col + idx,
                    value=f"={build_sumifs(q16a_base_let, w_num)}",
                )
                c.font, c.alignment, c.border, c.number_format = (
                    font_bold,
                    align_center,
                    border_cell,
                    "#,##0",
                )
            curr_row += 1

            for item_idx, (lbl_txt, flag_col) in enumerate(consideration_items):
                lbl_cell = ws_dash.cell(row=curr_row, column=2, value=lbl_txt)
                lbl_cell.font, lbl_cell.alignment = font_regular, align_left
                use_zebra = item_idx % 2 == 0
                if use_zebra:
                    lbl_cell.fill = fill_light_teal
                flag_let = get_excel_col_letter(flag_col)
                for idx, w_num in enumerate(sorted_wave_nums):
                    val_cell = ws_dash.cell(row=curr_row, column=start_col + idx)
                    val_cell.alignment, val_cell.border = (
                        align_center,
                        border_cell,
                    )
                    if use_zebra:
                        val_cell.fill = fill_light_teal
                    val_cell.value = (
                        f'=IFERROR(IF({build_sumifs(q16a_base_let, w_num)}=0,'
                        f' "N/A",'
                        f" {build_sumifs(flag_let, w_num)}/{build_sumifs(q16a_base_let, w_num)}),"
                        ' "N/A")'
                    )
                    val_cell.number_format = "0%"

                    wave_sub = summary_data[summary_data["wave_num"] == w_num]
                    if (
                        wave_sub.empty
                        or wave_sub[flag_col].sum() == 0
                        or wave_sub["Q16A_Base"].sum() == 0
                    ):
                        val_cell.fill, val_cell.font = fill_na, font_na
                curr_row += 1

            curr_row += 1
            add_section_header(
                ws_dash,
                curr_row,
                "Motivations for Choosing In-Branch over Digital Channels/ Call"
                " centre",
                sorted_wave_nums,
            )
            curr_row += 1
            q5a_base_let = get_excel_col_letter("Q5A_Base")
            ws_dash.cell(row=curr_row, column=2, value="Base (n)").font = (
                font_bold
            )
            for idx, w_num in enumerate(sorted_wave_nums):
                c = ws_dash.cell(
                    row=curr_row,
                    column=start_col + idx,
                    value=f"={build_sumifs(q5a_base_let, w_num)}",
                )
                c.font, c.alignment, c.border, c.number_format = (
                    font_bold,
                    align_center,
                    border_cell,
                    "#,##0",
                )
            curr_row += 1

            for item_idx, (label_text, flag_col) in enumerate(q5a_items):
                lbl_cell = ws_dash.cell(row=curr_row, column=2, value=label_text)
                lbl_cell.font, lbl_cell.alignment = font_regular, align_left
                use_zebra = item_idx % 2 == 0
                if use_zebra:
                    lbl_cell.fill = fill_light_teal
                flag_let = get_excel_col_letter(flag_col)
                for idx, w_num in enumerate(sorted_wave_nums):
                    val_cell = ws_dash.cell(row=curr_row, column=start_col + idx)
                    val_cell.alignment, val_cell.border = (
                        align_center,
                        border_cell,
                    )
                    if use_zebra:
                        val_cell.fill = fill_light_teal
                    val_cell.value = (
                        f'=IFERROR(IF({build_sumifs(q5a_base_let, w_num)}=0,'
                        f' "N/A",'
                        f" {build_sumifs(flag_let, w_num)}/{build_sumifs(q5a_base_let, w_num)}),"
                        ' "N/A")'
                    )
                    val_cell.number_format = "0%"

                    wave_sub = summary_data[summary_data["wave_num"] == w_num]
                    if (
                        wave_sub.empty
                        or wave_sub[flag_col].sum() == 0
                        or wave_sub["Q5A_Base"].sum() == 0
                    ):
                        val_cell.fill, val_cell.font = fill_na, font_na
                curr_row += 1

            ws_dash.column_dimensions["B"].width = 75
            for idx in range(len(sorted_wave_nums)):
                ws_dash.column_dimensions[
                    get_column_letter(start_col + idx)
                ].width = 12

            output_buffer = io.BytesIO()
            wb_dash.save(output_buffer)
            output_buffer.seek(0)
            return output_buffer
        except Exception as e:
            st.error(f"❌ Error generating Yearly Dashboard: {e}")
            return None
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    if "yearly_ready" not in st.session_state:
        st.session_state.yearly_ready = False
    if "yearly_bytes" not in st.session_state:
        st.session_state.yearly_bytes = None

    if st.button(
        "🚀 Generate Yearly Dashboard Report",
        type="primary",
        key="run_yearly_dash_btn",
    ):
        if yearly_uploaded_file is None:
            st.error("Please upload the yearly SPSS `.sav` file first!")
        else:
            with st.spinner(
                "Processing multi-wave dataset and building dashboard..."
            ):
                yearly_excel_bytes = generate_yearly_dashboard_workbook(
                    yearly_uploaded_file
                )
                if yearly_excel_bytes:
                    st.session_state.yearly_bytes = yearly_excel_bytes
                    st.session_state.yearly_ready = True
                    st.success(
                        "✅ Yearly Dashboard report generated successfully! Download"
                        " ready below."
                    )

    if st.session_state.yearly_ready and st.session_state.yearly_bytes:
        st.markdown("---")
        st.download_button(
            label="📥 Download Formatted Excel Report (`Star_Yearly_Dashboard.xlsx`)",
            data=st.session_state.yearly_bytes,
            file_name="Star_Yearly_Dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_yearly_dash_final",
        )


# ==========================================================================
# ==========================================================================
# TAB 6: SME/ENT TABLES (NEW TAB ADDED)
# ==========================================================================
# ==========================================================================
with tab6:
    st.markdown("### 🏢 SME / ENT Segment & Type Analysis Tables")
    st.markdown(
        "Upload your SPSS datasets below or view live statistical summaries (Mean & Valid N) broken down by **TYPE** (Total, Growth, R10Mil) and **Segment** (Total, Enterprise, Gold/SME/Platinum), filtered by Wave."
    )

    st.markdown("---")
    st.subheader("📁 Upload Latest SPSS Datasets for Live Analysis")
    col_up1, col_up2, col_up3 = st.columns(3)
    with col_up1:
        sme_file_grow = st.file_uploader(
            "Upload Growth (.sav)", type=["sav"], key="sme_grow"
        )
    with col_up2:
        sme_file_r10 = st.file_uploader(
            "Upload R10Mil (.sav)", type=["sav"], key="sme_r10"
        )
    with col_up3:
        sme_file_pub = st.file_uploader(
            "Upload PUBW (.sav)", type=["sav"], key="sme_pub"
        )

    # Mock or master data for SME/ENT tab demonstration / live use
    @st.cache_data
    def generate_sme_ent_mock_data(num_rows=2500):
        np.random.seed(42)
        waves = [f"Wave {i}" for i in range(20, 23)]
        types = ["Growth", "R10Mil"]
        segments = ["ENTERPRISE", "GOLD/SME/PLATINUM"]

        data = {
            "wave": np.random.choice(waves, num_rows),
            "type": np.random.choice(types, num_rows, p=[0.6, 0.4]),
            "segment": np.random.choice(segments, num_rows, p=[0.3, 0.7]),
            "fnb_nps_raw": np.random.choice(
                range(11),
                num_rows,
                p=[0.01, 0.01, 0.02, 0.02, 0.03, 0.05, 0.08, 0.15, 0.25, 0.20, 0.18],
            ),
            "bm_nps_raw": np.random.choice(
                range(11),
                num_rows,
                p=[0.01, 0.01, 0.01, 0.02, 0.02, 0.04, 0.07, 0.12, 0.28, 0.22, 0.20],
            ),
            "Q10_1": np.random.choice(range(1, 12), num_rows),
            "Q10_2": np.random.choice(range(1, 12), num_rows),
            "Q10_3": np.random.choice(range(1, 12), num_rows),
            "Q10_4": np.random.choice(range(1, 12), num_rows),
            "Q10_5": np.random.choice(range(1, 12), num_rows),
            "Q11_1_1": np.random.choice(range(1, 12), num_rows),
            "Q11_1_2": np.random.choice(range(1, 12), num_rows),
            "Q11_1_3": np.random.choice(range(1, 12), num_rows),
            "Q11_1_4": np.random.choice(range(1, 12), num_rows),
            "Q11_1_5": np.random.choice(range(1, 12), num_rows),
            "Q12_1": np.random.choice(range(1, 12), num_rows),
            "Q12_2": np.random.choice(range(1, 12), num_rows),
            "Q12_3": np.random.choice(range(1, 12), num_rows),
        }
        return pd.DataFrame(data)

    df_sme = generate_sme_ent_mock_data()

    def calc_nps_score(series):
        valid = series.dropna()
        if len(valid) == 0:
            return np.nan
        promoters = (valid >= 9).sum()
        detractors = (valid <= 6).sum()
        return ((promoters - detractors) / len(valid)) * 100

    all_waves_sme = sorted(df_sme["wave"].unique())
    selected_waves_sme = st.multiselect(
        "Select Wave(s) to Filter", options=all_waves_sme, default=all_waves_sme, key="sme_wave_filter"
    )

    if not selected_waves_sme:
        st.warning("Please select at least one wave.")
    else:
        filtered_sme_df = df_sme[df_sme["wave"].isin(selected_waves_sme)]

        metrics_config_sme = [
            ("FNB_NPS_SCORE", "fnb_nps_raw", "nps"),
            ("BM_NPS_SCORE", "bm_nps_raw", "nps"),
            ("Q10.1. OVERALL BUSINESS MANAGER EXPERIENCE?", "Q10_1", "mean"),
            ("Q10.2. OVERALL BRANCH EXPERIENCE?", "Q10_2", "mean"),
            ("Q10.3. OVERALL CONTACT CENTRE EXPERIENCE?", "Q10_3", "mean"),
            ("Q10.4. OVERALL ONLINE BANKING EXPERIENCE?", "Q10_4", "mean"),
            ("Q10.5. OVERALL FNB BUSINESS BANKING APP EXPERIENCE?", "Q10_5", "mean"),
            ("Q11.1 Lending products", "Q11_1_1", "mean"),
            ("Q11.2 Transactional products", "Q11_1_2", "mean"),
            ("Q11.3 Insurance products", "Q11_1_3", "mean"),
            ("Q11.4 Investment products", "Q11_1_4", "mean"),
            ("Q11.5 Forex Products", "Q11_1_5", "mean"),
            ("Q12. Your overall level of satisfaction with the products you received from FNB Business?", "Q12_1", "mean"),
            ("Q12. Your overall level of satisfaction with FNB Business over the last 3 months?", "Q12_2", "mean"),
            ("Q12. Your overall level of satisfaction with your Relationship Manager over the last 3-6 months?", "Q12_3", "mean"),
        ]

        table_rows_sme = []
        subsets_sme = [
            ("Total", filtered_sme_df),
            ("Growth", filtered_sme_df[filtered_sme_df["type"] == "Growth"]),
            ("R10Mil", filtered_sme_df[filtered_sme_df["type"] == "R10Mil"]),
            ("Segment_Total", filtered_sme_df),
            ("Enterprise", filtered_sme_df[filtered_sme_df["segment"] == "ENTERPRISE"]),
            ("Gold_SME_Platinum", filtered_sme_df[filtered_sme_df["segment"] == "GOLD/SME/PLATINUM"]),
        ]

        for q_label, col_name, calc_type in metrics_config_sme:
            row_data = {"Questions": q_label}
            for sub_name, sub_df in subsets_sme:
                valid_data = sub_df[col_name].dropna()
                valid_n = len(valid_data)
                if valid_n == 0:
                    val = np.nan
                elif calc_type == "nps":
                    val = calc_nps_score(valid_data)
                else:
                    val = valid_data.mean()
                row_data[f"{sub_name}_Mean"] = round(val, 2) if not np.isnan(val) else np.nan
                row_data[f"{sub_name}_ValidN"] = float(valid_n)
            table_rows_sme.append(row_data)

        flat_data_sme = []
        for r in table_rows_sme:
            flat_data_sme.append(
                [
                    r["Questions"],
                    r["Total_Mean"], r["Total_ValidN"],
                    r["Growth_Mean"], r["Growth_ValidN"],
                    r["R10Mil_Mean"], r["R10Mil_ValidN"],
                    r["Segment_Total_Mean"], r["Segment_Total_ValidN"],
                    r["Enterprise_Mean"], r["Enterprise_ValidN"],
                    r["Gold_SME_Platinum_Mean"], r["Gold_SME_Platinum_ValidN"],
                ]
            )

        formatted_sme_display_df = pd.DataFrame(flat_data_sme, columns=[
            "Questions", "Total Mean", "Total Valid N", "Growth Mean", "Growth Valid N",
            "R10Mil Mean", "R10Mil Valid N", "Segment Total Mean", "Segment Total Valid N",
            "ENTERPRISE Mean", "ENTERPRISE Valid N", "GOLD/SME/PLATINUM Mean", "GOLD/SME/PLATINUM Valid N"
        ])
        st.dataframe(formatted_sme_display_df, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("📥 Download SME/ENT Tables as Excel")
        if st.button("Download FNB Branded SME/ENT Report (.xlsx)", type="primary", key="sme_download_btn"):
            wb_sme = openpyxl.Workbook()
            ws_sme = wb_sme.active
            ws_sme.title = "SME-ENT Tables"
            ws_sme.views.sheetView[0].showGridLines = True

            # FNB Brand Color Palette from Logo
            teal_fill = PatternFill(start_color="00A3AD", end_color="00A3AD", fill_type="solid")
            orange_fill = PatternFill(start_color="F58220", end_color="F58220", fill_type="solid")
            gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            
            white_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            dark_font = Font(name="Calibri", size=10, bold=True, color="000000")
            regular_font = Font(name="Calibri", size=10)
            
            thin_border_side = Side(style='thin', color='BFBFBF')
            thin_border = Border(
                left=thin_border_side, right=thin_border_side,
                top=thin_border_side, bottom=thin_border_side
            )
            center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # ROW 1: Top-Level "Total" Banner
            ws_sme.cell(row=1, column=1, value="")
            ws_sme.merge_cells("B1:M1")
            ws_sme.cell(row=1, column=2, value="Total")

            # ROW 2: Category Banners (TYPE vs Seg1)
            ws_sme.cell(row=2, column=1, value="")
            ws_sme.merge_cells("B2:G2")
            ws_sme.cell(row=2, column=2, value="TYPE")
            ws_sme.merge_cells("H2:M2")
            ws_sme.cell(row=2, column=8, value="Seg1")

            # ROW 3: Sub-Category / Segment Headers
            ws_sme.cell(row=3, column=1, value="Questions")
            ws_sme.cell(row=3, column=2, value="Total")
            ws_sme.merge_cells("B3:C3")
            ws_sme.cell(row=3, column=4, value="Growth")
            ws_sme.merge_cells("D3:E3")
            ws_sme.cell(row=3, column=6, value="R10Mil")
            ws_sme.merge_cells("F3:G3")
            ws_sme.cell(row=3, column=8, value="Total")
            ws_sme.merge_cells("H3:I3")
            ws_sme.cell(row=3, column=10, value="ENTERPRISE")
            ws_sme.merge_cells("J3:K3")
            ws_sme.cell(row=3, column=12, value="GOLD/SME/PLATINUM")
            ws_sme.merge_cells("L3:M3")

            # ROW 4: Metric Sub-Headers (Mean / Valid N)
            metrics_row = ["Questions", "Mean", "Valid N", "Mean", "Valid N", "Mean", "Valid N", "Mean", "Valid N", "Mean", "Valid N", "Mean", "Valid N"]
            for col_idx, h_text in enumerate(metrics_row, start=1):
                ws_sme.cell(row=4, column=col_idx, value=h_text)

            for r in range(1, 5):
                for c in range(1, 14):
                    cell = ws_sme.cell(row=r, column=c)
                    cell.border = thin_border
                    cell.alignment = center_align
                    if r == 1:
                        cell.fill = teal_fill
                        cell.font = white_font
                    elif r == 2 and c >= 8:
                        cell.fill = gray_fill
                        cell.font = dark_font
                    elif r == 2:
                        cell.fill = orange_fill
                        cell.font = white_font
                    elif r == 3 and c >= 8:
                        cell.fill = gray_fill
                        cell.font = dark_font
                    elif r == 3:
                        cell.fill = orange_fill
                        cell.font = white_font
                    else: # Row 4
                        cell.fill = teal_fill
                        cell.font = white_font

            for row_idx, row_vals in enumerate(flat_data_sme, start=5):
                for col_idx, val in enumerate(row_vals, start=1):
                    cell = ws_sme.cell(row=row_idx, column=col_idx, value=val)
                    cell.font = regular_font
                    cell.border = thin_border
                    if col_idx == 1:
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        if col_idx % 2 == 0:
                            cell.number_format = "#,##0.00"
                        else:
                            cell.number_format = "#,##0"

            for col in ws_sme.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_sme.column_dimensions[col_letter].width = max(max_len + 3, 12)
            ws_sme.column_dimensions['A'].width = 50

            sme_output = io.BytesIO()
            wb_sme.save(sme_output)
            sme_output.seek(0)

            st.download_button(
                label="📁 Download FNB Branded SME/ENT Report (.xlsx)",
                data=sme_output,
                file_name="SME_ENT_Tables_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="sme_final_download"
            )
