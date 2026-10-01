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
# TAB 1: PROJECT STATUS & QUOTAS UPDATE
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
        col_p1, _ = st.columns([1, 3])
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

    def generate_exact_pm_update_workbook():
        output_buffer = io.BytesIO()
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        FNB_TEAL = PatternFill(start_color="00A3AD", end_color="00A3AD", fill_type="solid")
        FNB_ORANGE = PatternFill(start_color="F58220", end_color="F58220", fill_type="solid")
        LIGHT_TEAL = PatternFill(start_color="D9F2F4", end_color="D9F2F4", fill_type="solid")
        LIGHT_ORANGE = PatternFill(start_color="FDF3EC", end_color="FDF3EC", fill_type="solid")
        GRAY_HEADER = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        
        WHITE_BOLD_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        THIN_BORDER = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'), top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))
        CENTER_ALIGN = Alignment(horizontal="center", vertical="center")

        ws_sum = wb.create_sheet(title='Summary')
        ws_sum.append(["", "Segment", "TOTAL Target", "TOTAL Achieved", "Total Outstanding"])
        for col_let in ['B', 'C', 'D', 'E']:
            cell = ws_sum[f'{col_let}1']
            cell.fill = FNB_TEAL
            cell.font = WHITE_BOLD_FONT
            cell.alignment = CENTER_ALIGN
        for _, row in summary_df.iterrows():
            ws_sum.append(["", row["Segment"], row["TOTAL Target"], row["TOTAL Achieved"], row["Total Outstanding"]])

        # Business Sheet
        ws_bus = wb.create_sheet(title='Update Business')
        ws_bus.cell(row=1, column=2, value="Region").fill = GRAY_HEADER
        ws_bus.merge_cells("B1:H1")
        ws_bus.cell(row=2, column=2, value="Business").fill = FNB_TEAL
        ws_bus.cell(row=2, column=2).font = WHITE_BOLD_FONT
        for c_idx, reg in enumerate(["Cape", "Gauteng North", "Gauteng South Central", "Inland", "KwaZulu-Natal", "Total"], start=3):
            cell = ws_bus.cell(row=2, column=c_idx, value=reg)
            cell.fill = FNB_TEAL if c_idx < 8 else FNB_ORANGE
            cell.font = WHITE_BOLD_FONT

        bus_rows = [
            ("Eastern Cape", [180, 0, 0, 0, 0]), ("Free State", [0, 0, 0, 54, 0]),
            ("Gauteng East", [0, 0, 88, 0, 0]), ("Gauteng South Central", [0, 0, 93, 0, 0]),
            ("Gauteng Tshwane East", [0, 65, 0, 0, 0]), ("Gauteng Tshwane North", [0, 63, 0, 0, 0]),
            ("Gauteng Tshwane South", [0, 0, 0, 0, 0]), ("Gauteng West-Rand", [0, 0, 67, 0, 0]),
            ("Greater Sandton", [0, 115, 0, 0, 0]), ("Gauteng Midrand", [0, 59, 0, 0, 0]),
            ("KZN North", [0, 0, 0, 0, 42]), ("KZN South", [0, 0, 0, 0, 45]),
            ("KZN West", [0, 0, 0, 0, 50]), ("Limpopo", [0, 0, 0, 66, 0]),
            ("Mpumalanga", [0, 0, 0, 65, 0]), ("North West", [0, 0, 0, 72, 0]),
            ("Northern Cape", [0, 0, 0, 33, 0]), ("Western Cape", [105, 0, 0, 0, 0])
        ]
        for idx, (prov, vals) in enumerate(bus_rows, start=3):
            ws_bus.cell(row=idx, column=2, value=prov).border = THIN_BORDER
            for v_idx, val in enumerate(vals, start=3):
                c = ws_bus.cell(row=idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            tot_c = ws_bus.cell(row=idx, column=8, value=f"=SUM(C{idx}:G{idx})")
            tot_c.border = THIN_BORDER
            tot_c.fill = LIGHT_ORANGE

        tot_row_idx = len(bus_rows) + 3
        ws_bus.cell(row=tot_row_idx, column=2, value="TOTAL INLC R10-R60MIL").fill = FNB_TEAL
        ws_bus.cell(row=tot_row_idx, column=2).font = WHITE_BOLD_FONT
        for c_idx in range(3, 8):
            col_let = openpyxl.utils.get_column_letter(c_idx)
            c = ws_bus.cell(row=tot_row_idx, column=c_idx, value=f"=SUM({col_let}3:{col_let}{tot_row_idx-1})")
            c.fill = FNB_TEAL
            c.font = WHITE_BOLD_FONT
            c.border = THIN_BORDER
        tot_sum = ws_bus.cell(row=tot_row_idx, column=8, value=f"=SUM(H3:H{tot_row_idx-1})")
        tot_sum.fill = FNB_ORANGE
        tot_sum.font = WHITE_BOLD_FONT

        # Business Right Side Table
        ws_bus.cell(row=1, column=11, value="Region").fill = GRAY_HEADER
        ws_bus.merge_cells("K1:R1")
        ws_bus.cell(row=2, column=10, value="Business").fill = FNB_TEAL
        ws_bus.cell(row=2, column=10).font = WHITE_BOLD_FONT
        for c_idx, sc in enumerate(["Cape", "Gauteng North", "Gauteng South Central", "Inland", "KwaZulu-Natal", "Total", "Quota", "Outstanding"], start=11):
            cell = ws_bus.cell(row=2, column=c_idx, value=sc)
            cell.fill = FNB_TEAL if c_idx < 16 else (FNB_ORANGE if c_idx == 17 else RED_FILL)
            cell.font = WHITE_BOLD_FONT

        for idx, (s_name, s_vals, quota_val) in enumerate([
            ("R0M-R1M", [159, 127, 116, 157, 57], q_b1),
            ("R1M-R5M", [81, 90, 60, 97, 43], q_b2),
            ("R5M-R10M", [45, 85, 72, 36, 37], q_b3),
            ("R10-R60M", [56, 105, 107, 109, 82], q_b4)
        ], start=3):
            ws_bus.cell(row=idx, column=10, value=s_name).border = THIN_BORDER
            for v_idx, val in enumerate(s_vals, start=11):
                c = ws_bus.cell(row=idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            tot_seg = ws_bus.cell(row=idx, column=16, value=f"=SUM(K{idx}:O{idx})")
            tot_seg.border = THIN_BORDER
            tot_seg.fill = LIGHT_ORANGE
            
            q_c = ws_bus.cell(row=idx, column=17, value=quota_val)
            q_c.border = THIN_BORDER
            q_c.fill = LIGHT_TEAL
            
            r_c = ws_bus.cell(row=idx, column=18, value=f"=Q{idx}-P{idx}")
            r_c.border = THIN_BORDER
            r_c.fill = RED_FILL

        # Enterprise Sheet
        ws_ent = wb.create_sheet(title='Update Enterprise')
        ws_ent.cell(row=1, column=2, value="REGION").fill = GRAY_HEADER
        ws_ent.merge_cells("B1:H1")
        ws_ent.cell(row=2, column=2, value="Enterprise").fill = FNB_TEAL
        ws_ent.cell(row=2, column=2).font = WHITE_BOLD_FONT
        for c_idx, reg in enumerate(["Cape", "Gauteng-North", "Gauteng South and Central", "Inland", "KwaZulu-Natal", "Total"], start=3):
            cell = ws_ent.cell(row=2, column=c_idx, value=reg)
            cell.fill = FNB_TEAL if c_idx < 8 else FNB_ORANGE
            cell.font = WHITE_BOLD_FONT

        ent_rows = [
            ("Eastern Cape", [55, 0, 0, 0, 0]), ("Free State", [0, 0, 0, 21, 0]),
            ("Gauteng East", [0, 89, 0, 0, 0]), ("Gauteng Klipriver", [0, 69, 0, 0, 0]),
            ("Gauteng South-West", [0, 80, 0, 0, 0]), ("Gauteng Tshwane", [0, 0, 60, 0, 0]),
            ("Greater Sandton", [0, 0, 56, 0, 0]), ("KZN Coastal", [0, 0, 0, 0, 87]),
            ("KZN Inland", [0, 0, 0, 0, 50]), ("Limpopo", [0, 0, 0, 46, 0]),
            ("Midrand", [0, 0, 56, 0, 0]), ("Mpumalanga", [0, 0, 0, 67, 0]),
            ("North West", [0, 0, 0, 41, 0]), ("Northern Cape", [0, 0, 0, 26, 0]),
            ("Western Cape Inland", [35, 0, 0, 0, 0]), ("Western Cape Metro", [53, 0, 0, 0, 0])
        ]
        for idx, (prov, vals) in enumerate(ent_rows, start=3):
            ws_ent.cell(row=idx, column=2, value=prov).border = THIN_BORDER
            for v_idx, val in enumerate(vals, start=3):
                c = ws_ent.cell(row=idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            tot_c = ws_ent.cell(row=idx, column=8, value=f"=SUM(C{idx}:G{idx})")
            tot_c.border = THIN_BORDER
            tot_c.fill = LIGHT_ORANGE

        ent_tot_row = len(ent_rows) + 3
        ws_ent.cell(row=ent_tot_row, column=2, value="TOTAL EXCL R10 to R60MIL").fill = FNB_TEAL
        ws_ent.cell(row=ent_tot_row, column=2).font = WHITE_BOLD_FONT
        for c_idx in range(3, 8):
            col_let = openpyxl.utils.get_column_letter(c_idx)
            c = ws_ent.cell(row=ent_tot_row, column=c_idx, value=f"=SUM({col_let}3:{col_let}{ent_tot_row-1})")
            c.fill = FNB_TEAL
            c.font = WHITE_BOLD_FONT
            c.border = THIN_BORDER
        ent_sum = ws_ent.cell(row=ent_tot_row, column=8, value=f"=SUM(H3:H{ent_tot_row-1})")
        ent_sum.fill = FNB_ORANGE
        ent_sum.font = WHITE_BOLD_FONT

        # Enterprise Right Side Table
        ws_ent.cell(row=1, column=11, value="REGION").fill = GRAY_HEADER
        ws_ent.merge_cells("K1:R1")
        ws_ent.cell(row=2, column=10, value="Enterprise").fill = FNB_TEAL
        ws_ent.cell(row=2, column=10).font = WHITE_BOLD_FONT
        for c_idx, sc in enumerate(["Cape", "Gauteng-North", "Gauteng South and Central", "Inland", "KwaZulu-Natal", "Total", "Quota", "Outstanding"], start=11):
            cell = ws_ent.cell(row=2, column=c_idx, value=sc)
            cell.fill = FNB_TEAL if c_idx < 16 else (FNB_ORANGE if c_idx == 17 else RED_FILL)
            cell.font = WHITE_BOLD_FONT

        for idx, (s_name, s_vals, quota_val) in enumerate([
            ("R10-R60M", [56, 105, 107, 109, 82], q_e1),
            ("R60-R150M", [43, 38, 90, 55, 21], q_e2),
            ("R150M+", [44, 29, 41, 37, 34], q_e3)
        ], start=3):
            ws_ent.cell(row=idx, column=10, value=s_name).border = THIN_BORDER
            for v_idx, val in enumerate(s_vals, start=11):
                c = ws_ent.cell(row=idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            tot_seg = ws_ent.cell(row=idx, column=16, value=f"=SUM(K{idx}:O{idx})")
            tot_seg.border = THIN_BORDER
            tot_seg.fill = LIGHT_ORANGE
            
            q_c = ws_ent.cell(row=idx, column=17, value=quota_val)
            q_c.border = THIN_BORDER
            q_c.fill = LIGHT_TEAL
            
            r_c = ws_ent.cell(row=idx, column=18, value=f"=Q{idx}-P{idx}")
            r_c.border = THIN_BORDER
            r_c.fill = RED_FILL

        # PUBSC Sheet
        ws_pub = wb.create_sheet(title='Update PUBSC')
        ws_pub.append(["", "REGION"])
        ws_pub.append(["", "EASTERN CAPE", "FREE STATE", "GAUTENG", "KWAZULU-NATAL", "LIMPOPO", "MPUMALANGA", "NORTH WEST", "NORTHERN CAPE", "WESTERN CAPE", "TOTAL"])
        for col_idx in range(2, 12):
            cell = ws_pub.cell(row=2, column=col_idx)
            cell.fill = FNB_TEAL if col_idx < 12 else FNB_ORANGE
            cell.font = WHITE_BOLD_FONT
            cell.alignment = CENTER_ALIGN

        for prow in [
            ["NON-PROFIT ORGANISATION", 4, 1, 80, 5, 5, 2, 3, 2, 4, "=SUM(C4:K4)"],
            ["PUBLIC SECTOR COLLEGES & FET'S", 0, 0, 1, 0, 0, 0, 0, 0, 0, "=SUM(C5:K5)"],
            ["PUBLIC SECTOR EMBASSIES", 0, 0, 2, 0, 0, 0, 0, 0, 0, "=SUM(C6:K6)"],
            ["PUBLIC SECTOR LOCAL GOVERMENT", 0, 0, 1, 0, 0, 0, 0, 1, 0, "=SUM(C7:K7)"],
            ["PUBLIC SECTOR PROVINCIAL GOVER", 0, 0, 1, 0, 0, 0, 0, 0, 0, "=SUM(C8:K8)"],
            ["PUBLIC SECTOR PUBLIC SCHOOLS", 8, 1, 34, 12, 7, 5, 2, 0, 1, "=SUM(C9:K9)"],
            ["PUBLIC SECTOR UNIONS & POLITIC", 0, 0, 1, 0, 0, 0, 0, 0, 1, "=SUM(C10:K10)"]
        ]:
            ws_pub.append([""] + prow)

        for sheet in wb.worksheets:
            for col in sheet.columns:
                max_len = 0
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                for cell in col:
                    if cell.value is not None:
                        val_str = str(cell.value)
                        if len(val_str) > max_len: max_len = len(val_str)
                sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

        wb.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    st.markdown("---")
    if st.button("📥 Generate & Download Exact PM Update Workbook", type="primary", key="download_status_btn"):
        status_excel_bytes = generate_exact_pm_update_workbook()
        run_date_str = datetime.now().strftime("%Y-%m-%d")
        st.success("🎉 Project Status Update report generated successfully!")
        st.download_button(
            label="💾 Download Formatted Excel Report (`Star Detailed Update.xlsx`)",
            data=status_excel_bytes,
            file_name=f"Star Detailed Update-W22 {run_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_status_excel_final"
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
    st.info("Upload your 911 SPSS (.sav) files here to run pipeline automation.")


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
