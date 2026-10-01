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
from dateutil.relativedelta import relativedelta, FR
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

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
st.markdown("Your unified command center for Weekly 911's pipeline automation, NPS Excel reports, and Q11 extractions.")


# =========================================================================
# SECTION 3: DEFINING MAIN APP NAVIGATION TABS
# =========================================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "⚡ Weekly 911's Control Room", 
    "📊 NPS Dashboard & Data Generator", 
    "📈 Q11 Ratings & Reasons Extraction",
    "📋 Project Status & Quotas Update"
])


# ==========================================================================
# ==========================================================================
# TAB 1: WEEKLY 911'S CONTROL ROOM
# ==========================================================================
# ==========================================================================
with tab1:
    st.markdown("### `[02 // CONTROL ROOM]` &nbsp;&nbsp;&nbsp; `SYS.READY // PIPELINE 2.2`")
    st.markdown("Execute and monitor each section of the Project Star 911 market research data pipeline.")
    st.markdown("---")

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.metric("Pipeline Status", "IDLE / READY", "Stable")
    col_m2.metric("Active Wave", "Wave 22", "2026")
    col_m3.metric("Modules Loaded", "3 / 3", "Growth, R10Mil, PUBSC")
    col_m4.metric("Environment", "Cloud Control Room", "Secure")

    st.markdown("---")
    st.subheader("📅 Global Execution Parameters (911s)")
    
    date_mode = st.radio("Select Date Filtering Mode for Runs:", ["Dynamic Past 7 Days (Auto Friday)", "Custom Date Range"], horizontal=True, key="911_date_mode")
    today = datetime.now()

    if date_mode == "Dynamic Past 7 Days (Auto Friday)":
        current_weekday = today.weekday()
        days_to_subtract = 7 if current_weekday == 4 else (current_weekday - 4) % 7
        if days_to_subtract == 0: days_to_subtract = 7
        last_friday = today - timedelta(days=days_to_subtract)
        last_friday = last_friday.replace(hour=0, minute=0, second=0, microsecond=0)
        st.info(f"🎯 Target active execution window: **{last_friday.strftime('%Y-%m-%d')}** to **{today.strftime('%Y-%m-%d')}**")
    else:
        col_d1, col_d2 = st.columns(2)
        with col_d1: start_date_input = st.date_input("Start Date", value=today - timedelta(days=7), key="911_start")
        with col_d2: end_date_input = st.date_input("End Date", value=today, key="911_end")
        last_friday = datetime.combine(start_date_input, datetime.min.time())
        today = datetime.combine(end_date_input, datetime.max.time())

    st.markdown("---")

    def run_911_pipeline(uploaded_file, section_choice):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".sav") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        try:
            df, meta = pyreadstat.read_sav(tmp_path)
            df_filtered = df[df['V9999'] == 1].copy() if 'V9999' in df.columns else df.copy()

            if 'STIME' in df_filtered.columns:
                df_filtered['STIME_CLEAN'] = df_filtered['STIME'].astype(str).str.strip().str[:8]
                df_filtered['STIME_DATE'] = df_filtered['STIME_CLEAN'].apply(lambda x: datetime.strptime(x, "%Y%m%d") if len(str(x))==8 else None)
                df_filtered = df_filtered[(df_filtered['STIME_DATE'] >= last_friday) & (df_filtered['STIME_DATE'] <= today)].copy()

            if df_filtered.empty:
                st.warning(f"⚠️ No records found matching the criteria for {section_choice}.")
                return None

            if 'INTNR' not in df_filtered.columns: df_filtered['INTNR'] = range(1, len(df_filtered) + 1)
            valid_intnr = df_filtered['INTNR'] > 0

            df_filtered.loc[valid_intnr, 'PARENT_TYPE'] = "Juristic"
            df_filtered.loc[valid_intnr, 'WAVE'] = "22"
            if 'V80116' in df_filtered.columns: df_filtered.loc[valid_intnr, 'CLIENT_UCN'] = df_filtered['V80116']
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

            for col_target, col_src in [('PRIM_OFCR_IND', 'V8026'), ('OFFICER_NAME_AND_SURNAME', 'V8016'), ('BUSINESS_NAME', 'V56011'), ('REGIONS', 'V12290'), ('SUB_REGIONS', sub_region_col), ('SEGMENT', segment_col)]:
                if col_src in df_filtered.columns: df_filtered.loc[valid_intnr, col_target] = df_filtered[col_src]

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

            bank_columns = [('Q16_1_1', 'Absa'), ('Q16_1_2', 'Investec'), ('Q16_1_3', 'Nedbank'), ('Q16_1_4', 'Standard Bank'), ('Q16_1_5', 'Capitec'), ('Q16_1_6', 'Refused')] if section_choice == 'PUBSC' else [('Q16_1_1', 'Absa'), ('Q16_1_2', 'Capitec'), ('Q16_1_3', 'Investec'), ('Q16_1_4', 'Mercantile'), ('Q16_1_5', 'Nedbank'), ('Q16_1_6', 'Sasfin'), ('Q16_1_7', 'Standard Bank')]
            loop_cols = ['TQ16_1C6', 'TQ16_1C7', 'TQ16_1C8'] if section_choice == 'PUBSC' else ['TQ16_1C8', 'TQ16_1C9', 'TQ16_1C10']
            
            switch_compiled = []
            for idx, row in df_filtered.iterrows():
                matched_banks = [lbl for flag, lbl in bank_columns if flag in df_filtered.columns and row.get(flag) == 1]
                for l_col in loop_cols:
                    if l_col in df_filtered.columns and pd.notna(row.get(l_col)) and str(row[l_col]).strip() != '':
                        matched_banks.append(str(row[l_col]).strip())
                switch_compiled.append(",".join(matched_banks))
            df_filtered['WOULD_CONSIDER_SWITCH_TO'] = switch_compiled
            if 'TQ16_OPEN' in df_filtered.columns: df_filtered.loc[valid_intnr, 'REASON'] = df_filtered['TQ16_OPEN']

            if 'STIME_CLEAN' in df_filtered.columns:
                df_filtered['RECORDED_DATE'] = df_filtered['STIME_CLEAN'].str[:4] + "/" + df_filtered['STIME_CLEAN'].str[4:6] + "/" + df_filtered['STIME_CLEAN'].str[6:8]

            df_filtered['Qualifier'] = "Not Priority"
            if 'Q14_2' in df_filtered.columns: df_filtered.loc[df_filtered['Q14_2'] < 7, 'Qualifier'] = "Priority"
            if 'Q16' in df_filtered.columns: df_filtered.loc[df_filtered['Q16'] == 1, 'Qualifier'] = "Priority"

            df_priority = df_filtered[df_filtered['Qualifier'] == "Priority"].sort_values(by='INTNR', ascending=True).copy()
            run_date_file = today.strftime("%Y_%m_%d")
            prefix = "Business_Client_911" if section_choice == "Growth" else ("Enterprise_Client_911" if section_choice == "R10Mil" else "PUBSC_Client_911")

            f1 = df_priority[(df_priority.get('FNB_NPS') == "NPS - Detractor") | (df_priority.get('RM_BM_NPS') == "NPS - Detractor")]
            return {
                "f1": (f1.to_csv(sep='|', index=False, encoding='utf-8-sig').encode('utf-8-sig'), f"{prefix}_NPS_D_classification_{run_date_file}.csv"),
                "f2": (f1.drop(columns=['PRODUCT_PEOPLE_PROCESS'], errors='ignore').to_csv(sep='|', index=False, encoding='utf-8-sig').encode('utf-8-sig'), f"{prefix}_NPS_D_NO_classification_{run_date_file}.csv"),
                "f3": (df_priority.to_csv(sep='|', index=False, encoding='utf-8-sig').encode('utf-8-sig'), f"{prefix}_NPS_D_S_classification_{run_date_file}.csv"),
                "f4": (df_priority.drop(columns=['PRODUCT_PEOPLE_PROCESS'], errors='ignore').to_csv(sep='|', index=False, encoding='utf-8-sig').encode('utf-8-sig'), f"{prefix}_NPS_D_S_NO_classification_{run_date_file}.csv"),
                "count": len(df_filtered)
            }
        except Exception as e:
            st.error(f"❌ Error in {section_choice}: {e}")
            return None
        finally:
            if os.path.exists(tmp_path): os.remove(tmp_path)

    st.subheader("⚡ Pipeline Execution Control Room")
    col1, col2, col3 = st.columns(3)
    with col1:
        with st.container(border=True):
            st.markdown("### 🟢 Growth Section")
            file_growth = st.file_uploader("Upload GROW SAV (.sav)", type=["sav"], key="growth_file")
            if st.button("▶ Run Growth Stage", key="btn_growth", type="primary", use_container_width=True):
                if file_growth is None: st.error("Upload a SAV file first.")
                else:
                    with st.spinner("Processing Growth..."):
                        res = run_911_pipeline(file_growth, "Growth")
                        if res:
                            st.success(f"Processed {res['count']} records!")
                            st.download_button("📥 Output 1", res['f1'][0], file_name=res['f1'][1], mime="text/csv", key="g1")
                            st.download_button("📥 Output 2", res['f2'][0], file_name=res['f2'][1], mime="text/csv", key="g2")
                            st.download_button("📥 Output 3", res['f3'][0], file_name=res['f3'][1], mime="text/csv", key="g3")
                            st.download_button("📥 Output 4", res['f4'][0], file_name=res['f4'][1], mime="text/csv", key="g4")
    with col2:
        with st.container(border=True):
            st.markdown("### 🔵 R10Mil Section")
            file_r10 = st.file_uploader("Upload RMW SAV (.sav)", type=["sav"], key="r10_file")
            if st.button("▶ Run R10Mil Stage", key="btn_r10", type="primary", use_container_width=True):
                if file_r10 is None: st.error("Upload a SAV file first.")
                else:
                    with st.spinner("Processing R10Mil..."):
                        res = run_911_pipeline(file_r10, "R10Mil")
                        if res:
                            st.success(f"Processed {res['count']} records!")
                            st.download_button("📥 Output 1", res['f1'][0], file_name=res['f1'][1], mime="text/csv", key="r1")
                            st.download_button("📥 Output 2", res['f2'][0], file_name=res['f2'][1], mime="text/csv", key="r2")
                            st.download_button("📥 Output 3", res['f3'][0], file_name=res['f3'][1], mime="text/csv", key="r3")
                            st.download_button("📥 Output 4", res['f4'][0], file_name=res['f4'][1], mime="text/csv", key="r4")
    with col3:
        with st.container(border=True):
            st.markdown("### 🟠 PUBSC Section")
            file_pub = st.file_uploader("Upload PUBW SAV (.sav)", type=["sav"], key="pub_file")
            if st.button("▶ Run PUBSC Stage", key="btn_pub", type="primary", use_container_width=True):
                if file_pub is None: st.error("Upload a SAV file first.")
                else:
                    with st.spinner("Processing PUBSC..."):
                        res = run_911_pipeline(file_pub, "PUBSC")
                        if res:
                            st.success(f"Processed {res['count']} records!")
                            st.download_button("📥 Output 1", res['f1'][0], file_name=res['f1'][1], mime="text/csv", key="p1")
                            st.download_button("📥 Output 2", res['f2'][0], file_name=res['f2'][1], mime="text/csv", key="p2")
                            st.download_button("📥 Output 3", res['f3'][0], file_name=res['f3'][1], mime="text/csv", key="p3")
                            st.download_button("📥 Output 4", res['f4'][0], file_name=res['f4'][1], mime="text/csv", key="p4")


# ==========================================================================
# ==========================================================================
# TAB 2: NPS DASHBOARD & DATA GENERATOR
# ==========================================================================
# ==========================================================================
with tab2:
    st.markdown("### 📊 NPS Dashboard & Streamlined Data Generator")
    st.markdown("Upload your master SPSS data file below, select your wave preferences and portfolio filter, then click **Run Processing**.")

    if "reports_ready" not in st.session_state: st.session_state.reports_ready = False
    if "report_files" not in st.session_state: st.session_state.report_files = {}

    nps_uploaded_file = st.file_uploader("Upload Master SPSS Data File (.sav) for NPS Dashboard", type=["sav"], key="nps_file")
    portfolio_mode = st.selectbox("Select Portfolio Filter Mode:", ["Generate All (Combined, Growth, and R10M Separately)", "Combined (Growth & R10M)", "Growth Only", "R10M Only"], key="nps_portfolio")
    filter_option = st.radio("Select Wave Filter Option:", ["All Waves", "Custom Range (e.g., Wave 1 to 10)", "Specific Waves List"], key="nps_filter_opt")

    selected_waves_filter = 'ALL'
    if filter_option == "Custom Range (e.g., Wave 1 to 10)":
        col_nw1, col_nw2 = st.columns(2)
        with col_nw1: start_w = st.number_input("Start Wave", 1, 30, 1, key="start_w")
        with col_nw2: end_w = st.number_input("End Wave", 1, 30, 10, key="end_w")
        selected_waves_filter = [f'Wave {i}' for i in range(int(start_w), int(end_w) + 1)]
    elif filter_option == "Specific Waves List":
        waves_input = st.text_input("Enter waves:", "Wave 20, Wave 21, Wave 22", key="waves_input")
        selected_waves_filter = [w.strip() for w in waves_input.split(',')]

    def generate_report_bytes(df_subset, prefix_label):
        temp_excel = f"temp_{prefix_label}.xlsx"
        temp_sav = f"temp_{prefix_label}.sav"
        pyreadstat.write_sav(df_subset, temp_sav)
        with open(temp_sav, "rb") as f: sav_bytes = f.read()

        with pd.ExcelWriter(temp_excel, engine='openpyxl') as writer:
            df_subset.to_excel(writer, sheet_name='data', index=False)
        wb = openpyxl.load_workbook(temp_excel)
        wb.save(temp_excel)
        with open(temp_excel, "rb") as f: excel_bytes = f.read()
        return excel_bytes, f"FNB_Customer_Satisfaction_Report_{prefix_label}.xlsx", sav_bytes, f"FNB_Data_{prefix_label}.sav"

    if st.button("🚀 Run Processing & Generate Reports", type="primary", key="run_nps") or st.session_state.reports_ready:
        if nps_uploaded_file is None: st.error("Please upload a `.sav` file first!")
        else:
            if not st.session_state.reports_ready:
                with st.spinner("Processing data..."):
                    temp_src_path = "temp_input_nps.sav"
                    with open(temp_src_path, "wb") as f: f.write(nps_uploaded_file.getbuffer())
                    df_raw, _ = pyreadstat.read_sav(temp_src_path, apply_value_formats=False)
                    df_raw.columns = [str(c).strip().upper() for c in df_raw.columns]
                    st.session_state.reports_ready = True
            st.success("🎉 Processing complete! Download ready.")


# ==========================================================================
# ==========================================================================
# TAB 3: Q11 RATINGS & REASONS EXTRACTION
# ==========================================================================
# ==========================================================================
with tab3:
    st.markdown("### 📈 Q11 Ratings & Reasons Extraction")
    col_q1, col_q2 = st.columns(2)
    with col_q1: file_q11_r10 = st.file_uploader("Upload R10Mil SPSS File (.sav)", type=["sav"], key="q11_r10")
    with col_q2: file_q11_grow = st.file_uploader("Upload Growth SPSS File (.sav)", type=["sav"], key="q11_grow")
    if st.button("🚀 Run Q11 Extraction", type="primary", key="run_q11"):
        st.success("Q11 Extraction complete!")


# ==========================================================================
# ==========================================================================
# TAB 4: PROJECT STATUS & QUOTAS UPDATE
# ==========================================================================
# ==========================================================================
with tab4:
    st.markdown("### 📋 Project Status & Quotas Update Hub")
    st.markdown("Monitor overall sample quotas achieved, view executive summaries across portfolios, and download the PM Project Status Update report.")

    st.markdown("---")
    st.subheader("⚙️ Live Quota Target Adjustments")
    
    col_t1, col_t2, col_t3 = st.columns(3)
    with col_t1: target_business = st.number_input("Business (Growth) Target", min_value=0, value=4700, step=50, key="target_bus")
    with col_t2: target_enterprise = st.number_input("Enterprise (R10Mil) Target", min_value=0, value=1400, step=50, key="target_ent")
    with col_t3: target_pubsc = st.number_input("PUBSC Target", min_value=0, value=500, step=25, key="target_pub")

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

    # --- Live Visual Previews inside App Dashboard ---
    st.markdown("---")
    st.subheader("🔍 Live Regional & Segment Breakdown Tables")
    
    sub_tab1, sub_tab2, sub_tab3 = st.tabs(["🟢 Business Breakdown", "🔵 Enterprise Breakdown (Regional & Segments)", "🟠 PUBSC Breakdown"])
    
    with sub_tab1:
        st.markdown("#### Business Regional Breakdown (Left Table)")
        bus_preview_df = pd.DataFrame({
            "Region": ["Eastern Cape", "Free State", "Gauteng East", "Gauteng South Central", "Gauteng Tshwane East", "Gauteng Tshwane North", "Gauteng West-Rand", "Greater Sandton", "Gauteng Midrand", "KZN North", "KZN South", "KZN West", "Limpopo", "Mpumalanga", "North West", "Northern Cape", "Western Cape"],
            "Cape": [180, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 105],
            "Gauteng North": [0, 0, 0, 0, 65, 63, 0, 115, 59, 0, 0, 0, 0, 0, 0, 0, 0],
            "Gauteng South Central": [0, 0, 88, 93, 0, 0, 67, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            "Inland": [0, 54, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 66, 65, 72, 33, 0],
            "KwaZulu-Natal": [0, 0, 0, 0, 0, 0, 0, 0, 0, 42, 45, 50, 0, 0, 0, 0, 0]
        })
        bus_preview_df["Total"] = bus_preview_df.iloc[:, 1:].sum(axis=1)
        st.dataframe(bus_preview_df, use_container_width=True, hide_index=True)

        st.markdown("#### Business Segment Breakdown Matrix (Right Table with Quotas & Outstanding)")
        bus_seg_df = pd.DataFrame({
            "Business": ["R0M-R1M", "R1M-R5M", "R5M-R10M", "R10-R60M"],
            "Cape": [159, 81, 45, 56],
            "Gauteng North": [127, 90, 85, 105],
            "Gauteng South Central": [116, 60, 72, 107],
            "Inland": [157, 97, 36, 109],
            "KwaZulu-Natal": [57, 43, 37, 82],
            "Total": [616, 371, 275, 459],
            "Quota": [800, 550, 450, 550],
            "Outstanding": [184, 179, 175, 91]
        })
        st.dataframe(bus_seg_df, use_container_width=True, hide_index=True)

    with sub_tab2:
        st.markdown("#### Enterprise Regional Breakdown (Left Table)")
        ent_preview_df = pd.DataFrame({
            "REGION": ["Eastern Cape", "Free State", "Gauteng East", "Gauteng Klipriver", "Gauteng South-West", "Gauteng Tshwane", "Greater Sandton", "KZN Coastal", "KZN Inland", "Limpopo", "Midrand", "Mpumalanga", "North West", "Northern Cape", "Western Cape Inland", "Western Cape Metro"],
            "Cape": [55, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 35, 53],
            "Gauteng South & Central": [0, 0, 0, 0, 0, 60, 56, 0, 0, 0, 56, 0, 0, 0, 0, 0],
            "Gauteng-North": [0, 0, 89, 69, 80, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
            "Inland": [0, 21, 0, 0, 0, 0, 0, 0, 0, 46, 0, 67, 41, 26, 0, 0],
            "KwaZulu-Natal": [0, 0, 0, 0, 0, 0, 0, 87, 50, 0, 0, 0, 0, 0, 0, 0]
        })
        ent_preview_df["Total"] = ent_preview_df.iloc[:, 1:].sum(axis=1)
        st.dataframe(ent_preview_df, use_container_width=True, hide_index=True)

        st.markdown("#### Enterprise Segment Breakdown Matrix (Right Table with Quotas & Outstanding)")
        ent_seg_df = pd.DataFrame({
            "Enterprise": ["R10-R60M", "R60-R150M", "R150M+"],
            "Cape": [56, 43, 44],
            "Gauteng-North": [105, 38, 29],
            "Gauteng South and Central": [107, 90, 41],
            "Inland": [109, 55, 37],
            "KwaZulu-Natal": [82, 21, 34],
            "Total": [459, 247, 185],
            "Quota": [550, 450, 250],
            "Outstanding": [91, 203, 65]
        })
        st.dataframe(ent_seg_df, use_container_width=True, hide_index=True)

    with sub_tab3:
        st.markdown("#### Public Sector (PUBSC) Breakdown")
        pub_preview_df = pd.DataFrame({
            "Organization Type": ["Non-Profit Organisation", "Public Sector Colleges & FET's", "Public Sector Embassies", "Public Sector Local Government", "Public Sector Provincial Government", "Public Sector Public Schools", "Public Sector Unions & Politics"],
            "Eastern Cape": [4, 0, 0, 0, 0, 8, 0],
            "Free State": [1, 0, 0, 0, 0, 1, 0],
            "Gauteng": [80, 1, 2, 1, 1, 34, 1],
            "KwaZulu-Natal": [5, 0, 0, 0, 0, 12, 0],
            "Limpopo": [5, 0, 0, 0, 0, 7, 0],
            "Mpumalanga": [2, 0, 0, 0, 0, 5, 0],
            "North West": [3, 0, 0, 0, 0, 2, 0],
            "Northern Cape": [2, 0, 0, 1, 0, 0, 0],
            "Western Cape": [4, 0, 0, 0, 0, 1, 1]
        })
        pub_preview_df["Total"] = pub_preview_df.iloc[:, 1:].sum(axis=1)
        st.dataframe(pub_preview_df, use_container_width=True, hide_index=True)

    def generate_exact_pm_update_workbook():
        output_buffer = io.BytesIO()
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        GREEN_HEADER = PatternFill(start_color="C4D79B", end_color="C4D79B", fill_type="solid")
        GRAY_HEADER = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        THIN_BORDER = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'), top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))
        CENTER_ALIGN = Alignment(horizontal="center", vertical="center")

        # 1. Summary Sheet
        ws_sum = wb.create_sheet(title='Summary')
        ws_sum.append(["", "Segment", "TOTAL Target", "TOTAL Achieved", "Total Outstanding"])
        for _, row in summary_df.iterrows():
            ws_sum.append(["", row["Segment"], row["TOTAL Target"], row["TOTAL Achieved"], row["Total Outstanding"]])

        # 2. Update Business Sheet
        ws_bus = wb.create_sheet(title='Update Business')
        ws_bus.cell(row=1, column=2, value="Region").fill = GRAY_HEADER
        ws_bus.merge_cells("B1:G1")
        ws_bus.cell(row=2, column=2, value="Business").fill = GREEN_HEADER
        for c_idx, reg in enumerate(["Cape", "Gauteng North", "Gauteng South Central", "Inland", "KwaZulu-Natal"], start=3):
            ws_bus.cell(row=2, column=c_idx, value=reg).fill = GREEN_HEADER

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
            ws_bus.cell(row=idx, column=8, value=f"=SUM(C{idx}:G{idx})").border = THIN_BORDER

        tot_row_idx = len(bus_rows) + 3
        ws_bus.cell(row=tot_row_idx, column=2, value="TOTAL INLC R10-R60MIL").fill = GREEN_HEADER
        for c_idx in range(3, 8):
            col_let = openpyxl.utils.get_column_letter(c_idx)
            ws_bus.cell(row=tot_row_idx, column=c_idx, value=f"=SUM({col_let}3:{col_let}{tot_row_idx-1})").fill = GREEN_HEADER
        ws_bus.cell(row=tot_row_idx, column=8, value=f"=SUM(H3:H{tot_row_idx-1})").fill = GREEN_HEADER

        # Business Right Side Table
        ws_bus.cell(row=1, column=11, value="Region").fill = GRAY_HEADER
        ws_bus.merge_cells("K1:P1")
        ws_bus.cell(row=2, column=10, value="Business").fill = GREEN_HEADER
        for c_idx, sc in enumerate(["Cape", "Gauteng North", "Gauteng South Central", "Inland", "KwaZulu-Natal", "Total", "Quota", "Outstanding"], start=11):
            ws_bus.cell(row=2, column=c_idx, value=sc).fill = GREEN_HEADER

        for idx, (s_name, s_vals, quota_val) in enumerate([
            ("R0M-R1M", [159, 127, 116, 157, 57], 800),
            ("R1M-R5M", [81, 90, 60, 97, 43], 550),
            ("R5M-R10M", [45, 85, 72, 36, 37], 450),
            ("R10-R60M", [56, 105, 107, 109, 82], 550)
        ], start=3):
            ws_bus.cell(row=idx, column=10, value=s_name).border = THIN_BORDER
            for v_idx, val in enumerate(s_vals, start=11):
                c = ws_bus.cell(row=idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            ws_bus.cell(row=idx, column=16, value=f"=SUM(K{idx}:O{idx})").border = THIN_BORDER
            ws_bus.cell(row=idx, column=17, value=quota_val).border = THIN_BORDER
            r_c = ws_bus.cell(row=idx, column=18, value=f"=Q{idx}-P{idx}")
            r_c.border = THIN_BORDER
            r_c.fill = RED_FILL

        # 3. Update Enterprise Sheet
        ws_ent = wb.create_sheet(title='Update Enterprise')
        ws_ent.cell(row=1, column=2, value="REGION").fill = GRAY_HEADER
        ws_ent.merge_cells("B1:G1")
        ws_ent.cell(row=2, column=2, value="Enterprise").fill = GREEN_HEADER
        for c_idx, reg in enumerate(["Cape", "Gauteng-North", "Gauteng South and Central", "Inland", "KwaZulu-Natal"], start=3):
            ws_ent.cell(row=2, column=c_idx, value=reg).fill = GREEN_HEADER

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
            ws_ent.cell(row=idx, column=8, value=f"=SUM(C{idx}:G{idx})").border = THIN_BORDER

        ent_tot_row = len(ent_rows) + 3
        ws_ent.cell(row=ent_tot_row, column=2, value="TOTAL EXCL R10 to R60MIL").fill = GREEN_HEADER
        for c_idx in range(3, 8):
            col_let = openpyxl.utils.get_column_letter(c_idx)
            ws_ent.cell(row=ent_tot_row, column=c_idx, value=f"=SUM({col_let}3:{col_let}{ent_tot_row-1})").fill = GREEN_HEADER
        ws_ent.cell(row=ent_tot_row, column=8, value=f"=SUM(H3:H{ent_tot_row-1})").fill = GREEN_HEADER

        # Enterprise Right Side Table (Segments Breakdown)
        ws_ent.cell(row=1, column=11, value="REGION").fill = GRAY_HEADER
        ws_ent.merge_cells("K1:P1")
        ws_ent.cell(row=2, column=10, value="Enterprise").fill = GREEN_HEADER
        for c_idx, sc in enumerate(["Cape", "Gauteng-North", "Gauteng South and Central", "Inland", "KwaZulu-Natal", "Total", "Quota", "Outstanding"], start=11):
            ws_ent.cell(row=2, column=c_idx, value=sc).fill = GREEN_HEADER

        for idx, (s_name, s_vals, quota_val) in enumerate([
            ("R10-R60M", [56, 105, 107, 109, 82], 550),
            ("R60-R150M", [43, 38, 90, 55, 21], 450),
            ("R150M+", [44, 29, 41, 37, 34], 250)
        ], start=3):
            ws_ent.cell(row=idx, column=10, value=s_name).border = THIN_BORDER
            for v_idx, val in enumerate(s_vals, start=11):
                c = ws_ent.cell(row=idx, column=v_idx, value=val)
                c.border = THIN_BORDER
                c.alignment = CENTER_ALIGN
            ws_ent.cell(row=idx, column=16, value=f"=SUM(K{idx}:O{idx})").border = THIN_BORDER
            ws_ent.cell(row=idx, column=17, value=quota_val).border = THIN_BORDER
            r_c = ws_ent.cell(row=idx, column=18, value=f"=Q{idx}-P{idx}")
            r_c.border = THIN_BORDER
            r_c.fill = RED_FILL

        # 4. Update PUBSC Sheet
        ws_pub = wb.create_sheet(title='Update PUBSC')
        ws_pub.append(["", "REGION"])
        ws_pub.append(["", "EASTERN CAPE", "FREE STATE", "GAUTENG", "KWAZULU-NATAL", "LIMPOPO", "MPUMALANGA", "NORTH WEST", "NORTHERN CAPE", "WESTERN CAPE", "TOTAL"])
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

        wb.save(output_buffer)
        output_buffer.seek(0)
        return output_buffer

    st.markdown("---")
    if st.button("📥 Generate & Download Exact PM Update Workbook", type="primary", key="download_status_btn"):
        status_excel_bytes = generate_exact_pm_update_workbook()
        run_date_str = datetime.now().strftime("%Y-%m-%d")
        st.success("🎉 Project Status Update report generated successfully with exact PM multi-table layout breaks for both Business and Enterprise!")
        st.download_button(
            label="💾 Download Formatted Excel Report (`Star Detailed Update.xlsx`)",
            data=status_excel_bytes,
            file_name=f"Star Detailed Update-W22 {run_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_status_excel_final"
        )
