import streamlit as st
import pandas as pd
import pyreadstat
from datetime import datetime, timedelta
import os
import tempfile
import re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

st.set_page_config(
    page_title="Project Star: One-Stop Control Room",
    page_icon="⭐",
    layout="wide"
)

st.title("⭐ Project Star: One-Stop Operations Hub")
st.markdown("Your unified command center for Weekly 911's pipeline automation and NPS Excel Report generation.")

# --- TABS FOR THE ONE-STOP SHOP ---
tab1, tab2 = st.tabs(["⚡ Weekly 911's Control Room", "📊 BM/RM NPS Dashboard Portfolio & Data Generator"])

# ==========================================
# TAB 1: WEEKLY 911'S CONTROL ROOM
# ==========================================
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

    def run_911_pipeline(uploaded_file, section_choice):
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
                    try: return datetime.strptime(x, "%Y%m%d")
                    except: return None
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
                df_filtered['RECORDED_DATE'] = df_filtered['NYEAR'] + "/" + df_filtered['NMONTH'] + "/" + df_filtered['NDAY']

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
            st.caption("Target: Enterprise Client Pipeline")
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
            st.caption("Target: Public Sector Pipeline")
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


# ==========================================
# TAB 2: NPS DASHBOARD & DATA GENERATOR
# ==========================================
with tab2:
    st.markdown("### 📊 NPS Dashboard & Streamlined Data Generator")
    st.markdown("Upload your master SPSS data file below, select your wave preferences and portfolio filter, then click **Run Processing**.")

    if "reports_ready" not in st.session_state:
        st.session_state.reports_ready = False
    if "report_files" not in st.session_state:
        st.session_state.report_files = {}

    nps_uploaded_file = st.file_uploader("Upload Master SPSS Data File (Project Star_W? to W?.sav) for NPS Dashboard", type=["sav"], key="nps_file")

    portfolio_mode = st.selectbox(
        "Select Portfolio Filter Mode:",
        [
            "Generate All (Combined, Growth, and R10M Separately)",
            "Combined (Growth & R10M)",
            "Growth Only",
            "R10M Only"
        ],
        key="nps_portfolio"
    )

    filter_option = st.radio("Select Wave Filter Option:", ["All Waves", "Custom Range (e.g., Wave 1 to 10)", "Specific Waves List"], key="nps_filter_opt")

    selected_waves_filter = 'ALL'

    if filter_option == "Custom Range (e.g., Wave 1 to 10)":
        col_nw1, col_nw2 = st.columns(2)
        with col_nw1: start_w = st.number_input("Start Wave Number", min_value=1, max_value=30, value=1, key="start_w")
        with col_nw2: end_w = st.number_input("End Wave Number", min_value=1, max_value=30, value=10, key="end_w")
        selected_waves_filter = [f'Wave {i}' for i in range(int(start_w), int(end_w) + 1)]
    elif filter_option == "Specific Waves List":
        waves_input = st.text_input("Enter waves separated by commas:", "Wave 20, Wave 21, Wave 22", key="waves_input")
        selected_waves_filter = [w.strip() for w in waves_input.split(',')]

    def generate_report_bytes(df_subset, prefix_label):
        is_combined = (prefix_label == "Combined")
        temp_excel = f"temp_{prefix_label}.xlsx"
        temp_sav = f"temp_{prefix_label}.sav"

        pyreadstat.write_sav(df_subset, temp_sav)
        with open(temp_sav, "rb") as f: sav_bytes = f.read()

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
            if ws.max_row == 1 and ws['A1'].value is None: start_row = 1
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
                    if row_counter % 2 == 0: cell.fill = ZEBRA_FILL

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
                for col in ['E', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'N']: ws.column_dimensions[col].width = 11

        wb.save(temp_excel)
        with open(temp_excel, "rb") as f: excel_bytes = f.read()
        return excel_bytes, excel_name, sav_bytes, sav_name

    if st.button("🚀 Run Processing & Generate Reports", type="primary", key="run_nps") or st.session_state.reports_ready:
        if nps_uploaded_file is None:
            st.error("Please upload a `.sav` file first!")
            st.session_state.reports_ready = False
        else:
            if not st.session_state.reports_ready:
                with st.spinner("Processing data and building formatted reports..."):
                    temp_src_path = "temp_input_nps.sav"
                    with open(temp_src_path, "wb") as f: f.write(nps_uploaded_file.getbuffer())

                    df_raw, meta = pyreadstat.read_sav(temp_src_path, apply_value_formats=False)
                    df_raw.columns = [str(col).strip().upper() for col in df_raw.columns]

                    df_lbl, _ = pyreadstat.read_sav(temp_src_path, apply_value_formats=True)
                    df_lbl.columns = [str(col).strip().upper() for col in df_lbl.columns]

                    target_columns_mapping = {
                        'UNIQUEID': ['UNIQUEID', 'ID'], 'RUID': ['RUID'], 'TYPE': ['TYPE'], 'WAVE': ['WAVE'],
                        'REGION': ['REGION'], 'SUBREG': ['SUBREG'], 'SEGMENT': ['SEGMENT'],
                        'RM_NPS1': ['RM_NPS1', 'BM_NPS1', 'BMNPS01'], 'FNB_NPS1': ['FNB_NPS1', 'FNBNPS01'],
                        'PRIM_OFCR_IND': ['PRIM_OFCR_IND', 'PRIM_OFC'], 'OFFICER_NAME': ['OFFICER_NAME', 'OFFICER_M', 'OFFICER_NAME_']
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

                    df_base = pd.DataFrame()
                    for orig_col, target in rename_map.items():
                        if target in ['WAVE', 'REGION', 'SUBREG', 'SEGMENT'] and orig_col in df_lbl.columns:
                            df_base[target] = df_lbl[orig_col]
                        else:
                            df_base[target] = df_raw[orig_col]

                    if 'TYPE' in df_lbl.columns: df_base['TYPE'] = df_lbl[orig_type_col]

                    if selected_waves_filter != 'ALL' and 'WAVE' in df_base.columns:
                        df_base = df_base[df_base['WAVE'].isin(selected_waves_filter)].copy()

                    def clean_officer_name(name):
                        if pd.isna(name): return name
                        s = str(name).strip()
                        s = re.sub(r'\s*\(.*?\)', '', s)
                        return ' '.join(s.split())

                    if 'OFFICER_NAME' in df_base.columns: df_base['OFFICER_NAME'] = df_base['OFFICER_NAME'].apply(clean_officer_name)

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
                    st.download_button(label=f"📥 Download {label} Excel Dashboard", data=files["excel_bytes"], file_name=files["excel_name"], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"excel_{label}")
                with col_b:
                    st.download_button(label=f"📥 Download {label} .sav File", data=files["sav_bytes"], file_name=files["sav_name"], mime="application/octet-stream", key=f"sav_{label}")
