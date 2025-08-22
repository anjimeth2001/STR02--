import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side
from datetime import datetime

st.set_page_config(page_title="Exhaust Consolidate Plan Dashboard", layout="wide")
st.markdown("Upload all datasets to merge and download a formatted combined Excel file.")

# --- Session state for uploaded files ---
if 'files_uploaded' not in st.session_state:
    st.session_state.files_uploaded = {
        'exhaust_file': None,
        'finishing_file': None,
        'hank_file': None,
        'gre_file': None,
        'dye_file': None,
        'wf_file': None
    }

# --- File uploaders ---
col1, col2 = st.columns(2)
with col1:
    exhaust_file = st.file_uploader("Exhaust Consolidate Plan", type=["xlsx"], key="exhaust")
    finishing_file = st.file_uploader("Finishing PPO", type=["xlsx"], key="finishing")
    hank_file = st.file_uploader("Hank PPO", type=["xlsx"], key="hank")
with col2:
    gre_file = st.file_uploader("GRE Status", type=["xlsx"], key="gre")
    dye_file = st.file_uploader("Dye PPO", type=["xlsx"], key="dye")
    wf_file = st.file_uploader("WF PPO", type=["xlsx"], key="wf")

# --- Save uploaded files in session state ---
for key, file in zip(st.session_state.files_uploaded.keys(),
                     [exhaust_file, finishing_file, hank_file, gre_file, dye_file, wf_file]):
    if file is not None:
        st.session_state.files_uploaded[key] = file

# --- Display upload status ---
uploaded_files = []
file_names = ['Exhaust Consolidate Plan', 'Finishing PPO', 'Hank PPO', 'GRE Status', 'Dye PPO', 'WF PPO']
files = list(st.session_state.files_uploaded.values())

for name, file in zip(file_names, files):
    if file:
        upload_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        uploaded_files.append(f"✅ {name} (Uploaded at {upload_time})")
    else:
        uploaded_files.append(f"❌ {name} (Not uploaded)")

st.write("**Upload Status:**")
col_status1, col_status2 = st.columns(2)
for i in range(0, len(uploaded_files), 2):
    col_status1.write(uploaded_files[i])
for i in range(1, len(uploaded_files), 2):
    if i < len(uploaded_files):
        col_status2.write(uploaded_files[i])

# --- Merge datasets ---
if all(files):
    with st.spinner("Merging datasets..."):
        # Auto-detect today's dye plan sheet
        excel_file = pd.ExcelFile(st.session_state.files_uploaded['exhaust_file'])
        dye_sheet = [s for s in excel_file.sheet_names if "Dye plan" in s][0]
        df_main = pd.read_excel(st.session_state.files_uploaded['exhaust_file'], sheet_name=dye_sheet)

        df_gre = pd.read_excel(st.session_state.files_uploaded['gre_file'])
        df_finishing = pd.read_excel(st.session_state.files_uploaded['finishing_file'])
        df_dye = pd.read_excel(st.session_state.files_uploaded['dye_file'])
        df_hank = pd.read_excel(st.session_state.files_uploaded['hank_file'])
        df_wf = pd.read_excel(st.session_state.files_uploaded['wf_file'])

        # Strip column names
        for df in [df_main, df_gre, df_finishing, df_dye, df_hank, df_wf]:
            df.columns = df.columns.str.strip()

        # Merge GRE Status (drop duplicates)
        gre_merge = df_gre[['Origin order code', 'Receiving status', 'Last update DateTime Cmp/Div']].drop_duplicates()
        df_main = df_main.merge(
            gre_merge,
            left_on='Production order',
            right_on='Origin order code',
            how='left'
        ).drop(columns=['Origin order code']).drop_duplicates()

        df_main['Receiving status'].fillna('-', inplace=True)
        df_main['Last update DateTime Cmp/Div'].fillna('-', inplace=True)

        # Function to merge PPOs safely (drop duplicates each time)
        def merge_ppo(df_main, df_ppo, col_name):
            ppo_merge = df_ppo[['Prod Order', 'Operation']].drop_duplicates()
            df_main = df_main.merge(
                ppo_merge,
                left_on='Production order',
                right_on='Prod Order',
                how='left'
            ).drop(columns=['Prod Order', 'Operation']).drop_duplicates()
            df_main[col_name] = df_main[col_name] if col_name in df_main else '-'
            return df_main

        # Merge PPO tables
        df_main = merge_ppo(df_main, df_finishing, 'Finishing PPO')
        df_main = merge_ppo(df_main, df_dye, 'Dye PPO')
        df_main = merge_ppo(df_main, df_hank, 'Hank PPO')
        df_main = merge_ppo(df_main, df_wf, 'WF PPO')

        # Final deduplication -> ensures only original Dye plan rows
        df_main = df_main.drop_duplicates(subset=df_main.columns)

        st.session_state.merged_df = df_main.copy()

    st.success(f"✅ Data merged successfully! ({len(df_main)} rows)")

    with st.expander("Preview Merged Data"):
        st.dataframe(df_main, use_container_width=True)

# --- Download button ---
if 'merged_df' in st.session_state:
    df_main = st.session_state.merged_df

    output = BytesIO()
    df_main.to_excel(output, index=False)
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    font = Font(name='Aptos Narrow', size=9, color='000000')
    for row in ws.iter_rows():
        for cell in row:
            cell.font = font
            cell.border = thin_border

    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    formatted_output = BytesIO()
    wb.save(formatted_output)
    formatted_output.seek(0)

    st.download_button(
        label="📥 Download Formatted Dataset",
        data=formatted_output,
        file_name="Combined_Exhaust_Consolidate.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
