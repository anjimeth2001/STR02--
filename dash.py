import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side

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

# --- Update session state for uploaded files ---
for key, file in zip(st.session_state.files_uploaded.keys(),
                     [exhaust_file, finishing_file, hank_file, gre_file, dye_file, wf_file]):
    st.session_state.files_uploaded[key] = file  # sets None if cleared

# --- Display upload status ---
uploaded_files = []
file_names = ['Exhaust Consolidate Plan', 'Finishing PPO', 'Hank PPO', 'GRE Status', 'Dye PPO', 'WF PPO']
files = list(st.session_state.files_uploaded.values())

for name, file in zip(file_names, files):
    uploaded_files.append(f"✅ {name}" if file else f"❌ {name}")

st.write("**Upload Status:**")
col_status1, col_status2 = st.columns(2)
for i in range(0, len(uploaded_files), 2):
    col_status1.write(uploaded_files[i])
for i in range(1, len(uploaded_files), 2):
    if i < len(uploaded_files):
        col_status2.write(uploaded_files[i])

# --- Sheet selection for main dataset ---
if st.session_state.files_uploaded['exhaust_file'] is not None:
    excel_file = pd.ExcelFile(st.session_state.files_uploaded['exhaust_file'])
    sheet_options = excel_file.sheet_names
    selected_sheet = st.selectbox("Select Main Dataset Sheet", sheet_options)
else:
    selected_sheet = None

# --- Merge datasets if all uploaded and sheet selected ---
if all(files) and selected_sheet:
    with st.spinner("Merging datasets..."):
        # Load selected sheet
        df_main = pd.read_excel(st.session_state.files_uploaded['exhaust_file'], sheet_name=selected_sheet)
        original_row_count = len(df_main)

        df_gre = pd.read_excel(st.session_state.files_uploaded['gre_file'])
        df_finishing = pd.read_excel(st.session_state.files_uploaded['finishing_file'])
        df_dye = pd.read_excel(st.session_state.files_uploaded['dye_file'])
        df_hank = pd.read_excel(st.session_state.files_uploaded['hank_file'])
        df_wf = pd.read_excel(st.session_state.files_uploaded['wf_file'])

        # Strip column names
        for df in [df_main, df_gre, df_finishing, df_dye, df_hank, df_wf]:
            df.columns = df.columns.str.strip()

        # Drop duplicates in lookup tables
        if 'Origin order code' in df_gre.columns:
            df_gre = df_gre.drop_duplicates(subset=['Origin order code'], keep='first')
        for df in [df_finishing, df_dye, df_hank, df_wf]:
            if 'Prod Order' in df.columns:
                df.drop_duplicates(subset=['Prod Order'], keep='first', inplace=True)

        # Ensure main dataset has unique Production orders
        df_main = df_main.drop_duplicates(subset=['Production order'], keep='first')

        # Merge GRE Status using mapping
        if 'Origin order code' in df_gre.columns:
            gre_status = df_gre.set_index('Origin order code')['Receiving status'].to_dict()
            gre_datetime = df_gre.set_index('Origin order code')['Last update DateTime Cmp/Div'].to_dict()
            df_main['Receiving status'] = df_main['Production order'].map(gre_status).fillna('-')
            df_main['Last update DateTime Cmp/Div'] = df_main['Production order'].map(gre_datetime).fillna('-')
        else:
            df_main['Receiving status'] = '-'
            df_main['Last update DateTime Cmp/Div'] = '-'

        # Merge PPO tables safely
        def merge_ppo_safe(df_main, df_ppo, col_name):
            if 'Prod Order' in df_ppo.columns and 'Operation' in df_ppo.columns:
                ppo_map = df_ppo.set_index('Prod Order')['Operation'].to_dict()
                df_main[col_name] = df_main['Production order'].map(ppo_map).fillna('-')
            else:
                df_main[col_name] = '-'
            return df_main

        df_main = merge_ppo_safe(df_main, df_finishing, 'Finishing PPO')
        df_main = merge_ppo_safe(df_main, df_dye, 'Dye PPO')
        df_main = merge_ppo_safe(df_main, df_hank, 'Hank PPO')
        df_main = merge_ppo_safe(df_main, df_wf, 'WF PPO')

        final_row_count = len(df_main)
        st.session_state.merged_df = df_main.copy()
        st.session_state.row_info = {'original': original_row_count, 'final': final_row_count}

    # Show row count verification
    
    with st.expander("Preview Merged Data"):
        st.dataframe(df_main, use_container_width=True)

# --- Download button ---
if 'merged_df' in st.session_state and all(files) and selected_sheet:
    df_main = st.session_state.merged_df

    output = BytesIO()
    df_main.to_excel(output, index=False)
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    # Borders & font
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    font = Font(name='Aptos Narrow', size=9, color='000000')
    for row in ws.iter_rows():
        for cell in row:
            cell.font = font
            cell.border = thin_border

    # Adjust column widths
    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    formatted_output = BytesIO()
    wb.save(formatted_output)
    formatted_output.seek(0)

    st.download_button(
        label="📥 Download Formatted Dataset",
        data=formatted_output,
        file_name=f"Combined_{selected_sheet}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
