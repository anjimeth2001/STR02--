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
        uploaded_files.append(f"✅ {name}")
    else:
        uploaded_files.append(f"❌ {name}")

st.write("**Upload Status:**")
col_status1, col_status2 = st.columns(2)
for i in range(0, len(uploaded_files), 2):
    col_status1.write(uploaded_files[i])
for i in range(1, len(uploaded_files), 2):
    if i < len(uploaded_files):
        col_status2.write(uploaded_files[i])

# --- Merge datasets if all uploaded ---
if all(files):
    with st.spinner("Merging datasets..."):
        # Load main sheet "Dye plan 8.22" explicitly
        df_main = pd.read_excel(
    st.session_state.files_uploaded['exhaust_file'], 
    sheet_name=[s for s in pd.ExcelFile(st.session_state.files_uploaded['exhaust_file']).sheet_names if s.lower().startswith("dye plan")][-1]
)

        
        # Ensure we have exactly 37 rows from main dataset
        original_row_count = len(df_main)
        
        df_gre = pd.read_excel(st.session_state.files_uploaded['gre_file'])
        df_finishing = pd.read_excel(st.session_state.files_uploaded['finishing_file'])
        df_dye = pd.read_excel(st.session_state.files_uploaded['dye_file'])
        df_hank = pd.read_excel(st.session_state.files_uploaded['hank_file'])
        df_wf = pd.read_excel(st.session_state.files_uploaded['wf_file'])

        # Strip column names
        for df in [df_main, df_gre, df_finishing, df_dye, df_hank, df_wf]:
            df.columns = df.columns.str.strip()

        # Remove duplicates from lookup tables to avoid row multiplication
        # Only keep the FIRST occurrence of each Production order match
        if 'Origin order code' in df_gre.columns:
            df_gre = df_gre.drop_duplicates(subset=['Origin order code'], keep='first')
        
        if 'Prod Order' in df_finishing.columns:
            df_finishing = df_finishing.drop_duplicates(subset=['Prod Order'], keep='first')
        if 'Prod Order' in df_dye.columns:
            df_dye = df_dye.drop_duplicates(subset=['Prod Order'], keep='first')
        if 'Prod Order' in df_hank.columns:
            df_hank = df_hank.drop_duplicates(subset=['Prod Order'], keep='first')
        if 'Prod Order' in df_wf.columns:
            df_wf = df_wf.drop_duplicates(subset=['Prod Order'], keep='first')

        # Also ensure main dataset has unique Production orders
        df_main = df_main.drop_duplicates(subset=['Production order'], keep='first')

        # Merge GRE Status using mapping to avoid row multiplication
        if 'Origin order code' in df_gre.columns:
            gre_dict_status = df_gre.set_index('Origin order code')['Receiving status'].to_dict()
            gre_dict_datetime = df_gre.set_index('Origin order code')['Last update DateTime Cmp/Div'].to_dict()
            
            df_main['Receiving status'] = df_main['Production order'].map(gre_dict_status).fillna('-')
            df_main['Last update DateTime Cmp/Div'] = df_main['Production order'].map(gre_dict_datetime).fillna('-')
        else:
            df_main['Receiving status'] = '-'
            df_main['Last update DateTime Cmp/Div'] = '-'

        # Function to merge PPOs safely - ensure no row multiplication
        def merge_ppo_safe(df_main, df_ppo, col_name):
            if 'Prod Order' in df_ppo.columns and 'Operation' in df_ppo.columns:
                # Create a mapping dictionary to avoid merge issues
                ppo_dict = df_ppo.set_index('Prod Order')['Operation'].to_dict()
                df_main[col_name] = df_main['Production order'].map(ppo_dict).fillna('-')
            else:
                df_main[col_name] = '-'
            return df_main

        # Merge all PPO tables
        df_main = merge_ppo_safe(df_main, df_finishing, 'Finishing PPO')
        df_main = merge_ppo_safe(df_main, df_dye, 'Dye PPO')
        df_main = merge_ppo_safe(df_main, df_hank, 'Hank PPO')
        df_main = merge_ppo_safe(df_main, df_wf, 'WF PPO')

        # Verify row count hasn't changed
        final_row_count = len(df_main)
        
        # Save merged df in session state
        st.session_state.merged_df = df_main.copy()
        st.session_state.row_info = {
            'original': original_row_count,
            'final': final_row_count
        }

    # Show row count verification
    if st.session_state.row_info['original'] == st.session_state.row_info['final']:
        st.success(f"✅ Data merged successfully! Maintained {st.session_state.row_info['final']} rows from original dataset.")
    else:
        st.warning(f"⚠️ Row count changed: {st.session_state.row_info['original']} → {st.session_state.row_info['final']}")

    # Preview merged data
    with st.expander("Preview Merged Data"):
        st.dataframe(df_main, use_container_width=True)

# --- Download button always visible if merged df exists ---
if 'merged_df' in st.session_state:
    df_main = st.session_state.merged_df

    # Format Excel
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

    # Save formatted Excel
    formatted_output = BytesIO()
    wb.save(formatted_output)
    formatted_output.seek(0)

    st.download_button(
        label="📥 Download Formatted Dataset",
        data=formatted_output,
        file_name="Combined_Exhaust_Consolidate.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

