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

# --- Display upload status with date and time ---
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

# --- Merge datasets if all uploaded ---
if all(files):
    with st.spinner("Merging datasets..."):
        # Load main sheet "Dye plan 8.22" explicitly
        df_main = pd.read_excel(st.session_state.files_uploaded['exhaust_file'], sheet_name="Dye plan 8.22")
        df_gre = pd.read_excel(st.session_state.files_uploaded['gre_file'])
        df_finishing = pd.read_excel(st.session_state.files_uploaded['finishing_file'])
        df_dye = pd.read_excel(st.session_state.files_uploaded['dye_file'])
        df_hank = pd.read_excel(st.session_state.files_uploaded['hank_file'])
        df_wf = pd.read_excel(st.session_state.files_uploaded['wf_file'])

        # Strip column names
        for df in [df_main, df_gre, df_finishing, df_dye, df_hank, df_wf]:
            df.columns = df.columns.str.strip()

        # Merge GRE Status
        gre_merge = df_gre[['Origin order code', 'Receiving status', 'Last update DateTime Cmp/Div']].copy()
        df_main = df_main.merge(
            gre_merge,
            left_on='Production order',
            right_on='Origin order code',
            how='left'
        )
        df_main['Receiving status'].fillna('-', inplace=True)
        df_main['Last update DateTime Cmp/Div'].fillna('-', inplace=True)
        df_main.drop(columns=['Origin order code'], inplace=True)

        # Function to merge PPOs
        def merge_ppo(df_main, df_ppo, col_name):
            ppo_merge = df_ppo[['Prod Order', 'Operation']].copy()
            df_main = df_main.merge(
                ppo_merge,
                left_on='Production order',
                right_on='Prod Order',
                how='left'
            )
            df_main[col_name] = df_main['Operation'].fillna('-')
            df_main.drop(columns=['Operation', 'Prod Order'], inplace=True)
            return df_main

        # Merge all PPO tables
        for df, name in zip([df_finishing, df_dye, df_hank, df_wf],
                            ['Finishing PPO', 'Dye PPO', 'Hank PPO', 'WF PPO']):
            df_main = merge_ppo(df_main, df, name)

        # Save merged df in session state
        st.session_state.merged_df = df_main.copy()

    st.success("✅ Data merged successfully!")

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


