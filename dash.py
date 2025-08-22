import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook

st.set_page_config(page_title="Dye Plan Consolidate Dashboard", layout="wide")

# --- File uploader section ---
st.title("📊 Dye Plan Consolidate Dashboard")

if "files_uploaded" not in st.session_state:
    st.session_state.files_uploaded = {
        "exhaust_file": None,
        "gre_file": None,
        "finishing_file": None,
        "packing_file": None,
        "shipment_file": None,
    }
if "upload_times" not in st.session_state:
    st.session_state.upload_times = {}

uploaded_exhaust = st.file_uploader("📂 Upload Consolidate Dye Plan file", type=["xlsx"], key="exhaust")
uploaded_gre = st.file_uploader("📂 Upload GRE Status file", type=["xlsx"], key="gre")
uploaded_finishing = st.file_uploader("📂 Upload Finishing PPO file", type=["xlsx"], key="finishing")
uploaded_packing = st.file_uploader("📂 Upload Packing PPO file", type=["xlsx"], key="packing")
uploaded_shipment = st.file_uploader("📂 Upload Shipment PPO file", type=["xlsx"], key="shipment")

# Save to session_state
for k, v in {
    "exhaust_file": uploaded_exhaust,
    "gre_file": uploaded_gre,
    "finishing_file": uploaded_finishing,
    "packing_file": uploaded_packing,
    "shipment_file": uploaded_shipment,
}.items():
    if v is not None:
        st.session_state.files_uploaded[k] = v
        st.session_state.upload_times[k] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Display upload status with date/time
st.subheader("📌 Upload Status")
for k, v in st.session_state.files_uploaded.items():
    if v:
        st.write(f"✅ {k} uploaded at {st.session_state.upload_times[k]}")
    else:
        st.write(f"❌ {k} not uploaded")

# --- Processing ---
if all(st.session_state.files_uploaded.values()):
    try:
        # Find latest Dye plan sheet
        xls = pd.ExcelFile(st.session_state.files_uploaded["exhaust_file"])
        dye_sheets = [s for s in xls.sheet_names if s.lower().startswith("dye plan")]
        latest_sheet = sorted(dye_sheets)[-1]

        df_main = pd.read_excel(st.session_state.files_uploaded["exhaust_file"], sheet_name=latest_sheet)
        df_gre = pd.read_excel(st.session_state.files_uploaded["gre_file"])
        df_finishing = pd.read_excel(st.session_state.files_uploaded["finishing_file"])
        df_packing = pd.read_excel(st.session_state.files_uploaded["packing_file"])
        df_shipment = pd.read_excel(st.session_state.files_uploaded["shipment_file"])

        # --- GRE merge (deduplicate) ---
        gre_merge = (
            df_gre[["Origin order code", "Receiving status", "Last update DateTime Cmp/Div"]]
            .drop_duplicates(subset=["Origin order code"], keep="last")
        )
        df_main = df_main.merge(
            gre_merge,
            left_on="Production order",
            right_on="Origin order code",
            how="left"
        )
        df_main.drop(columns=["Origin order code"], inplace=True)

        # --- PPO merge helper ---
        def merge_ppo(df_main, df_ppo, col_name):
            ppo_merge = (
                df_ppo[["Prod Order", "Operation"]]
                .drop_duplicates(subset=["Prod Order"], keep="last")
            )
            df_main = df_main.merge(
                ppo_merge,
                left_on="Production order",
                right_on="Prod Order",
                how="left"
            )
            df_main[col_name] = df_main["Operation"].fillna("-")
            df_main.drop(columns=["Operation", "Prod Order"], inplace=True)
            return df_main

        # Apply PPO merges
        df_main = merge_ppo(df_main, df_finishing, "Finishing_PPO")
        df_main = merge_ppo(df_main, df_packing, "Packing_PPO")
        df_main = merge_ppo(df_main, df_shipment, "Shipment_PPO")

        st.success(f"✅ Consolidated successfully! Final rows: {len(df_main)}")

        # --- Download updated workbook ---
        exhaust_file = st.session_state.files_uploaded["exhaust_file"]
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            book = load_workbook(exhaust_file)
            writer.book = book
            writer.sheets = {ws.title: ws for ws in book.worksheets}
            # overwrite the latest Dye plan sheet
            df_main.to_excel(writer, sheet_name=latest_sheet, index=False)
            writer.close()

        st.download_button(
            label="💾 Download Updated Consolidate Dye Plan",
            data=output.getvalue(),
            file_name="Updated_Dye_Plan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
