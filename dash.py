import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Exhaust Consolidate Plan Dashboard", layout="wide")

# --- File Uploads ---
exhaust_file = st.file_uploader("Upload Consolidate Dye Plan Excel file", type=["xlsx"], key="exhaust_file")
post_file = st.file_uploader("Upload POST Excel file", type=["xlsx"], key="post_file")
gre_file = st.file_uploader("Upload GRE Excel file", type=["xlsx"], key="gre_file")
ppos_file = st.file_uploader("Upload PPOs Excel file", type=["xlsx"], key="ppos_file")

if exhaust_file and post_file and gre_file and ppos_file:
    try:
        # --- Find latest Dye plan sheet ---
        xls = pd.ExcelFile(exhaust_file)
        dyeplan_sheets = [s for s in xls.sheet_names if s.lower().startswith("dye plan")]
        if not dyeplan_sheets:
            st.error("No 'Dye plan xx.xx' sheets found in the file.")
            st.stop()
        latest_sheet = sorted(dyeplan_sheets)[-1]  # pick latest
        st.success(f"Using latest sheet: {latest_sheet}")

        # --- Load DataFrames ---
        df_main = pd.read_excel(exhaust_file, sheet_name=latest_sheet)
        post_df = pd.read_excel(post_file)
        gre_df = pd.read_excel(gre_file)
        ppos_df = pd.read_excel(ppos_file)

        # --- Merge with GRE ---
        if "GRE Prod Order" in gre_df.columns and "Production Order" in df_main.columns:
            df_main = df_main.merge(
                gre_df[["GRE Prod Order", "Project"]],
                left_on="Production Order",
                right_on="GRE Prod Order",
                how="left"
            ).drop(columns=["GRE Prod Order"])
        else:
            st.warning("GRE merge skipped (columns missing).")

        # --- Merge with PPOs ---
        if "Production Order" in ppos_df.columns and "Production Order" in df_main.columns:
            df_main = df_main.merge(
                ppos_df[["Production Order", "PPO"]],
                on="Production Order",
                how="left"
            )
        else:
            st.warning("PPOs merge skipped (columns missing).")

        # --- Merge with POST ---
        if "Production Order" in post_df.columns and "Production Order" in df_main.columns:
            df_main = df_main.merge(
                post_df[["Production Order", "POST"]],
                on="Production Order",
                how="left"
            )
        else:
            st.warning("POST merge skipped (columns missing).")

        st.dataframe(df_main.head())

        # --- Replace latest sheet inside original workbook ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            book = load_workbook(exhaust_file)
            writer.book = book
            writer.sheets = {ws.title: ws for ws in book.worksheets}

            # Delete old sheet if exists
            if latest_sheet in writer.book.sheetnames:
                std = writer.book[latest_sheet]
                writer.book.remove(std)

            # Write new updated sheet
            df_main.to_excel(writer, sheet_name=latest_sheet, index=False)

        # --- Download button ---
        st.download_button(
            label="📥 Download Updated Consolidate Plan",
            data=output.getvalue(),
            file_name="Updated_Consolidate_Plan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")
