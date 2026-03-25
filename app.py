import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="HS Code Processor", layout="wide")

st.title("📦 Excel HS Code & COO Processor")

# -----------------------
# 🧠 SMART MAPPING
# -----------------------
def smart_map(columns):
    mapping = {}

    def find_col(keywords):
        for col in columns:
            col_lower = col.lower()
            if any(k in col_lower for k in keywords):
                return col
        return None

    mapping["QTY"] = find_col(["qty", "quantity"])
    mapping["Description"] = find_col(["desc", "description", "item"])
    mapping["Amount"] = find_col(["amount", "value", "price"])
    mapping["HS Code"] = find_col(["hs", "code"])
    mapping["COO"] = find_col(["coo", "origin", "country"])
    mapping["GW"] = find_col(["gw", "gross"])
    mapping["NW"] = find_col(["nw", "net"])

    return mapping


# -----------------------
# 📂 FILE UPLOAD
# -----------------------
uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    # -----------------------
    # 🔗 MERGE FILES
    # -----------------------
    df_list = []

    for file in uploaded_files:
        temp_df = pd.read_excel(file)
        temp_df["Source File"] = file.name
        df_list.append(temp_df)

    df = pd.concat(df_list, ignore_index=True)

    st.success(f"✅ {len(uploaded_files)} file(s) merged successfully")
    st.subheader("Preview")
    st.dataframe(df.head())

    # -----------------------
    # 🔗 SMART MAPPING UI
    # -----------------------
    st.subheader("🔗 Column Mapping")

    columns = df.columns.tolist()
    auto_map = smart_map(columns)

    def get_index(col_name):
        return columns.index(col_name) if col_name in columns else 0

    col_qty = st.selectbox("QTY", columns, index=get_index(auto_map["QTY"]))
    col_desc = st.selectbox("Description", columns, index=get_index(auto_map["Description"]))
    col_amount = st.selectbox("Amount", columns, index=get_index(auto_map["Amount"]))
    col_hs = st.selectbox("HS Code", columns, index=get_index(auto_map["HS Code"]))
    col_coo = st.selectbox("COO", columns, index=get_index(auto_map["COO"]))
    col_gw = st.selectbox("GW", columns, index=get_index(auto_map["GW"]))
    col_nw = st.selectbox("NW", columns, index=get_index(auto_map["NW"]))

    # -----------------------
    # 🚀 PROCESS BUTTON
    # -----------------------
    if st.button("Process File"):

        # Rename columns
        df_processed = df.rename(columns={
            col_qty: "QTY",
            col_desc: "Description",
            col_amount: "Amount",
            col_hs: "HS Code",
            col_coo: "COO",
            col_gw: "GW",
            col_nw: "NW"
        })

        df_processed = df_processed[
            ["QTY", "Description", "Amount", "HS Code", "COO", "GW", "NW"]
        ].copy()

        # -----------------------
        # 🧹 CLEAN DATA
        # -----------------------
        def clean_numeric(col):
            return (
                df_processed[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.strip()
            )

        for col in ["QTY", "Amount", "GW", "NW", "HS Code"]:
            df_processed[col] = clean_numeric(col)

        # Convert numeric fields
        for col in ["QTY", "Amount", "GW", "NW"]:
            df_processed[col] = pd.to_numeric(df_processed[col], errors="coerce")

        # HS Code validation (numeric but kept as string later)
        hs_numeric = pd.to_numeric(df_processed["HS Code"], errors="coerce")

        # -----------------------
        # ❌ VALIDATION
        # -----------------------
        errors = {}

        for col in ["QTY", "Amount", "GW", "NW"]:
            invalid_rows = df_processed[df_processed[col].isna()]
            if not invalid_rows.empty:
                errors[col] = invalid_rows.index.tolist()

        invalid_hs = df_processed[hs_numeric.isna()]
        if not invalid_hs.empty:
            errors["HS Code"] = invalid_hs.index.tolist()

        if errors:
            st.error("❌ Data validation failed. Fix the following rows:")

            for col, rows in errors.items():
                st.write(f"**{col}** invalid at rows: {rows}")

            st.dataframe(df_processed)
            st.stop()

        # Convert HS Code back to string
        df_processed["HS Code"] = df_processed["HS Code"].astype(str)

        # -----------------------
        # 📄 SHEET 1
        # -----------------------
        sheet1 = df_processed.sort_values(by=["HS Code", "COO"])

        # -----------------------
        # 📄 SHEET 2
        # -----------------------
        sheet2 = df_processed.groupby("COO").agg({
            "QTY": "sum",
            "Amount": "sum",
            "GW": "sum",
            "NW": "sum"
        }).reset_index()

        sheet2["Description"] = "Auto Spare Parts"

        sheet2 = sheet2[
            ["COO", "Description", "QTY", "Amount", "GW", "NW"]
        ]

        # -----------------------
        # 📤 EXPORT
        # -----------------------
        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            sheet1.to_excel(writer, index=False, sheet_name="HS Code Grouped")
            sheet2.to_excel(writer, index=False, sheet_name="COO Summary")

        output.seek(0)

        st.success("✅ File processed successfully!")

        st.download_button(
            label="📥 Download Result",
            data=output,
            file_name="processed_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )