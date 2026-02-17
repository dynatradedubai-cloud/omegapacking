import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Final Packaging List Generator")
st.title("Final Packaging List Generator")
st.write("Upload both Excel files to generate the final packaging list.")

# Upload files
uploaded_order = st.file_uploader("Upload Order list.xlsx", type=["xlsx"])
uploaded_packing = st.file_uploader("Upload Packing list.xlsx", type=["xlsx"])

if uploaded_order and uploaded_packing:
    try:
        # Load packing list
        packing_df = pd.read_excel(uploaded_packing, engine="openpyxl")
        packing_df.columns = packing_df.columns.str.strip()

        # Remove zero quantity rows
        packing_df = packing_df[packing_df["QUANTITY"] > 0].copy()

        # Calculate UNIT PRICE
        packing_df["UNIT PRICE"] = packing_df["NETVALUE"] / packing_df["QUANTITY"]

        # Load Order list
        order_df = pd.read_excel(uploaded_order, engine="openpyxl")
        order_df.columns = order_df.columns.str.strip()

        # Check Brand exists in Order list
        if "Brand" not in order_df.columns:
            st.error("Order list.xlsx must contain 'Brand' column.")
            st.stop()

        # Ensure row count matches to avoid mismatch
        if len(order_df) != len(packing_df):
            st.warning("Row count mismatch! Brand will be assigned by order of rows.")
            # Optionally, you can truncate or pad Order list to match Packing list length
            min_len = min(len(order_df), len(packing_df))
            packing_df["Brand"] = order_df["Brand"].iloc[:min_len].reset_index(drop=True)
        else:
            # Assign Brand strictly from Order list
            packing_df["Brand"] = order_df["Brand"].values

        # MANFPART: if blank, copy PARTNO
        packing_df["MANFPART"] = packing_df["MANFPART"].fillna(packing_df["PARTNO"])
        packing_df.loc[packing_df["MANFPART"].astype(str).str.strip() == "", "MANFPART"] = packing_df["PARTNO"]

        # Build final DataFrame
        final_df = pd.DataFrame({
            "SL.NO": range(1, len(packing_df) + 1),
            "CARTONNO": packing_df["CARTONNO"],
            "Brand": packing_df["Brand"],  # strictly from Order list
            "PARTNO": packing_df["PARTNO"],
            "PART DESC": packing_df["PARTDESC"],
            "COO": "",
            "QTY": packing_df["QUANTITY"],
            "REF1": packing_df["REF1"],
            "WEIGHT": packing_df["WEIGHT"],
            "HSCODE": "87089900",
            "UNIT PRICE": packing_df["UNIT PRICE"],
            "AMOUNT AED": packing_df["NETVALUE"],
            "MANFPART": packing_df["MANFPART"],
            "CRTN WEIGHT": packing_df["CRTNWEIGHT"],
            "ORDER NUMBER": "",
            "REFERENCES": ""
        })

        # Export to Excel in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Sheet1")
        output.seek(0)

        st.success("Final Packaging List Generated Successfully")
        st.download_button(
            label="Download Final Packaging List",
            data=output,
            file_name="Final packaging list.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")

