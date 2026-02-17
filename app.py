import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Final Packaging List Generator")
st.title("Final Packaging List Generator")
st.write("Upload both Excel files to generate the final packaging list.")

uploaded_order = st.file_uploader("Upload Order list.xlsx", type=["xlsx"])
uploaded_packing = st.file_uploader("Upload Packing list.xlsx", type=["xlsx"])

if uploaded_packing and uploaded_order:
    try:
        # Load packing list
        packing_df = pd.read_excel(uploaded_packing)
        packing_df.columns = packing_df.columns.str.strip().str.upper().str.replace(" ", "")
        
        # Define standard column mapping for packing list
        column_map_packing = {
            "PARTDESC": "PARTDESC",
            "QUANTITY": "QUANTITY",
            "CARTONNO": "CARTONNO",
            "REF1": "REF1",
            "WEIGHT": "WEIGHT",
            "NETVALUE": "NETVALUE",
            "CRTNWEIGHT": "CRTNWEIGHT",
            "MANFPART": "MANFPART"
        }

        # Check all mapped columns exist in packing_df
        for col in column_map_packing.keys():
            if col not in packing_df.columns:
                st.error(f"Packing list missing column: {col}")
                st.stop()

        # Keep only required columns
        packing_df = packing_df[list(column_map_packing.keys())].copy()

        # Remove zero quantity rows
        packing_df = packing_df[packing_df["QUANTITY"] > 0].copy()

        # Load Order list
        order_df = pd.read_excel(uploaded_order)
        order_df.columns = order_df.columns.str.strip().str.upper().str.replace(" ", "")

        # Check Order list has Partref and Brand
        if "PARTREF" not in order_df.columns or "BRAND" not in order_df.columns:
            st.error("Order list must contain 'Partref' and 'Brand' columns.")
            st.stop()

        # Prepare Order list for merge
        order_df = order_df[["PARTREF", "BRAND"]].copy()
        order_df = order_df.rename(columns={"PARTREF": "PARTNO", "BRAND": "BRAND"})

        # Add PARTNO and Brand to packing list
        final_df = packing_df.copy()
        if len(order_df) != len(final_df):
            st.warning("Row count mismatch! Matching by order of rows.")
        final_df["PARTNO"] = order_df["PARTNO"].values
        final_df["Brand"] = order_df["BRAND"].values

        # MANFPART: if blank, copy PARTNO
        final_df["MANFPART"] = final_df["MANFPART"].fillna(final_df["PARTNO"])
        final_df.loc[final_df["MANFPART"].astype(str).str.strip() == "", "MANFPART"] = final_df["PARTNO"]

        # UNIT PRICE calculation
        final_df["UNIT PRICE"] = final_df["NETVALUE"] / final_df["QUANTITY"]

        # Add fixed columns
        final_df["HSCODE"] = "87089900"
        final_df["COO"] = ""
        final_df["ORDER NUMBER"] = ""
        final_df["REFERENCES"] = ""

        # Reorder columns for final sheet
        final_df = final_df[[
            "PARTNO",
            "Brand",
            "PARTDESC",
            "QUANTITY",
            "CARTONNO",
            "REF1",
            "WEIGHT",
            "MANFPART",
            "CRTNWEIGHT",
            "UNIT PRICE",
            "NETVALUE",
            "HSCODE",
            "COO",
            "ORDER NUMBER",
            "REFERENCES"
        ]]

        # Add SL.NO
        final_df.insert(0, "SL.NO", range(1, len(final_df) + 1))

        # Rename columns to match final packaging list exactly
        final_df = final_df.rename(columns={
            "PARTDESC": "PART DESC",
            "QUANTITY": "QTY",
            "CRTNWEIGHT": "CRTN WEIGHT",
            "NETVALUE": "AMOUNT AED",
            "CARTONNO": "CARTONNO"
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


