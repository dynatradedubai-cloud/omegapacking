import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Final Packaging List Generator")
st.title("Final Packaging List Generator")
st.write("Upload both Excel files to generate the final packaging list.")

uploaded_order = st.file_uploader("Upload Order list.xlsx", type=["xlsx"])
uploaded_packing = st.file_uploader("Upload Packing list.xlsx", type=["xlsx"])

if uploaded_order and uploaded_packing:
    try:
        # Load Excel files
        order_df = pd.read_excel(uploaded_order)
        packing_df = pd.read_excel(uploaded_packing)

        # Clean column names
        order_df.columns = order_df.columns.str.strip()
        packing_df.columns = packing_df.columns.str.strip()

        # Map actual Excel columns to standardized names
        # Order list: only PRICE and PARTNO needed
        column_map_order = {
            "Partnumber": "PARTNO",
            "Price": "PRICE"
        }
        order_df = order_df.rename(columns=column_map_order)

        # Packing list: full mapping
        column_map_packing = {
            "PartNo": "PARTNO",
            "PartDesc": "PARTDESC",
            "Quantity": "QUANTITY",
            "CartonNo": "CARTONNO",
            "Ref1": "REF1",
            "Weight": "WEIGHT",
            "NetValue": "NETVALUE",
            "CrtWeight": "CRTNWEIGHT",
            "MANFPART": "MANFPART",
            "Brand": "BRAND"
        }
        packing_df = packing_df.rename(columns=column_map_packing)

        # Required columns
        required_order_cols = ["PARTNO", "PRICE"]
        required_packing_cols = [
            "CARTONNO", "PARTNO", "PARTDESC", "QUANTITY",
            "REF1", "WEIGHT", "NETVALUE", "CRTNWEIGHT",
            "MANFPART", "BRAND"
        ]

        # Validate columns
        for col in required_order_cols:
            if col not in order_df.columns:
                st.warning(f"Column '{col}' not found in Order list.xlsx — will ignore.")

        for col in required_packing_cols:
            if col not in packing_df.columns:
                st.error(f"Missing column in Packing list.xlsx: {col}")
                st.stop()

        # Remove zero quantity rows
        packing_df = packing_df[packing_df["QUANTITY"] > 0].copy()

        # Merge with Order list only for PRICE
        merged_df = packing_df.merge(
            order_df[["PARTNO", "PRICE"]],
            on="PARTNO",
            how="left"
        )

        # BRAND and MANFPART always from Packing list
        merged_df["BRAND"] = merged_df["BRAND"]
        merged_df["MANFPART"] = merged_df["MANFPART"]

        # Calculate UNIT PRICE
        merged_df["UNIT PRICE"] = merged_df["NETVALUE"] / merged_df["QUANTITY"]
        merged_df["UNIT PRICE"] = merged_df["UNIT PRICE"].fillna(merged_df["PRICE"])

        # Build final DataFrame
        final_df = pd.DataFrame({
            "SL.NO": range(1, len(merged_df) + 1),
            "CARTONNO": merged_df["CARTONNO"],
            "Brand": merged_df["BRAND"],
            "PARTNO": merged_df["PARTNO"],
            "PART DESC": merged_df["PARTDESC"],
            "COO": "",
            "QTY": merged_df["QUANTITY"],
            "REF1": merged_df["REF1"],
            "WEIGHT": merged_df["WEIGHT"],
            "HSCODE": "87089900",
            "UNIT PRICE": merged_df["UNIT PRICE"].round(2),
            "AMOUNT AED": merged_df["NETVALUE"].round(2),
            "MANFPART": merged_df["MANFPART"],
            "CRTN WEIGHT": merged_df["CRTNWEIGHT"],
            "ORDER NUMBER": "",
            "REFERENCES": ""
        })

        # Convert to Excel in memory
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
