import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Final Packaging List Generator")
st.title("Final Packaging List Generator")
st.write("Upload both Excel files to generate the final packaging list.")

uploaded_order = st.file_uploader("Upload Order list.xlsx", type=["xlsx"])
uploaded_packing = st.file_uploader("Upload Packing list.xlsx", type=["xlsx"])

if uploaded_packing:
    try:
        # Load packing list
        packing_df = pd.read_excel(uploaded_packing)
        packing_df.columns = packing_df.columns.str.strip()

        # Map packing list columns
        column_map_packing = {
            "PartNo": "PARTNO",
            "PartDesc": "PARTDESC",
            "Quantity": "QUANTITY",
            "CartonNo": "CARTONNO",
            "Ref1": "REF1",
            "Weight": "WEIGHT",
            "NetValue": "NETVALUE",
            "CrtWeight": "CRTNWEIGHT",
            "MANFPART": "MANFPART"
        }
        packing_df = packing_df.rename(columns=column_map_packing)

        # Validate required packing columns
        required_packing_cols = [
            "CARTONNO", "PARTNO", "PARTDESC", "QUANTITY",
            "REF1", "WEIGHT", "NETVALUE", "CRTNWEIGHT",
            "MANFPART"
        ]
        for col in required_packing_cols:
            if col not in packing_df.columns:
                st.error(f"Missing column in Packing list.xlsx: {col}")
                st.stop()

        # Remove zero quantity rows
        packing_df = packing_df[packing_df["QUANTITY"] > 0].copy()

        # Initialize BRAND column as blank
        packing_df["BRAND"] = ""

        # Load Order list if provided
        if uploaded_order:
            order_df = pd.read_excel(uploaded_order)
            order_df.columns = order_df.columns.str.strip()

            # Standardize column names in Order list
            column_map_order = {
                "Partnumber": "PARTNO",
                "PartNumber": "PARTNO",
                "PARTNO": "PARTNO",
                "Brand": "BRAND",
                "PRICE": "PRICE",
                "Price": "PRICE"
            }
            order_df = order_df.rename(columns=column_map_order)

            # Keep only relevant columns
            order_cols = ["PARTNO"]
            if "BRAND" in order_df.columns:
                order_cols.append("BRAND")
            if "PRICE" in order_df.columns:
                order_cols.append("PRICE")

            order_df = order_df[order_cols]

            # Merge Packing list with Order list on PARTNO
            merged_df = pd.merge(
                packing_df,
                order_df,
                on="PARTNO",
                how="left",
                suffixes=("", "_ORDER")
            )

            # BRAND strictly from Order list (no fallback)
            if "BRAND" in merged_df.columns:
                merged_df["BRAND"] = merged_df["BRAND"]
            else:
                merged_df["BRAND"] = ""
        else:
            merged_df = packing_df.copy()

        # MANFPART: if blank, copy PARTNO
        merged_df["MANFPART"] = merged_df["MANFPART"].fillna(merged_df["PARTNO"])
        merged_df.loc[merged_df["MANFPART"].astype(str).str.strip() == "", "MANFPART"] = merged_df["PARTNO"]

        # UNIT PRICE calculation
        merged_df["UNIT PRICE"] = merged_df["NETVALUE"] / merged_df["QUANTITY"]
        if "PRICE" in merged_df.columns:
            merged_df["UNIT PRICE"] = merged_df["UNIT PRICE"].fillna(merged_df["PRICE"])

        # Build final DataFrame
        final_df = pd.DataFrame({
            "SL.NO": range(1, len(merged_df) + 1),
            "CARTONNO": merged_df["CARTONNO"],
            "Brand": merged_df["BRAND"],  # strictly from Order list
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
