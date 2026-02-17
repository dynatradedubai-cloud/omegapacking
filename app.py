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
        packing_df.columns = packing_df.columns.str.strip().str.upper()

        # Rename packing columns to standard names
        column_map_packing = {
            "PARTNO": "PARTNO",
            "PARTDESC": "PARTDESC",
            "QUANTITY": "QUANTITY",
            "CARTONNO": "CARTONNO",
            "REF1": "REF1",
            "WEIGHT": "WEIGHT",
            "NETVALUE": "NETVALUE",
            "CRTNWEIGHT": "CRTNWEIGHT",
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
            order_df.columns = order_df.columns.str.strip().str.upper()

            # Standardize Order list column names
            order_df = order_df.rename(columns={
                "PARTNUMBER": "PARTNO",
                "PART NUMBER": "PARTNO",
                "PART_NO": "PARTNO",
                "BRAND": "BRAND",
                "PRICE-AED": "PRICE",
                "PRICE": "PRICE"
            })

            # Keep only relevant columns
            keep_cols = [col for col in ["PARTNO", "BRAND", "PRICE"] if col in order_df.columns]
            order_df = order_df[keep_cols]

            # Merge Packing list with Order list on PARTNO
            merged_df = pd.merge(
                packing_df,
                order_df,
                on="PARTNO",
                how="left"
            )

            # BRAND strictly from Order list
            merged_df["BRAND"] = merged_df["BRAND"].fillna("")
        else:
            merged_df = packing_df.copy()

        # MANFPART fallback
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
