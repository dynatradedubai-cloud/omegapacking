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
        packing_df.columns = packing_df.columns.str.strip()

        # Map packing list columns
        column_map_packing = {
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

        # Remove zero quantity rows
        packing_df = packing_df[packing_df["QUANTITY"] > 0].copy()

        # Load Order list
        order_df = pd.read_excel(uploaded_order)
        order_df.columns = order_df.columns.str.strip()

        # Map Order list columns
        # "Partref" (J column) becomes PARTNO in final
        # "Brand" copied to Brand column in final
        if "Partref" not in order_df.columns or "Brand" not in order_df.columns:
            st.error("Order list must contain 'Partref' and 'Brand' columns.")
            st.stop()

        order_df = order_df[["Partref", "Brand"]]
        order_df = order_df.rename(columns={"Partref": "PARTNO", "Brand": "BRAND"})

        # Build final DataFrame
        # Merge Packing list with Order list on row index (assume same order) if needed
        final_df = packing_df.copy()
        final_df["PARTNO"] = order_df["PARTNO"]
        final_df["Brand"] = order_df["BRAND"]

        # MANFPART: if blank, copy PARTNO
        final_df["MANFPART"] = final_df["MANFPART"].fillna(final_df["PARTNO"])
        final_df.loc[final_df["MANFPART"].astype(str).str.strip() == "", "MANFPART"] = final_df["PARTNO"]

        # UNIT PRICE calculation
        final_df["UNIT PRICE"] = final_df["NETVALUE"] / final_df["QUANTITY"]

        # Build final columns in order
        final_df = final_df[[
            "CartonNo",  # CARTONNO
            "Brand",
            "PARTNO",
            "PARTDESC",
            "QUANTITY",
            "REF1",
            "WEIGHT",
            "MANFPART",
            "CRTNWEIGHT",
            "UNIT PRICE",
            "NETVALUE"
        ]]

        # Rename columns to match final packaging list
        final_df = final_df.rename(columns={
            "CartonNo": "CARTONNO",
            "PARTDESC": "PART DESC",
            "QUANTITY": "QTY",
            "CRTNWEIGHT": "CRTN WEIGHT",
            "NETVALUE": "AMOUNT AED"
        })

        # Add fixed columns
        final_df.insert(8, "HSCODE", "87089900")
        final_df.insert(9, "COO", "")
        final_df.insert(10, "ORDER NUMBER", "")
        final_df.insert(11, "REFERENCES", "")

        # Add SL.NO
        final_df.insert(0, "SL.NO", range(1, len(final_df) + 1))

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
