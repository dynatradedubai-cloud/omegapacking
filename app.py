import streamlit as st
import pandas as pd

st.title("Final Packaging List Generator")

uploaded_order = st.file_uploader("Upload Order list.xlsx", type=["xlsx"])
uploaded_packing = st.file_uploader("Upload Packing list.xlsx", type=["xlsx"])

if uploaded_order and uploaded_packing:

    order_df = pd.read_excel(uploaded_order)
    packing_df = pd.read_excel(uploaded_packing)

    packing_df = packing_df[packing_df["QUANTITY"] > 0].copy()

    merged_df = packing_df.merge(
        order_df[["PARTNO", "Brand", "MANFPART", "Price"]],
        on="PARTNO",
        how="left",
        suffixes=("", "_order")
    )

    merged_df["Brand"] = merged_df["Brand"].fillna(merged_df["Brand_order"])
    merged_df["MANFPART"] = merged_df["MANFPART"].fillna(merged_df["MANFPART_order"])

    merged_df["UNIT PRICE"] = merged_df["NETVALUE"] / merged_df["QUANTITY"]
    merged_df["UNIT PRICE"] = merged_df["UNIT PRICE"].fillna(merged_df["Price"])

    final_df = pd.DataFrame({
        "SL.NO": range(1, len(merged_df) + 1),
        "CARTONNO": merged_df["CARTONNO"],
        "Brand": merged_df["Brand"],
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

    st.success("File Generated Successfully")

    st.download_button(
        label="Download Final Packaging List",
        data=final_df.to_excel(index=False),
        file_name="Final packaging list.xlsx"
    )
