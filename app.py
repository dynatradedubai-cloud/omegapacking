
import pandas as pd


def generate_final_packing_list(order_file, packing_file, output_file):
    # Load Excel files
    order_df = pd.read_excel(order_file, sheet_name=0)
    packing_df = pd.read_excel(packing_file, sheet_name=0)

    # Remove zero quantity rows
    packing_df = packing_df[packing_df["QUANTITY"] > 0].copy()

    # Merge Order list for fallback values
    merged_df = packing_df.merge(
        order_df[["PARTNO", "Brand", "MANFPART", "Price"]],
        on="PARTNO",
        how="left",
        suffixes=("", "_order")
    )

    # Fallback logic
    merged_df["Brand"] = merged_df["Brand"].fillna(merged_df["Brand_order"])
    merged_df["MANFPART"] = merged_df["MANFPART"].fillna(merged_df["MANFPART_order"])

    # Calculate UNIT PRICE
    merged_df["UNIT PRICE"] = merged_df["NETVALUE"] / merged_df["QUANTITY"]

    # If UNIT PRICE missing, fallback to Order price
    merged_df["UNIT PRICE"] = merged_df["UNIT PRICE"].fillna(merged_df["Price"])

    # Build final dataframe
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

    # Save output
    final_df.to_excel(output_file, sheet_name="Sheet1", index=False)

    print(f"Final packaging list generated: {output_file}")


if __name__ == "__main__":
    order_file = "Order list.xlsx"
    packing_file = "Packing list.xlsx"
    output_file = "Final packaging list.xlsx"

    generate_final_packing_list(order_file, packing_file, output_file)
