import pandas as pd

def generate_final_packing_list(order_file, packing_file, output_file):
    # Load sheets
    order_df = pd.read_excel(order_file, sheet_name=0)
    packing_df = pd.read_excel(packing_file, sheet_name=0)

    # Remove zero quantity rows
    packing_df = packing_df[packing_df["QUANTITY"] > 0].copy()

    # Calculate unit price from NETVALUE / QUANTITY
    packing_df["UNIT PRICE"] = packing_df["NETVALUE"] / packing_df["QUANTITY"]

    # Ensure Brand column comes strictly from Order list
    if "Brand" in order_df.columns:
        if len(order_df) != len(packing_df):
            raise ValueError("Row count mismatch: Order list and Packing list must have the same number of rows.")
        packing_df["Brand"] = order_df["Brand"]
    else:
        raise ValueError("Order list.xlsx must contain 'Brand' column.")

    # MANFPART: if blank, copy PARTNO
    packing_df["MANFPART"] = packing_df["MANFPART"].fillna(packing_df["PARTNO"])
    packing_df.loc[packing_df["MANFPART"].astype(str).str.strip() == "", "MANFPART"] = packing_df["PARTNO"]

    # Build final DataFrame using required column order
    final_df = pd.DataFrame({
        "SL.NO": range(1, len(packing_df) + 1),
        "CARTONNO": packing_df["CARTONNO"],
        "Brand": packing_df["Brand"],  # strictly from Order list
        "PARTNO": packing_df["PARTNO"],
        "PART DESC": packing_df["PARTDESC"],
        "COO": "",  # left blank
        "QTY": packing_df["QUANTITY"],
        "REF1": packing_df["REF1"],
        "WEIGHT": packing_df["WEIGHT"],
        "HSCODE": "87089900",
        "UNIT PRICE": packing_df["UNIT PRICE"],
        "AMOUNT AED": packing_df["NETVALUE"],
        "MANFPART": packing_df["MANFPART"],
        "CRTN WEIGHT": packing_df["CRTNWEIGHT"],
        "ORDER NUMBER": "",  # left blank
        "REFERENCES": ""  # left blank
    })

    # Save output
    final_df.to_excel(output_file, index=False)
    print(f"Final packaging list generated: {output_file}")


if __name__ == "__main__":
    # File names
    order_file = "Order list.xlsx"
    packing_file = "Packing list.xlsx"
    output_file = "Final packaging list.xlsx"

    generate_final_packing_list(order_file, packing_file, output_file)


