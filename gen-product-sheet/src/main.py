import pandas as pd
from configs.constants import OUTPUT_COLUMNS


def read_xlsx_file(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    return df


def gen_empty_product_row():
    new_row = {col["name"]: "" for col in OUTPUT_COLUMNS}
    return new_row.copy()


def gen_product_sheet(df: pd.DataFrame):
    result_rows = []

    for index, row in df.iterrows():
        product_name, colors, price, discount, short_description, category = (
            row["商品名稱"],
            row["顏色"].split(", "),
            row["定價"],
            row["折數"],
            row["簡短內容說明"],
            row["分類"]
        )

        ## handle parent row
        new_row_parent = gen_empty_product_row()
        new_row_parent["類型"] = "variable" if len(colors) > 1 else "simple"
        new_row_parent["貨號"] = f"L{index}"
        new_row_parent["名稱"] = product_name
        new_row_parent["目錄的可見度"] = "visible"
        new_row_parent["簡短內容說明"] = short_description
        new_row_parent["特價"] = round(price * discount)
        new_row_parent["原價"] = price
        new_row_parent["分類"] = category
        new_row_parent["位置"] = 0
        new_row_parent["屬性 1 名稱"] = "顏色"
        new_row_parent["屬性 1 值"] = ", ".join(colors)
        new_row_parent["屬性 1 可見"] = "1"
        new_row_parent["屬性 1 全域"] = "0"
        new_row_parent["屬性 1 預設"] = colors[0]
        result_rows.append(new_row_parent)

        ## handle child row if any
        for i, color in enumerate(colors):
            new_row_child = gen_empty_product_row()
            new_row_child["類型"] = "variation"
            new_row_child["貨號"] = f"L{index}-{i+1}"
            new_row_child["名稱"] = f"{product_name} - {color}"
            new_row_child["目錄的可見度"] = "visible"
            new_row_child["特價"] = round(price * discount)
            new_row_child["原價"] = price
            new_row_child["上層"] = f"L{index}"
            new_row_child["位置"] = i + 1
            new_row_child["屬性 1 名稱"] = "顏色"
            new_row_child["屬性 1 值"] = color
            new_row_child["屬性 1 全域"] = "0"

            result_rows.append(new_row_child)

    out_df = pd.DataFrame(result_rows)
    return out_df


def main():
    ## read input xlsx file
    file_path = "src/data"
    file_name = "LELO_20240916"
    input_file_name = f"{file_path}/{file_name}.xlsx"
    output_file_name = f"{file_path}/{file_name}_output.csv"
    df = read_xlsx_file(input_file_name)

    ## generate product sheet
    out_df = gen_product_sheet(df)

    ## write to output xlsx file
    ordered_columns = [
        col["name"] for col in sorted(OUTPUT_COLUMNS, key=lambda x: x["position"])
    ]
    out_df = out_df[ordered_columns]
    out_df.to_csv(output_file_name, index=False)


if __name__ == "__main__":
    main()
