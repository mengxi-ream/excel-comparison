# @author Kuiliang Zhang (Xi Meng)
# @create date 2022-01-07 16:38:13
# @modify date 2022-01-12 15:34:07
# @desc compare differences between excel files

import pandas as pd
import numpy as np


def load_file(order):
    file_path = input(f"Please specify the path of excel file ({order}/2): ")
    file = pd.read_excel(file_path, sheet_name=None, header=None, dtype=str)
    return file


def enlarge_df_to_same_shape(origin_df_1, origin_df_2):
    df_1 = origin_df_1.copy(deep=True)
    df_2 = origin_df_2.copy(deep=True)
    max_row_num = max(df_1.shape[0], df_2.shape[0])
    while df_1.shape[0] < max_row_num:
        df_1.loc[df_1.shape[0]] = np.nan
    while df_2.shape[0] < max_row_num:
        df_2.loc[df_2.shape[0]] = np.nan

    max_col_num = max(df_1.shape[1], df_2.shape[1])
    while df_1.shape[1] < max_col_num:
        df_1[df_1.shape[1]] = np.nan
    while df_2.shape[1] < max_col_num:
        df_2[df_2.shape[1]] = np.nan

    return df_1, df_2


def main():
    file_1 = load_file(1)
    file_2 = load_file(2)

    # get sheet names of the excel files in key_list
    key_list = list(file_1.keys())
    for key in file_2:
        if key not in key_list:
            key_list.append(key)

    # define writer and related format
    writer = pd.ExcelWriter("file_diff.xlsx", engine="xlsxwriter")
    workbook = writer.book
    grey_fmt = workbook.add_format({"font_color": "#E0E0E0"})
    highlight_fmt = workbook.add_format(
        {"font_color": "#FF0000", "bg_color": "#B1B3B3"}
    )

    for sheet_name in key_list:
        if sheet_name not in file_1.keys() or sheet_name not in file_2.keys():
            print(f'Sheet "{sheet_name}" does not exist in both files')
            continue

        # enlarge df_1 and df_2 to the same shape
        df_1, df_2 = enlarge_df_to_same_shape(file_1[sheet_name], file_2[sheet_name])
        df_diff = df_1.copy(deep=True)

        # compare values
        for row in range(df_diff.shape[0]):
            for col in range(df_diff.shape[1]):
                value_1 = df_1.iloc[row, col]
                value_2 = df_2.iloc[row, col]
                if pd.isnull(value_1) and pd.isnull(value_2):
                    continue
                if value_1 == value_2:
                    continue
                if pd.isnull(value_1):
                    value_1 = "NaN"
                if pd.isnull(value_2):
                    value_2 = "NaN"
                df_diff.iloc[row, col] = f"{value_1} → {value_2}"

        # write df_diff
        df_diff.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        worksheet = writer.sheets[sheet_name]

        ## highlight changed cells
        worksheet.conditional_format(
            "A1:ZZ1000",
            {
                "type": "text",
                "criteria": "containing",
                "value": "→",
                "format": highlight_fmt,
            },
        )
        ## highlight unchanged cells
        worksheet.conditional_format(
            "A1:ZZ1000",
            {
                "type": "text",
                "criteria": "not containing",
                "value": "→",
                "format": grey_fmt,
            },
        )

    writer.save()


if __name__ == "__main__":
    main()
