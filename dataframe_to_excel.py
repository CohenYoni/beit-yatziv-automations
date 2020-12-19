from reports_maker import ReportMaker
from typing import Sequence
from datetime import date
import pandas as pd
import xlsxwriter
import shutil
import os


class SheetDataFrame:
    def __init__(self, df: pd.DataFrame, first_row_header: str = None):
        self.df = df
        self.first_row_header = first_row_header

    def get_col_widths(self, index: bool):
        df = self.df.copy()
        has_multi_idx = type(df.columns) == pd.MultiIndex
        if has_multi_idx:
            df.columns = df.columns.droplevel(0)
        # First we find the maximum length of the index column
        idx_max = [max([len(str(s)) for s in df.index.values] + [len(str(df.index.name))])]
        # Then, we concatenate this to the max of the lengths of column name and its values for each column
        columns_width = [max([len(str(s)) for s in df[col].values] + [len(col)]) for col in df.columns]
        return idx_max + columns_width if index else columns_width


class Sheet:
    def __init__(self, sheet_name: str, sheet_dataframes: Sequence[SheetDataFrame]):
        self.sheet_name = sheet_name
        self.sheet_dataframes = sheet_dataframes


class DataFrameToExcel:
    START_ROW_IF_SHEET_HEADER_EXISTS = 2
    BORDER_WIDTH = 1

    @staticmethod
    def non_unique_col_idx_handler(df: pd.DataFrame, style_properties: dict) -> pd.DataFrame:
        is_df_unique = True
        org_idx = df.index
        org_cols = df.columns
        if not df.index.is_unique or not df.columns.is_unique:
            is_df_unique = False
            df.columns = [f'{col}{i}' for i, col in enumerate(df.columns)]
            df.reset_index(drop=True, inplace=True)
        styled_df = df.style.set_properties(**style_properties)
        if not is_df_unique:
            df.index = org_idx
            df.columns = org_cols
            styled_df = df.copy()
        return styled_df

    def __init__(
            self,
            file_path: str,
            sheets: Sequence[Sheet],
            date_format: str = 'dd/mm/yyyy',
            datetime_format: str = 'dd/mm/yyyy',
            with_index: bool = False,
            with_header: bool = True,
            float_format: str = '%.2f',
            header_bold: bool = True,
            header_text_wrap: bool = False,
            header_color: str = '#B4C6E7',
            horizontal_align: str = 'center',
            vertical_align: str = 'center',
            with_all_borders: bool = True,
            styled_header: bool = True
    ):
        self.file_path = file_path
        self.sheets = sheets
        self.with_index = with_index
        self.styled_header = styled_header
        self.with_header = with_header
        self.header_text_wrap = header_text_wrap
        self.with_all_borders = with_all_borders
        self.horizontal_align = horizontal_align
        self.writer_config = {
            'date_format': date_format,
            'datetime_format': datetime_format
        }
        self.df_to_excel_config = {
            'index': self.with_index,
            'header': self.with_header and not self.styled_header,
            'startrow': 1 if self.with_header and self.styled_header else 0,
            'float_format': float_format
        }
        self.first_row_format_config = {
            'bold': header_bold,
            'reading_order': 2
        }
        self.header_format_config = {
            'bold': header_bold,
            'text_wrap': self.header_text_wrap,
            'fg_color': header_color,
            'align': self.horizontal_align,
            'valign': vertical_align,
            'reading_order': 2
        }
        self.borders_format_config = {
            'bottom': self.BORDER_WIDTH,
            'top': self.BORDER_WIDTH,
            'left': self.BORDER_WIDTH,
            'right': self.BORDER_WIDTH,
            'align': self.horizontal_align
        }

    def get_multi_column_first_row(self, df: pd.DataFrame) -> list:
        first_col_names = [col[0] for col in df.columns]
        new_col = [first_col_names[0], ]
        for i in range(1, len(first_col_names)):
            if first_col_names[i] == first_col_names[i - 1]:
                new_col.append('')
            else:
                new_col.append(first_col_names[i])
        if self.with_index:
            new_col.insert(0, '')
        return new_col

    def get_multi_column_range_to_merge(self, multi_column_first_row: list) -> dict:
        start = 1 if self.with_index else 0
        to_merge = {start: 0}
        if self.with_index:
            to_merge[0] = 0
        for i, col_name in enumerate(multi_column_first_row[1:], 1):
            if col_name == '':
                to_merge[start] += 1
            else:
                start = i
                to_merge[start] = start
        return to_merge

    def write(self) -> None:
        with pd.ExcelWriter(self.file_path, **self.writer_config) as writer:
            workbook = writer.book
            header_format = workbook.add_format(self.header_format_config)
            border_format = workbook.add_format(self.borders_format_config)
            first_row_format = workbook.add_format(self.first_row_format_config)
            for sheet in self.sheets:
                sheet_name = sheet.sheet_name
                current_start_row = 0
                for sheet_df in sheet.sheet_dataframes:
                    df = sheet_df.df.copy()
                    first_row_header = sheet_df.first_row_header
                    data_start_row = current_start_row + 1 if self.with_header and self.styled_header else current_start_row
                    if first_row_header:
                        data_start_row += self.START_ROW_IF_SHEET_HEADER_EXISTS
                    has_multi_idx = type(df.columns) == pd.MultiIndex
                    first_row_multi_index = []
                    col_range_to_merge = dict()
                    if has_multi_idx:
                        data_start_row += 1
                        first_row_multi_index = self.get_multi_column_first_row(df)
                        df.columns = df.columns.droplevel(0)
                        col_range_to_merge = self.get_multi_column_range_to_merge(first_row_multi_index)
                    self.df_to_excel_config['startrow'] = data_start_row
                    # styled_df = self.non_unique_col_idx_handler(df, {'text-align': self.horizontal_align})
                    df.to_excel(writer, sheet_name=sheet_name, **self.df_to_excel_config)
                    worksheet = writer.sheets[sheet_name]
                    if self.with_header and self.styled_header:
                        for col_num, value in enumerate(df.columns.values):
                            if self.with_index:
                                col_num += 1
                            worksheet.write(data_start_row - 1, col_num, value, header_format)
                    if first_row_header:
                        worksheet.write(current_start_row, 0, first_row_header, first_row_format)
                    first_row = data_start_row
                    if self.with_header and self.styled_header:
                        first_row -= 1
                    last_row = first_row + len(df)
                    if not self.with_header:
                        last_row -= 1
                    last_col = len(df.columns)
                    if not self.with_index:
                        last_col -= 1
                    if has_multi_idx:
                        first_row -= 1
                        for first_col_idx, last_col_idx in col_range_to_merge.items():
                            if first_col_idx < last_col_idx:
                                data = first_row_multi_index[first_col_idx]
                                worksheet.merge_range(first_row, first_col_idx, first_row, last_col_idx, data,
                                                      header_format)
                    xl_range = xlsxwriter.utility.xl_range(first_row, 0, last_row, last_col)
                    worksheet.conditional_format(xl_range, {'type': 'no_errors', 'format': border_format})
                    worksheet.right_to_left()
                    for i, width in enumerate(sheet_df.get_col_widths(self.with_index)):
                        worksheet.set_column(i, i, width)
                    current_start_row += len(df) + 2
                    if first_row_header:
                        current_start_row += self.START_ROW_IF_SHEET_HEADER_EXISTS
