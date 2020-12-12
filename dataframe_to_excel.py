from typing import Sequence
import pandas as pd
import xlsxwriter


class SheetDataFrame:
    def __init__(self, df: pd.DataFrame, sheet_name: str, first_row_header: str = None):
        self.df = df
        self.sheet_name = sheet_name
        self.first_row_header = first_row_header

    def get_col_widths(self, index: bool):
        # First we find the maximum length of the index column
        idx_max = [max([len(str(s)) for s in self.df.index.values] + [len(str(self.df.index.name))])]
        # Then, we concatenate this to the max of the lengths of column name and its values for each column
        columns_width = [max([len(str(s)) for s in self.df[col].values] + [len(col)]) for col in self.df.columns]
        return idx_max + columns_width if index else columns_width


class DataFrameToExcel:
    START_ROW_IF_SHEET_HEADER_EXISTS = 2
    BORDER_WIDTH = 1

    def __init__(
            self,
            file_path: str,
            sheet_dataframes: Sequence[SheetDataFrame],
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
        self.sheet_dataframes = sheet_dataframes
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
        }
        self.header_format_config = {
            'bold': header_bold,
            'text_wrap': self.header_text_wrap,
            'fg_color': header_color,
            'align': self.horizontal_align,
            'valign': vertical_align
        }
        self.borders_format_config = {
            'bottom': self.BORDER_WIDTH,
            'top': self.BORDER_WIDTH,
            'left': self.BORDER_WIDTH,
            'right': self.BORDER_WIDTH
        }

    def write(self) -> None:
        with pd.ExcelWriter(self.file_path, **self.writer_config) as writer:
            workbook = writer.book
            header_format = workbook.add_format(self.header_format_config)
            border_format = workbook.add_format(self.borders_format_config)
            first_row_format = workbook.add_format(self.first_row_format_config)
            for sheet_df in self.sheet_dataframes:
                sheet_name = sheet_df.sheet_name
                df = sheet_df.df
                first_row_header = sheet_df.first_row_header
                data_start_row = 1 if self.with_header and self.styled_header else 0
                if first_row_header:
                    data_start_row += self.START_ROW_IF_SHEET_HEADER_EXISTS
                self.df_to_excel_config['startrow'] = data_start_row
                styled_df = df.style.set_properties(**{'text-align': self.horizontal_align})
                styled_df.to_excel(writer, sheet_name=sheet_name, **self.df_to_excel_config)
                worksheet = writer.sheets[sheet_name]
                if self.with_header and self.styled_header:
                    for col_num, value in enumerate(df.columns.values):
                        if self.with_index:
                            col_num += 1
                        worksheet.write(data_start_row - 1, col_num, value, header_format)
                if first_row_header:
                    worksheet.write(0, 0, first_row_header, first_row_format)
                first_row = data_start_row
                if self.with_header and self.styled_header:
                    first_row -= 1
                last_row = first_row + len(df)
                if not self.with_header:
                    last_row -= 1
                last_col = len(df.columns)
                if not self.with_index:
                    last_col -= 1
                xl_range = xlsxwriter.utility.xl_range(first_row, 0, last_row, last_col)
                worksheet.conditional_format(xl_range, {'type': 'no_errors', 'format': border_format})
                worksheet.right_to_left()
                for i, width in enumerate(sheet_df.get_col_widths(self.with_index)):
                    worksheet.set_column(i, i, width)
