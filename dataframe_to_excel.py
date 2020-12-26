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


class MashovReportsToExcel:
    SCHOOLS = list(range(100151, 100158))
    DESTINATION_FOLDER_NAME = 'דוחות משוב'

    def __init__(self, heb_year: str, class_codes: Sequence[str], username: str, password: str,
                 destination_folder_path: str, from_date: date = None, to_date: date = None):
        self.class_codes = class_codes
        self.heb_year = heb_year
        self.destination_folder_path = os.path.join(destination_folder_path, self.DESTINATION_FOLDER_NAME)
        if os.path.exists(self.destination_folder_path):
            try:
                shutil.rmtree(self.destination_folder_path)
            except:
                raise OSError('נא לסגור את הדוחות הפתוחים!')
        os.makedirs(self.destination_folder_path)
        self.report_makers_for_class = dict()
        for class_code in self.class_codes:
            report_maker = ReportMaker(self.SCHOOLS, heb_year, class_code, username, password)
            if not from_date:
                from_date = report_maker.first_school_year_date
            if not to_date:
                to_date = report_maker.last_school_year_date
            self.from_date = from_date
            self.to_date = to_date
            report_maker.fetch_data_from_server(from_date, to_date)
            self.report_makers_for_class[class_code] = report_maker

    def write_raw_behavior_report(self, from_date: date, to_date: date) -> None:
        date_range_str = f'{from_date.strftime("%d.%m.%Y")}-{to_date.strftime("%d.%m.%Y")}'
        for class_code, report_maker in self.report_makers_for_class.items():
            file_name = f'דוח התנהגות גולמי שכבה {class_code} {date_range_str}.xlsx'
            file_path = os.path.join(self.destination_folder_path, file_name)
            sheets = []
            school_behavior_data = report_maker.create_raw_behavior_report_by_schools(from_date, to_date)
            for school_name, behavior_df in school_behavior_data.items():
                sheet_name = f'{school_name} {class_code}'
                header = f'דוח התנהגות גולמי {school_name}'
                sheet = Sheet(sheet_name, [SheetDataFrame(behavior_df, header)])
                sheets.append(sheet)
            excel_writer = DataFrameToExcel(file_path=file_path, sheets=sheets)
            excel_writer.write()

    def write_summary_report(self, from_date: date, to_date: date) -> None:
        date_range_str = f'{from_date.strftime("%d.%m.%Y")}-{to_date.strftime("%d.%m.%Y")}'
        for class_code, report_maker in self.report_makers_for_class.items():
            file_name = f'דוח סיכום שכבה {class_code} {date_range_str}.xlsx'
            file_path = os.path.join(self.destination_folder_path, file_name)
            sheets = []
            municipal_average_presence_df = report_maker.create_municipal_average_presence_report(from_date, to_date)
            municipal_average_presence_sheet_name = 'ממוצע נוכחות עירוני'
            municipal_average_presence_header = f'{municipal_average_presence_sheet_name} - שכבה {class_code}'
            municipal_average_presence_sheet = Sheet(
                municipal_average_presence_sheet_name,
                [SheetDataFrame(municipal_average_presence_df, municipal_average_presence_header)]
            )
            sheets.append(municipal_average_presence_sheet)
            schools_data = report_maker.create_summary_report_by_schools(from_date, to_date)
            for school_name, presence_df in schools_data.items():
                presence_sheet_name = f'{school_name} {class_code}'
                presence_header = f'דוח סיכום {school_name}'
                presence_sheet = Sheet(presence_sheet_name, [SheetDataFrame(presence_df, presence_header)])
                sheets.append(presence_sheet)
            excel_writer = DataFrameToExcel(file_path=file_path, sheets=sheets)
            excel_writer.write()

    def write_mashov_report(self, from_date: date, to_date: date) -> None:
        date_range_str = f'{from_date.strftime("%d.%m.%Y")}-{to_date.strftime("%d.%m.%Y")}'
        for class_code, report_maker in self.report_makers_for_class.items():
            file_name = f'דוח משוב שכבה {class_code} {date_range_str}.xlsx'
            file_path = os.path.join(self.destination_folder_path, file_name)
            sheets = []
            presence_summary_df = report_maker.create_presence_summary_report(from_date, to_date)
            presence_summary_sheet_name = f'סיכום נוכחות {class_code}'
            presence_summary_header = f'דוח סיכום נוכחות שכבה {class_code}'
            presence_summary_sheet = Sheet(
                presence_summary_sheet_name, [SheetDataFrame(presence_summary_df, presence_summary_header)]
            )
            sheets.append(presence_summary_sheet)
            events_without_remarks_df = report_maker.create_events_without_remarks_report(from_date, to_date)
            events_without_remarks_sheet_name = f'אירועים בלי הערות {class_code}'
            events_without_remarks_header = f'דוח אירועים בלי הערות שכבה {class_code}'
            events_without_remarks_sheet = Sheet(
                events_without_remarks_sheet_name,
                [SheetDataFrame(events_without_remarks_df, events_without_remarks_header)]
            )
            sheets.append(events_without_remarks_sheet)
            middle_week_lessons_df = report_maker.create_middle_week_lessons_report(from_date, to_date)
            middle_week_lessons_sheet_name = f'לימודים שהתקיימו באמצע השבוע {class_code}'
            middle_week_lessons_header = f'דוח לימודים שהתקיימו באמצע השבוע שכבה {class_code}'
            middle_week_lessons_sheet = Sheet(
                middle_week_lessons_sheet_name,
                [SheetDataFrame(middle_week_lessons_df, middle_week_lessons_header)]
            )
            sheets.append(middle_week_lessons_sheet)
            presence_distribution_df = report_maker.create_presence_distribution_report(from_date, to_date)
            presence_distribution_sheet_name = f'התפלגות נוכחות {class_code}'
            presence_distribution_header = f'דוח התפלגות נוכחות שכבה {class_code}'
            presence_distribution_sheet = Sheet(
                presence_distribution_sheet_name,
                [SheetDataFrame(presence_distribution_df, presence_distribution_header)]
            )
            sheets.append(presence_distribution_sheet)
            excel_writer = DataFrameToExcel(file_path=file_path, sheets=sheets)
            excel_writer.write()

    def write_periodical_report(self, from_date: date, to_date: date) -> None:
        date_range_str = f'{from_date.strftime("%d.%m.%Y")}-{to_date.strftime("%d.%m.%Y")}'
        for class_code, report_maker in self.report_makers_for_class.items():
            file_name = f'דוח תקופתי שכבה {class_code} {date_range_str}.xlsx'
            file_path = os.path.join(self.destination_folder_path, file_name)
            sheets = []
            presence_by_schools = report_maker.create_presence_report_by_schools(from_date, to_date)
            dfs_in_sheet = []
            presence_schools_sheet_name = f'ועדת היגוי לכל בי"ס {class_code}'
            for school, school_df in presence_by_schools.items():
                curr_df = SheetDataFrame(school_df, f'דוח תקופתי ועדת היגוי עירונית שכבה {class_code} {school}')
                dfs_in_sheet.append(curr_df)
            schools_sheet = Sheet(presence_schools_sheet_name, dfs_in_sheet)
            sheets.append(schools_sheet)
            municipal_presence_by_level = report_maker.create_municipal_presence_report_by_levels(from_date, to_date)
            dfs_in_sheet = []
            municipal_presence_sheet_name = f'ועדת היגוי עירונית {class_code}'
            for level, level_df in municipal_presence_by_level.items():
                curr_df = SheetDataFrame(level_df, f'דוח תקופתי ועדת היגוי עירונית שכבה {class_code} {level}')
                dfs_in_sheet.append(curr_df)
            municipal_presence_sheet = Sheet(municipal_presence_sheet_name, dfs_in_sheet)
            sheets.append(municipal_presence_sheet)
            grades_colors_by_level = report_maker.create_grades_colors_report_by_levels()
            grades_sheet_name = f'שותפים - ציונים {class_code}'
            dfs_in_sheet = []
            for level, level_df in grades_colors_by_level.items():
                curr_df = SheetDataFrame(level_df, f'דוח תקופתי ציונים סמסטריאלים שכבה {class_code} {level}')
                dfs_in_sheet.append(curr_df)
            grades_sheet = Sheet(grades_sheet_name, dfs_in_sheet)
            sheets.append(grades_sheet)
            monthly_presence_by_level = report_maker.create_presence_report_of_month_by_levels(from_date.month,
                                                                                               to_date.month,
                                                                                               from_date.year,
                                                                                               to_date.year)
            dfs_in_sheet = []
            monthly_presence_sheet_name = f'שותפים - נוכחות {class_code}'
            for level, level_df in monthly_presence_by_level.items():
                curr_df = SheetDataFrame(level_df, f'דוח תקופתי נוכחות ממוצעת לפי חודש שכבה {class_code} {level}')
                dfs_in_sheet.append(curr_df)
            monthly_presence_sheet = Sheet(monthly_presence_sheet_name, dfs_in_sheet)
            sheets.append(monthly_presence_sheet)
            excel_writer = DataFrameToExcel(file_path=file_path, sheets=sheets)
            excel_writer.write()
