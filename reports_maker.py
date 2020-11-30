from data_server import MashovServer, School
from datetime import datetime, date
from typing import Dict, Sequence
import pandas as pd
import numpy as np
import calendar


MONTHS_IN_HEBREW = {
    1: 'ינואר',
    2: 'פברואר',
    3: 'מרץ',
    4: 'אפריל',
    5: 'מאי',
    6: 'יוני',
    7: 'יולי',
    8: 'אוגוסט',
    9: 'ספטמבר',
    10: 'אוקטובר',
    11: 'נובמבר',
    12: 'דצמבר',
}


class SchoolData(School):
    def __init__(self, school_id: int, name: str, class_code: str):
        super().__init__(school_id, name)
        self.class_code = class_code
        self._behavior_report = None
        self._phonebook = None
        self._grades_report = None
        self._organic_teachers: Dict[int, str] = dict()
        self._practitioners: Dict[int, str] = dict()
        self._levels: Dict[int, str] = dict()
        self._num_of_students: Dict[int, int] = dict()
        self._num_of_active_classes = 0

    @property
    def behavior_report(self) -> pd.DataFrame:
        return self._behavior_report

    @behavior_report.setter
    def behavior_report(self, behavior_report: pd.DataFrame) -> None:
        self._behavior_report = behavior_report

    @property
    def phonebook(self) -> pd.DataFrame:
        return self._phonebook

    @phonebook.setter
    def phonebook(self, phonebook: pd.DataFrame) -> None:
        self._phonebook = phonebook

    @property
    def grades_report(self) -> pd.DataFrame:
        return self._grades_report

    @grades_report.setter
    def grades_report(self, grades_report: pd.DataFrame) -> None:
        self._grades_report = grades_report

    @property
    def num_of_active_classes(self):
        return self._num_of_active_classes

    @num_of_active_classes.setter
    def num_of_active_classes(self, num_of_active_classes: int) -> None:
        if num_of_active_classes < 0:
            num_of_active_classes = 0
        self._num_of_active_classes = num_of_active_classes

    def set_organic_teacher(self, class_num: int, organic_teacher_name: str):
        self._organic_teachers[class_num] = organic_teacher_name

    def set_practitioner(self, class_num: int, practitioner_name: str):
        self._practitioners[class_num] = practitioner_name

    def set_level(self, class_num: int, level: str):
        self._levels[class_num] = level

    def set_num_of_students(self, class_num: int, num_of_students: int):
        self._num_of_students[class_num] = num_of_students

    def get_organic_teacher(self, class_num: int) -> str:
        return self._organic_teachers.get(class_num, '')

    def get_practitioner(self, class_num) -> str:
        return self._practitioners.get(class_num, '')

    def get_level(self, class_num: int) -> str:
        return self._levels.get(class_num, '')

    def get_num_of_students(self, class_num: int) -> int:
        return self._num_of_students.get(class_num, 0)


class ReportMaker:
    class LessonEvents:
        PRESENCE = 'נוכחות'
        MISSING = 'חיסור'
        REINFORCEMENT = 'חיזוק חיובי'
        LATE = 'איחור'
        DISTURB = 'הפרעה'

    HEB_WEEKDAYS = ['שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת', 'ראשון']
    HEB_DAYS_MAPPER = {idx: day for idx, day in enumerate(HEB_WEEKDAYS, 0)}
    NO_REMARKS = 'ללא הערות'
    FAIL_GRADE_THRESHOLD = 85
    NEGATIVE_GRADE_THRESHOLD = 56
    DATE_FORMAT = '%d/%m/%Y'

    @staticmethod
    def count_events(events_df, event_type):
        event_filter = events_df['event_type'] == event_type
        result = events_df.loc[event_filter, 'event_type']
        num_of_result = result.count()
        return num_of_result

    @staticmethod
    def sort_datetime_columns_names(df: pd.DataFrame, non_datetime_names: Sequence, datetime_format: str):
        non_datetime_names_columns = df[non_datetime_names]
        datetime_names_columns = df.drop(non_datetime_names_columns, axis=1)
        datetime_names_columns.columns = [datetime.strptime(d, datetime_format) for d in datetime_names_columns.columns]
        datetime_names_columns = datetime_names_columns.sort_index(axis=1)
        datetime_names_columns.columns = [datetime.strftime(d, datetime_format) for d in datetime_names_columns.columns]
        return pd.concat([non_datetime_names_columns, datetime_names_columns], axis=1)

    @staticmethod
    def get_date_range_of_week(year: int, week_number: int) -> str:
        week_first_date = date.fromisocalendar(year, week_number, 1)
        week_last_date = date.fromisocalendar(year, week_number, 7)
        rng = f'{week_first_date.strftime(ReportMaker.DATE_FORMAT)}-{week_last_date.strftime(ReportMaker.DATE_FORMAT)}'
        return rng

    def __init__(self, schools_ids: list, heb_year: str, class_code: str, username: str, password: str):
        self.schools_data: Dict[int, SchoolData] = {_id: None for _id in schools_ids}
        self.heb_year = heb_year
        self.class_code = class_code
        self.username = username
        self.password = password
        self.from_date = None
        self.to_date = None

    def fetch_data_from_server(self, from_date: date, to_date: date) -> None:
        assert from_date <= to_date, 'From date must be less than to date'
        self.from_date = from_date
        self.to_date = to_date
        for school_id in self.schools_data.keys():
            server = MashovServer(school_id=school_id, school_year=self.heb_year)
            try:
                server.login(username=self.username, password=self.password)
                behavior_report = server.get_behavior_report_by_dates(from_date=from_date,
                                                                      to_date=to_date,
                                                                      class_code=self.class_code)
                phonebook = server.get_students_phonebook(class_code=self.class_code)
                grades_report = server.get_grades_report(from_date=from_date,
                                                         to_date=to_date,
                                                         class_code=self.class_code)
                school_class_data = SchoolData(school_id, server.school.name, self.class_code)
                school_class_data.behavior_report = behavior_report
                school_class_data.phonebook = phonebook
                school_class_data.grades_report = grades_report
                school_class_data.num_of_active_classes = server.get_num_of_active_classes(self.class_code)
                for class_num in range(1, school_class_data.num_of_active_classes + 1):
                    organic_teacher_name = server.get_organic_teacher_name(self.class_code, class_num)
                    practitioner_name = server.get_class_practitioner(self.class_code, class_num)
                    class_level = server.get_class_level(self.class_code, class_num)
                    school_class_data.set_organic_teacher(class_num, organic_teacher_name)
                    school_class_data.set_practitioner(class_num, practitioner_name)
                    school_class_data.set_level(class_num, class_level)
                archives_class_num = school_class_data.num_of_active_classes + 1
                archives_class_level = server.get_class_level(self.class_code, archives_class_num)
                school_class_data.set_level(archives_class_num, archives_class_level)
                self.schools_data[school_id] = school_class_data
            except Exception:
                raise
            finally:
                server.logout()
        self.calculate_num_of_students()

    def calculate_num_of_students(self):
        for school_id, school_data in self.schools_data.items():
            for class_num in range(1, school_data.num_of_active_classes + 1):
                class_num_filter = school_data.behavior_report['class_num'] == class_num
                class_num_students = school_data.behavior_report.loc[class_num_filter, 'student_id']
                school_data.set_num_of_students(class_num, class_num_students.nunique())

    def create_presence_summary_report(self) -> pd.DataFrame:
        presence_summary_df = pd.DataFrame(columns=['בית ספר'])
        for school_id, school_data in self.schools_data.items():
            presence_filter = school_data.behavior_report['event_type'] == self.LessonEvents.PRESENCE
            presence_df = school_data.behavior_report.loc[presence_filter, ['lesson_date', 'event_type']]
            presents_grp = presence_df.groupby(['lesson_date'])
            count_grp = presents_grp.agg('count')
            presence_df = pd.DataFrame(columns=count_grp.index.to_list())
            presence_df = presence_df.append(count_grp['event_type'], ignore_index=True)
            presence_df['בית ספר'] = [school_data.name]
            presence_summary_df = pd.concat([presence_summary_df, presence_df], ignore_index=True)
        return presence_summary_df

    def create_events_without_remarks_report(self) -> pd.DataFrame:
        columns = ['שם המורה', 'מקצוע', 'תאריך', 'מספר שיעור', 'שם התלמיד', 'שכבה', 'כיתה', 'סוג האירוע',
                   'הערה מילולית', 'הוצדק ע"י', 'הצדקה', 'בית ספר', 'יום']
        events_without_remarks = pd.DataFrame(columns=columns)
        for school_id, school_data in self.schools_data.items():
            no_remark_events_df = school_data.behavior_report.drop('student_id', axis=1)
            event_filter = no_remark_events_df['event_type'] == self.LessonEvents.MISSING
            remark_filter = no_remark_events_df['remark'].fillna('') == ''
            justification_filter = no_remark_events_df['justification'] == self.NO_REMARKS
            no_remark_events_df = no_remark_events_df.loc[event_filter & remark_filter & justification_filter]
            no_remark_events_df['school_name'] = school_data.name
            no_remark_events_df['day_of_week'] = no_remark_events_df['lesson_date'].dt.weekday.apply(
                lambda week_day: self.HEB_DAYS_MAPPER[week_day])
            no_remark_events_df.columns = columns
            events_without_remarks = pd.concat([events_without_remarks, no_remark_events_df], ignore_index=True)
        return events_without_remarks

    def create_middle_week_lessons_report(self) -> pd.DataFrame:
        columns = ['בית ספר', 'נוכחים', 'חיסורים', 'חיזוקים', 'איחור', 'הפרעה', 'מצבת']
        middle_week_lessons_df = pd.DataFrame(columns=columns)
        for school_id, school_data in self.schools_data.items():
            not_in_saturday_filter = school_data.behavior_report['lesson_date'].dt.weekday != calendar.SATURDAY
            required_columns = ['lesson_date', 'event_type']
            not_in_saturday_df = school_data.behavior_report.loc[not_in_saturday_filter, required_columns]
            num_of_students = not_in_saturday_df['lesson_date'].count()
            num_of_presents = self.count_events(not_in_saturday_df, self.LessonEvents.PRESENCE)
            num_of_missing = self.count_events(not_in_saturday_df, self.LessonEvents.MISSING)
            num_of_reinforcements = self.count_events(not_in_saturday_df, self.LessonEvents.REINFORCEMENT)
            num_of_lateness = self.count_events(not_in_saturday_df, self.LessonEvents.LATE)
            num_of_disturbs = self.count_events(not_in_saturday_df, self.LessonEvents.DISTURB)
            data = [[
                school_data.name,
                num_of_presents,
                num_of_missing,
                num_of_reinforcements,
                num_of_lateness,
                num_of_disturbs,
                num_of_students
            ]]
            curr_df = pd.DataFrame(data, columns=columns)
            middle_week_lessons_df = pd.concat([middle_week_lessons_df, curr_df], ignore_index=True)
        return middle_week_lessons_df

    def assert_dates_in_range(self, from_date: date, to_date: date):
        current_date_range = f'{self.from_date.strftime(self.DATE_FORMAT)} - {self.to_date.strftime(self.DATE_FORMAT)}'
        required_date_range = f'{from_date.strftime(self.DATE_FORMAT)} - {to_date.strftime(self.DATE_FORMAT)}'
        error_msg = 'You must fetch the data again with the new date range!'
        error_msg += f'{error_msg} (current: {current_date_range}, required: {required_date_range})'
        assert self.from_date <= from_date <= to_date <= self.to_date, error_msg

    def create_presence_report_by_schools(self, from_date: date, to_date: date) -> Dict[str, pd.DataFrame]:
        self.assert_dates_in_range(from_date, to_date)
        from_date = pd.to_datetime(from_date.strftime(self.DATE_FORMAT), format=self.DATE_FORMAT)
        to_date = pd.to_datetime(to_date.strftime(self.DATE_FORMAT), format=self.DATE_FORMAT)
        const_columns = ['מורה אורגני', 'מתרגל', 'יח"ל', 'מצבת']
        periodic_attendance: Dict[str, pd.DataFrame] = dict()
        for school_id, school_data in self.schools_data.items():
            all_behaviors = school_data.behavior_report
            from_date_filter = all_behaviors['lesson_date'] >= pd.to_datetime(from_date)
            to_date_filter = all_behaviors['lesson_date'] <= pd.to_datetime(to_date)
            period_behavior_report = all_behaviors.loc[from_date_filter & to_date_filter]
            school_name = school_data.name
            week_groups = period_behavior_report.groupby(pd.Grouper(key='lesson_date', freq='W'))
            current_weeks_columns = [self.get_date_range_of_week(w.year, w.week) for w in week_groups.groups.keys()]
            current_school_columns = const_columns + current_weeks_columns
            current_school_df = pd.DataFrame(columns=current_school_columns)
            for class_num in range(1, school_data.num_of_active_classes + 1):
                new_row_data = {
                    'מורה אורגני': school_data.get_organic_teacher(class_num),
                    'מתרגל': school_data.get_practitioner(class_num),
                    'יח"ל': school_data.get_level(class_num),
                    'מצבת': school_data.get_num_of_students(class_num)
                }
                for week_key in week_groups.groups.keys():
                    try:
                        week_df = week_groups.get_group(week_key)
                    except KeyError:  # there are no data for that week
                        continue
                    presence_filter = week_df['event_type'] == self.LessonEvents.PRESENCE
                    class_num_filter = week_df['class_num'] == class_num
                    week_events = week_df.loc[presence_filter & class_num_filter, ['student_id', 'lesson_date']]
                    presence_events_in_week_groups = week_events.groupby('lesson_date')
                    if presence_events_in_week_groups.groups.keys():  # calculate average presence in that week
                        num_of_presence = int(round(presence_events_in_week_groups.nunique().mean()))
                    else:  # there is no presence events in that lesson
                        num_of_presence = 0
                    week_column_name = self.get_date_range_of_week(week_key.year, week_key.week)
                    new_row_data[week_column_name] = num_of_presence
                current_school_df = current_school_df.append(new_row_data, ignore_index=True)
            periodic_attendance[school_name] = current_school_df
        return periodic_attendance

    def create_municipal_presence_report_by_levels(self, from_date: date, to_date: date) -> Dict[str, pd.DataFrame]:
        unwanted_columns = ['מתרגל', 'מורה אורגני']
        schools_presence_report = self.create_presence_report_by_schools(from_date, to_date)
        for school_name, school_data_df in schools_presence_report.items():
            school_data_df.insert(0, 'בית ספר', school_name)
        all_schools_data = pd.concat([*schools_presence_report.values()], ignore_index=True)
        all_schools_data.drop(unwanted_columns, axis=1, inplace=True)
        level_groups = all_schools_data.groupby('יח"ל')
        municipal_presence_by_levels = dict()
        for level in level_groups.groups.keys():
            level_group = level_groups.get_group(level).drop('יח"ל', axis=1)
            schools_groups = level_group.groupby('בית ספר')
            each_school_in_row_df = schools_groups.sum().astype(int).reset_index()
            municipal_presence_by_levels[level] = each_school_in_row_df
        return municipal_presence_by_levels

    def create_presence_report_by_month(self, from_month_num: int, to_month_num: int,
                                        from_year: int, to_year: int) -> Dict[str, pd.DataFrame]:
        to_month_week_day, to_month_last_day = calendar.monthrange(to_year, to_month_num)
        from_date = date(from_year, from_month_num, 1)
        to_date = date(to_year, to_month_num, to_month_last_day)
        self.assert_dates_in_range(from_date, to_date)
        from_date = pd.to_datetime(from_date)
        to_date = pd.to_datetime(to_date)
        presence_by_month: Dict[str, pd.DataFrame] = dict()
        columns = ['בית ספר', 'מצבת'] + [MONTHS_IN_HEBREW[month_num] for month_num in
                                         range(from_date.month, to_date.month + 1)]
        all_schools_behavior = pd.DataFrame()
        for school_id, school_data in self.schools_data.items():
            school_name = school_data.name
            school_behavior_df = school_data.behavior_report.copy()
            school_behavior_df.insert(0, 'school_name', school_name)
            school_behavior_df['level'] = school_behavior_df['class_num'].apply(
                self.schools_data[school_id].get_level)
            all_schools_behavior = pd.concat([all_schools_behavior, school_behavior_df])
        from_date_filter = all_schools_behavior['lesson_date'] >= pd.to_datetime(from_date)
        to_date_filter = all_schools_behavior['lesson_date'] <= pd.to_datetime(to_date)
        period_all_schools_behavior = all_schools_behavior.loc[from_date_filter & to_date_filter]
        no_archive_filter = period_all_schools_behavior['level'] != MashovServer.ClassLevel.ARCHIVES
        behavior_no_archive_df = period_all_schools_behavior.loc[no_archive_filter]
        level_groups = behavior_no_archive_df.groupby('level')
        for level in level_groups.groups.keys():
            presence_by_month_df = pd.DataFrame(columns=columns)
            level_df = level_groups.get_group(level)
            school_groups = level_df.groupby('school_name')
            for school_name in school_groups.groups.keys():
                school_df = school_groups.get_group(school_name)
                num_of_students = school_df['student_id'].nunique()
                school_data = {
                    'בית ספר': school_name,
                    'מצבת': num_of_students
                }
                month_groups = school_df.groupby(pd.Grouper(key='lesson_date', freq='M'))
                for month_key in month_groups.groups.keys():
                    try:
                        month_df = month_groups.get_group(month_key)
                    except KeyError:  # there is no data for that month
                        continue
                    presence_filter = month_df['event_type'] == self.LessonEvents.PRESENCE
                    month_presence = month_df.loc[presence_filter]
                    lesson_date_groups = month_presence.groupby('lesson_date')
                    num_of_average_presence = int(round(lesson_date_groups['lesson_date'].count().mean()))
                    school_data[MONTHS_IN_HEBREW[month_key.month]] = num_of_average_presence
                presence_by_month_df = presence_by_month_df.append(school_data, ignore_index=True)
            presence_by_month[level] = presence_by_month_df
        return presence_by_month
