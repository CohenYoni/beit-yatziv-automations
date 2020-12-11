from data_server import MashovServer, School
from datetime import datetime, date
from typing import Dict, Sequence
import pandas as pd
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
        self._semesters_grades_report = None
        self._all_grades_report = None
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
    def semesters_grades_report(self) -> pd.DataFrame:
        return self._semesters_grades_report

    @semesters_grades_report.setter
    def semesters_grades_report(self, semesters_grades_report: pd.DataFrame) -> None:
        self._semesters_grades_report = semesters_grades_report

    @property
    def all_grades_report(self) -> pd.DataFrame:
        return self._all_grades_report

    @all_grades_report.setter
    def all_grades_report(self, all_grades_report: pd.DataFrame) -> None:
        self._all_grades_report = all_grades_report

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

    class Semester:
        BEGIN_YEAR_EXAM = 'תחילת שנה'
        END_SEMESTER_1 = MashovServer.SEMESTER_EXAM_MAPPER['end_semester1']
        BEGIN_SEMESTER_2 = MashovServer.SEMESTER_EXAM_MAPPER['begin_semester2']
        END_SEMESTER_2 = MashovServer.SEMESTER_EXAM_MAPPER['end_semester2']

    HEB_WEEKDAYS = ['שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת', 'ראשון']
    NUM_TO_HEB_CLASS_MAPPER = {num: code for num, code in enumerate('א ב ג ד ה ו ז ח ט י יא יב'.split(), 1)}
    HEB_CLASS_TO_NUM_MAPPER = {code: num for num, code in enumerate('א ב ג ד ה ו ז ח ט י יא יב'.split(), 1)}
    HEB_DAYS_MAPPER = {idx: day for idx, day in enumerate(HEB_WEEKDAYS, 0)}
    SEMESTER_EXAMS_MAPPER = {
        'first_year': Semester.BEGIN_YEAR_EXAM,
        'end_semester1': Semester.END_SEMESTER_1,
        'begin_semester2': Semester.BEGIN_SEMESTER_2,
        'end_semester2': Semester.END_SEMESTER_2
    }
    NO_REMARKS = 'ללא הערות'
    FAIL_GRADE_THRESHOLD = 85
    NEGATIVE_GRADE_THRESHOLD = 56
    GREEN_GRADE_THRESHOLD = 76
    RED_GRADE_THRESHOLD = 55
    GRADES_COLORS_MAPPER = {
        'red': f'אדום (קטן מ-{RED_GRADE_THRESHOLD})',
        'orange': f'כתום (בין {RED_GRADE_THRESHOLD + 1} ל-{GREEN_GRADE_THRESHOLD - 1})',
        'green': f'ירוק (מעל {GREEN_GRADE_THRESHOLD})'
    }
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

    @staticmethod
    def get_previous_class_code(class_code: str) -> str:
        if class_code not in ReportMaker.HEB_CLASS_TO_NUM_MAPPER:
            raise ValueError(f'{class_code} is not a valid class code!')
        class_code_num = ReportMaker.HEB_CLASS_TO_NUM_MAPPER[class_code]
        if class_code_num == 1:
            raise ValueError(f'{class_code} is the first class!')
        return ReportMaker.NUM_TO_HEB_CLASS_MAPPER[class_code_num - 1]

    def __init__(self, schools_ids: list, heb_year: str, class_code: str, username: str, password: str):
        self.schools_data: Dict[int, SchoolData] = {_id: None for _id in schools_ids}
        self.heb_year = heb_year
        self.class_code = class_code
        self.username = username
        self.password = password
        self.from_date = None
        self.to_date = None
        self._greg_year = MashovServer.map_heb_year_to_greg(self.heb_year)
        self._first_school_year_date = date(year=self._greg_year - 1, month=8, day=1)
        self._last_school_year_date = date(year=self._greg_year, month=11, day=30)
        self._previous_heb_year = MashovServer.map_greg_year_to_heb(self._greg_year - 1)

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
                semesters_grades_report = server.get_grades_report(from_date=from_date,
                                                                   to_date=to_date,
                                                                   class_code=self.class_code,
                                                                   exam_type=MashovServer.ExamType.SEMESTER_EXAM)
                all_grades_report = server.get_grades_report(from_date=from_date,
                                                             to_date=to_date,
                                                             class_code=self.class_code,
                                                             exam_type=MashovServer.ExamType.ALL)
                school_class_data = SchoolData(school_id, server.school.name, self.class_code)
                school_class_data.behavior_report = behavior_report
                school_class_data.phonebook = phonebook
                school_class_data.semesters_grades_report = semesters_grades_report
                school_class_data.all_grades_report = all_grades_report
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

    def create_grades_colors_report_by_levels(self):
        all_schools_grades_df = pd.DataFrame()
        for school_id in self.schools_data.keys():
            server = MashovServer(school_id=school_id, school_year=self.heb_year)
            server.login(username=self.username, password=self.password)
            current_year_grades_df = server.get_grades_report(from_date=self._first_school_year_date,
                                                              to_date=self._last_school_year_date,
                                                              class_code=self.class_code,
                                                              exam_type=MashovServer.ExamType.SEMESTER_EXAM)
            server.logout()
            try:
                server.school_year = self._previous_heb_year
                there_is_prev_year = True
            except TypeError:  # there are no data of previous year in the server
                there_is_prev_year = False
            if there_is_prev_year:
                prev_greg_year = self._greg_year - 1
                prev_from_date = self._first_school_year_date.replace(year=prev_greg_year - 1)
                prev_to_date = self._last_school_year_date.replace(year=prev_greg_year)
                prev_class_code = self.get_previous_class_code(self.class_code)
                server.login(username=self.username, password=self.password)
                prev_year_grades_df = server.get_grades_report(from_date=prev_from_date, to_date=prev_to_date,
                                                               class_code=prev_class_code,
                                                               exam_type=MashovServer.ExamType.SEMESTER_EXAM)
                server.logout()
            else:
                prev_year_grades_df = pd.DataFrame(columns=current_year_grades_df.columns,
                                                   index=current_year_grades_df.index)
            prev_first_year_grade_df = prev_year_grades_df['end_semester2']
            new_col_pos = current_year_grades_df.columns.get_loc('end_semester1')
            current_year_grades_df.insert(new_col_pos, 'first_year', prev_first_year_grade_df)
            all_schools_grades_df = pd.concat([all_schools_grades_df, current_year_grades_df])
        grades_colors_by_level = dict()
        levels_groups = all_schools_grades_df.groupby('level')
        for level_key in levels_groups.groups.keys():
            level_df = levels_groups.get_group(level_key)
            schools_grades = dict()
            schools_groups = level_df.groupby('school_name')
            for school_key in schools_groups.groups.keys():
                school_df = schools_groups.get_group(school_key)
                exam_periods = dict()
                for exam_period_key, exam_period_col in self.SEMESTER_EXAMS_MAPPER.items():
                    red_filter = school_df[exam_period_key] <= self.RED_GRADE_THRESHOLD
                    low_orange_filter = school_df[exam_period_key] > self.RED_GRADE_THRESHOLD
                    high_orange_filter = school_df[exam_period_key] < self.GREEN_GRADE_THRESHOLD
                    green_filter = school_df[exam_period_key] >= self.GREEN_GRADE_THRESHOLD
                    num_of_red = school_df[red_filter][exam_period_key].count()
                    num_of_orange = school_df[high_orange_filter & low_orange_filter][exam_period_key].count()
                    num_of_green = school_df[green_filter][exam_period_key].count()
                    columns = list(self.GRADES_COLORS_MAPPER.values()) + ['סה"כ', ]
                    data = [num_of_red, num_of_orange, num_of_green]
                    sum_of_grades = sum(data)
                    data.append(sum_of_grades)
                    exam_periods[exam_period_col] = pd.DataFrame([data], columns=columns)
                schools_grades[school_key] = exam_periods
            grades_colors_by_level[level_key] = schools_grades
        return grades_colors_by_level

    def create_summary_report_by_school(self, from_date: date, to_date: date) -> Dict[str, pd.DataFrame]:
        self.assert_dates_in_range(from_date, to_date)
        all_schools_behavior_df = pd.DataFrame()
        school_name_to_id_mapper = dict()
        for school_id in self.schools_data.keys():
            behavior_df = self.schools_data[school_id].behavior_report.copy()
            school_name = self.schools_data[school_id].name
            behavior_df['school_name'] = school_name
            school_name_to_id_mapper[school_name] = school_id
            all_schools_behavior_df = pd.concat([all_schools_behavior_df, behavior_df])
        from_date = pd.to_datetime(from_date.strftime(self.DATE_FORMAT), format=self.DATE_FORMAT)
        to_date = pd.to_datetime(to_date.strftime(self.DATE_FORMAT), format=self.DATE_FORMAT)
        from_date_filter = all_schools_behavior_df['lesson_date'] >= pd.to_datetime(from_date)
        to_date_filter = all_schools_behavior_df['lesson_date'] <= pd.to_datetime(to_date)
        period_behavior_report = all_schools_behavior_df.loc[from_date_filter & to_date_filter]
        required_columns = ['lesson_date', 'school_name', 'class_num', 'lesson_num', 'student_id', 'event_type']
        period_behavior_report = period_behavior_report[required_columns]
        schools_groups = period_behavior_report.groupby('school_name')
        schools_summary = dict()
        for school_key in schools_groups.groups.keys():
            school_id = school_name_to_id_mapper[school_key]
            school_details_df = schools_groups.get_group(school_key)
            school_details_df['level'] = school_details_df['class_num'].apply(
                self.schools_data[school_id].get_level)
            no_archive_filter = school_details_df['level'] != MashovServer.ClassLevel.ARCHIVES
            school_details_df = school_details_df.loc[no_archive_filter]
            lesson_groups = school_details_df.groupby(['class_num', 'lesson_date', 'lesson_num'])
            school_summary_df = pd.DataFrame()
            total_num_students = lesson_groups.apply(lambda group: group['student_id'].nunique())
            school_summary_df['מצבת'] = total_num_students
            num_of_presence = lesson_groups.apply(
                lambda group: group.loc[
                    group['event_type'] == self.LessonEvents.PRESENCE, 'student_id'].nunique())
            school_summary_df['נוכחים'] = num_of_presence
            num_of_missing = lesson_groups.apply(
                lambda group: group.loc[
                    group['event_type'] == self.LessonEvents.MISSING, 'student_id'].nunique())
            school_summary_df['חיסורים'] = num_of_missing
            num_of_disturbs = lesson_groups.apply(
                lambda group: group.loc[
                    group['event_type'] == self.LessonEvents.DISTURB, 'student_id'].nunique())
            school_summary_df['הפרעה'] = num_of_disturbs
            school_summary_df['אחוז נוכחות'] = round(
                (school_summary_df['נוכחים'] / school_summary_df['מצבת']) * 100).astype(int).astype(str) + '%'
            all_grades_report_df = self.schools_data[school_id].all_grades_report.copy()
            from_date_filter = all_grades_report_df['exam_date'] >= pd.to_datetime(from_date)
            to_date_filter = all_grades_report_df['exam_date'] <= pd.to_datetime(to_date)
            period_grades_report = all_grades_report_df.loc[from_date_filter & to_date_filter]
            period_grades_report.rename(columns={'exam_date': 'lesson_date'}, inplace=True)
            grades_groups = period_grades_report.groupby(['class_num', 'lesson_date'])
            num_of_testing_students = grades_groups.apply(lambda group: group['student_id'].nunique())
            num_of_failed_students = grades_groups.apply(lambda group: group.loc[
                group['exam_grade'] < self.FAIL_GRADE_THRESHOLD, 'student_id'].nunique())
            grps_to_max_lesson_num = school_summary_df.reset_index().groupby(['class_num', 'lesson_date'])
            max_lesson_num = grps_to_max_lesson_num.apply(lambda grp: grp['lesson_num'].max())
            grades_summary_df = pd.DataFrame()
            grades_summary_df['lesson_num'] = max_lesson_num
            if num_of_testing_students.empty:
                num_of_testing_students = pd.Series()
            if num_of_failed_students.empty:
                num_of_failed_students = pd.Series()
            grades_summary_df['מגישים'] = num_of_testing_students
            grades_summary_df[f'נכשלים (מתחת {self.FAIL_GRADE_THRESHOLD})'] = num_of_failed_students
            grades_summary_df.set_index('lesson_num', append=True, inplace=True)
            school_summary_df = school_summary_df.join(grades_summary_df)
            school_summary_df.reset_index(inplace=True)
            cols_to_replace = ['מצבת', 'נוכחים', 'חיסורים', 'הפרעה', 'מגישים',
                               f'נכשלים (מתחת {self.FAIL_GRADE_THRESHOLD})']
            school_summary_df[cols_to_replace] = school_summary_df[cols_to_replace].replace(0, pd.NA)
            school_summary_df.sort_values(['lesson_date', 'class_num', 'lesson_num'], ascending=[True, True, True],
                                          inplace=True, ignore_index=True)
            date_range = f'{from_date.strftime(format=self.DATE_FORMAT)}-{to_date.strftime(format=self.DATE_FORMAT)}'
            school_summary_df['טווח זמן'] = date_range
            school_summary_df['יח"ל'] = school_summary_df['class_num'].apply(
                self.schools_data[school_id].get_level)
            school_summary_df['מורה אורגני'] = school_summary_df['class_num'].apply(
                self.schools_data[school_id].get_organic_teacher)
            school_summary_df['כיתה/קבוצת לימוד'] = school_summary_df['class_num'].apply(
                self.schools_data[school_id].get_practitioner)
            school_summary_df['שכבה'] = self.class_code
            school_summary_df.rename(columns={
                'lesson_date': 'תאריך שיעור',
                'class_num': 'כיתה',
                'lesson_num': 'מספר שיעור'
            }, inplace=True)
            cols_order = [
                'טווח זמן',
                'תאריך שיעור',
                'שכבה',
                'כיתה',
                'מורה אורגני',
                'כיתה/קבוצת לימוד',
                'מספר שיעור',
                'יח"ל',
                'מצבת',
                'נוכחים',
                'חיסורים',
                'מגישים',
                f'נכשלים (מתחת {self.FAIL_GRADE_THRESHOLD})',
                'הפרעה',
                'אחוז נוכחות'
            ]
            school_summary_df = school_summary_df[cols_order]
            school_summary_df['סיבת החיסורים וטיפול בהפרעות (ואסים)'] = pd.NA
            school_summary_df['הערות'] = pd.NA
            schools_summary[school_key] = school_summary_df
        return schools_summary

    def create_raw_behavior_report(self, from_date: date, to_date: date) -> Dict[str, pd.DataFrame]:
        raw_behavior_by_schools = dict()
        for school_id in self.schools_data.keys():
            behavior_df = self.schools_data[school_id].behavior_report.copy()
            from_date = pd.self(from_date.strftime(self.DATE_FORMAT), format=self.DATE_FORMAT)
            to_date = pd.to_datetime(to_date.strftime(self.DATE_FORMAT), format=self.DATE_FORMAT)
            from_date_filter = behavior_df['lesson_date'] >= pd.to_datetime(from_date)
            to_date_filter = behavior_df['lesson_date'] <= pd.to_datetime(to_date)
            behavior_df = behavior_df.loc[from_date_filter & to_date_filter]
            behavior_df['level'] = behavior_df['class_num'].apply(self.schools_data[school_id].get_level)
            behavior_df['practitioner'] = behavior_df['class_num'].apply(
                self.schools_data[school_id].get_practitioner)
            no_archive_filter = behavior_df['level'] != MashovServer.ClassLevel.ARCHIVES
            behavior_df = behavior_df.loc[no_archive_filter]
            column_names_mapper = {
                'teacher_name': 'שם המורה',
                'subject': 'מקצוע',
                'lesson_date': 'תאריך',
                'lesson_num': 'מספר שיעור',
                'student_id': 'ת.ז',
                'student_name': 'שם התלמיד',
                'class_code': 'שכבה',
                'class_num': 'כיתה',
                'event_type': 'סוג האירוע',
                'remark': 'הערה מילולית',
                'justified_by': 'הוצדק ע"י',
                'justification': 'הצדקה',
                'level': 'יח"ל',
                'practitioner': 'מתרגל',
            }
            behavior_df.rename(columns=column_names_mapper, inplace=True)
            behavior_df.reset_index(drop=True, inplace=True)
            raw_behavior_by_schools[self.schools_data[school_id].name] = behavior_df
        return raw_behavior_by_schools
