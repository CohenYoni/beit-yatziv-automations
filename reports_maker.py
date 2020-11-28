from data_server import MashovServer, School
from datetime import date
import pandas as pd
import calendar
import typing


class SchoolData(School):
    def __init__(self, school_id: int, name: str, class_code: str):
        super().__init__(school_id, name)
        self.class_code = class_code
        self._behavior_report = None
        self._phonebook = None
        self._grades_report = None
        self._organic_teachers: typing.Dict[int, str] = dict()
        self._teachers: typing.Dict[int, str] = dict()
        self._practitioners: typing.Dict[int, str] = dict()
        self._levels: typing.Dict[int, str] = dict()
        self._num_of_classes = 0

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
    def num_of_classes(self):
        return self._num_of_classes

    @num_of_classes.setter
    def num_of_classes(self, num_of_classes: int) -> None:
        if num_of_classes < 0:
            num_of_classes = 0
        self._num_of_classes = num_of_classes

    def set_organic_teacher(self, class_num: int, organic_teacher_name: str):
        self._organic_teachers[class_num] = organic_teacher_name

    def set_teacher(self, class_num: int, teacher_name: str):
        self._teachers[class_num] = teacher_name

    def set_practitioner(self, class_num: int, practitioner_name: str):
        self._practitioners[class_num] = practitioner_name

    def set_level(self, class_num: int, level: str):
        self._levels[class_num] = level

    def get_organic_teacher(self, class_num: int) -> str:
        return self._organic_teachers.get(class_num, '')

    def get_teacher(self, class_num: int) -> str:
        return self._teachers.get(class_num, '')

    def get_practitioner(self, class_num) -> str:
        return self._practitioners.get(class_num, '')

    def get_level(self, class_num: int) -> str:
        return self._levels.get(class_num, '')


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

    @staticmethod
    def count_events(events_df, event_type):
        event_filter = events_df['event_type'] == event_type
        result = events_df.loc[event_filter, 'event_type']
        num_of_result = result.count()
        return num_of_result

    def __init__(self, schools_ids: list, heb_year: str, class_code: str, username: str, password: str):
        self.schools_data: typing.Dict[int, SchoolData] = {_id: None for _id in schools_ids}
        self.heb_year = heb_year
        self.class_code = class_code
        self.username = username
        self.password = password

    def fetch_data_from_server(self, from_date: date, to_date: date) -> None:
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
                self.schools_data[school_id] = school_class_data
            except Exception:
                raise
            finally:
                server.logout()

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
