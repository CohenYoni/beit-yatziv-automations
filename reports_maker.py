from data_scraper import MashovScraper, School
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

    def __init__(self, schools_ids: list, heb_year: str, class_code: str, username: str, password: str):
        self.schools_data: typing.Dict[int, SchoolData] = {_id: None for _id in schools_ids}
        self.heb_year = heb_year
        self.class_code = class_code
        self.username = username
        self.password = password

    def fetch_data_from_server(self, from_date: date, to_date: date) -> None:
        for school_id in self.schools_data.keys():
            scraper = MashovScraper(school_id=school_id, school_year=self.heb_year)
            try:
                scraper.login(username=self.username, password=self.password)
                behavior_report = scraper.get_behavior_report_by_dates(from_date=from_date,
                                                                       to_date=to_date,
                                                                       class_code=self.class_code)
                phonebook = scraper.get_students_phonebook(class_code=self.class_code)
                grades_report = scraper.get_grades_report(from_date=from_date,
                                                          to_date=to_date,
                                                          class_code=self.class_code)
                school_class_data = SchoolData(school_id, scraper.school.name, self.class_code)
                school_class_data.behavior_report = behavior_report
                school_class_data.phonebook = phonebook
                school_class_data.grades_report = grades_report
                self.schools_data[school_id] = school_class_data
            except Exception:
                raise
            finally:
                scraper.logout()

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
