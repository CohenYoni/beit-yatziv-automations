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
