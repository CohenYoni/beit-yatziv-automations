from datetime import datetime
from datetime import date
from typing import Dict
import pandas as pd
import numpy as np
import requests
import urllib
import re


class School:
    def __init__(self, school_id: int, name: str):
        self.school_id = school_id
        self.name = name

    @property
    def school_id(self) -> int:
        return self._school_id

    @school_id.setter
    def school_id(self, school_id: int) -> None:
        if type(school_id) != int:
            raise TypeError(f'School ID must be an integer, not {type(school_id)}')
        self._school_id = school_id

    @property
    def name(self) -> str:
        return self._name

    @name.setter
    def name(self, name: str) -> None:
        if type(name) != str:
            raise TypeError(f'School name must be a string, not {type(name)}')
        self._name = name

    def __str__(self) -> str:
        return self.name

    def __repr__(self) -> str:
        return str(self)

    def __eq__(self, other: 'School') -> bool:
        return self.school_id == other.school_id


class Class:
    def __init__(self, class_code: str, class_num: int, practitioner: str, level: str):
        self._class_code = class_code
        self._class_num = class_num
        self._practitioner = practitioner
        self._level = level

    @property
    def class_code(self) -> str:
        return self._class_code

    @class_code.setter
    def class_code(self, class_code: str) -> None:
        self._class_code = class_code

    @property
    def class_num(self) -> int:
        return self._class_num

    @class_num.setter
    def class_num(self, class_num: int) -> None:
        self._class_num = class_num

    @property
    def practitioner(self) -> str:
        return self._practitioner

    @practitioner.setter
    def practitioner(self, practitioner: str) -> None:
        self._practitioner = practitioner

    @property
    def level(self) -> str:
        return self._level

    @level.setter
    def level(self, level: str) -> None:
        self._level = level

    def __str__(self) -> str:
        return f'{self.class_code}{self._class_num} - {self.practitioner} - {self.level}'

    def __repr__(self) -> str:
        return str(self)


class MashovServer:
    class ClassLevel:
        NO_LEVEL = 'ללא'
        ARCHIVES = 'ארכיון'

    class ExamType:
        SEMESTER_EXAM = 1
        ALL = 2

    HEB_TO_GREG_YEAR_MAPPER = {
        'תשעה': 2015,
        'תשעו': 2016,
        'תשעז': 2017,
        'תשעח': 2018,
        'תשעט': 2019,
        'תשפ': 2020,
        'תשף': 2020,
        'תשפא': 2021,
        'תשפב': 2022,
        'תשפג': 2023,
        'תשפד': 2024,
        'תשפה': 2025,
        'תשפו': 2026,
        'תשפז': 2027,
        'תשפח': 2028,
        'תשפט': 2029,
        'תשצ': 2030,
        'תשץ': 2030,
        'תשצא': 2031,
        'תשצב': 2032,
        'תשצג': 2033,
        'תשצד': 2034,
        'תשצה': 2035,
        'תשצו': 2036,
        'תשצז': 2037,
        'תשצח': 2038,
        'תשצט': 2039,
        'תת': 2040
    }
    SEMESTER_EXAM_MAPPER = {
        'end_semester1': 'סוף סמסטר א',
        'begin_semester2': 'תחילת סמסטר ב',
        'end_semester2': 'סוף סמסטר ב'
    }
    CHROME_VERSION = '86.0.4240.198'
    CHROME_UA = f'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) ' \
                f'Chrome/{CHROME_VERSION} Safari/537.36'
    BASE_URL = 'https://web.mashov.info'
    LOGIN_PAGE_URL = f'{BASE_URL}/teachers/login'
    CLEAR_SESSION_URL = f'{BASE_URL}/api/clearSession'
    LOGIN_API_URL = f'{BASE_URL}/api/login'
    MAIN_DASHBOARD_PAGE_URL = f'{BASE_URL}/teachers/main/dashboard'
    LOGOUT_URL = f'{BASE_URL}/api/logout'
    FAILED_GRADE_THRESHOLD = 56
    DATE_FORMAT = '%d/%m/%Y'
    EXAM_TYPE_WORD = 'מבחן'

    @staticmethod
    def map_heb_year_to_greg(heb_year: str) -> int:
        clean_heb_year = heb_year.replace('\"', '').replace('\'', '')
        if clean_heb_year not in MashovServer.HEB_TO_GREG_YEAR_MAPPER.keys():
            raise TypeError(f'{heb_year} is not a correct Hebrew year!')
        return MashovServer.HEB_TO_GREG_YEAR_MAPPER[clean_heb_year]

    @staticmethod
    def map_greg_year_to_heb(greg_year: int) -> str:
        heb_year = ''
        for heb, greg in MashovServer.HEB_TO_GREG_YEAR_MAPPER.items():
            if greg_year == greg:
                heb_year = heb
                break
        assert heb_year, f'{greg_year} is not a correct year!'
        return heb_year

    def __init__(self, school_id: int, school_year: str):
        res = requests.get('https://web.mashov.info/api/schools', headers={'User-Agent': self.CHROME_UA})
        try:
            res.raise_for_status()
        except requests.exceptions.HTTPError:
            raise requests.exceptions.HTTPError('An error occurred while fetching school list from the server')
        all_schools_json = res.json()
        self._api_version = res.headers['apiversion']
        self._all_schools = dict()
        for school_detail in all_schools_json:
            self._all_schools[school_detail.pop('semel')] = school_detail
        self.school = school_id
        self.school_year = school_year
        self._session = requests.Session()
        const_headers = {
            'Connection': 'close',
            'User-Agent': self.CHROME_UA,
            'Accept': 'application/json, text/plain, */*',
        }
        self._session.headers.update(const_headers)
        self._csrf_token = ''
        self._auth_json_response = dict()
        self._logged_in = False
        self.classes_details: Dict[str, Dict[int, Class]] = dict()

    @property
    def school(self) -> School:
        return self._school

    @school.setter
    def school(self, school_id: int) -> None:
        if type(school_id) != int:
            raise TypeError(f'School must be an integer, not {type(school_id)}')
        if school_id not in self._all_schools.keys():
            raise TypeError(f'{school_id} is no a correct school')
        self._school = School(school_id=school_id, name=self._all_schools[school_id]['name'])

    @property
    def school_year(self) -> int:
        return self._school_year

    @school_year.setter
    def school_year(self, school_year: str) -> None:
        if type(school_year) != str:
            raise TypeError(f'School year must be a string, not {type(school_year)}')
        assert self.school, 'You must initialize the school first!'
        gregorian_year = self.map_heb_year_to_greg(school_year)
        school_years = self._all_schools[self.school.school_id]['years']
        if gregorian_year not in school_years:
            possible_years = ', '.join([self.map_greg_year_to_heb(year) for year in school_years])
            raise TypeError(f'There is not {school_year} in {self.school.name} (possible years are {possible_years})')
        self._school_year = gregorian_year

    def assert_logged_in(self):
        assert self._logged_in, 'You must log in first!'

    def login(self, username: str, password: str) -> None:
        login_json_data = {
            'apiVersion': self._api_version,
            'appBuild': self._api_version,
            'appName': 'info.mashov.teachers',
            'appVersion': self._api_version,
            'deviceManufacturer': 'win',
            'deviceModel': 'desktop',
            'devicePlatform': 'chrome',
            'deviceUuid': 'chrome',
            'deviceVersion': self.CHROME_VERSION,
            'password': password,
            'semel': self.school.school_id,
            'username': username,
            'year': self.school_year
        }
        self._session.get(self.LOGIN_PAGE_URL)
        self._session.get(self.CLEAR_SESSION_URL, headers={'Referer': self.LOGIN_PAGE_URL})
        res = self._session.post(self.LOGIN_API_URL, headers={'Referer': self.LOGIN_PAGE_URL}, json=login_json_data)
        try:
            res.raise_for_status()
        except requests.exceptions.HTTPError:
            raise requests.exceptions.HTTPError(f'Login process failed: {res.reason}: {res.text}')
        self._auth_json_response = res.json()
        self._csrf_token = self._session.cookies.get('Csrf-Token')
        self._session.headers.update({'X-Csrf-Token': self._csrf_token})
        self._logged_in = True
        self.get_classes_details()

    def logout(self) -> None:
        if not self._logged_in:
            return
        self._session.get(self.LOGOUT_URL, headers={'Referer': self.MAIN_DASHBOARD_PAGE_URL})
        self._session.close()
        self._logged_in = False

    def get_behavior_report_by_dates(self, from_date: date, to_date: date, class_code: str) -> pd.DataFrame:
        self.assert_logged_in()

        def parse_json_res(res_obj: dict) -> list:
            student_json = res_obj.get('student', {})
            lesson_log_json = res_obj.get('lessonLog', {})
            teacher_json = res_obj.get('teacher', {})
            achva_json = res_obj.get('achva', {})
            achva_remark_json = res_obj.get('achvaRemark', {})
            justified_by_json = res_obj.get('justifiedBy', {})
            achva_justification_json = res_obj.get('achvaJustification', {})
            lesson_date = lesson_log_json.get('lessonDate', '')
            date_format = '%Y-%m-%dT%H:%M:%S'
            try:
                lesson_date = datetime.strptime(lesson_date, date_format).strftime('%d/%m/%Y')
            except:
                if lesson_date:
                    print(f'Warning: Lesson date format changed! ("{lesson_date}" instead of {date_format})')
            required_data = [
                teacher_json.get('teacherName', ''),
                res_obj.get('subjectName', ''),
                lesson_date,
                lesson_log_json.get('lesson', ''),
                student_json.get('studentId', ''),
                f'{student_json.get("familyName", "")} {student_json.get("privateName", "")}',
                student_json.get('classCode', ''),
                student_json.get('classNum', ''),
                achva_json.get('name', ''),
                achva_remark_json.get('remarkText', ''),
                justified_by_json.get('teacherName', ''),
                achva_justification_json.get('justification', '')
            ]
            return required_data

        encoded_class = urllib.parse.quote(class_code)
        from_date = f"{from_date.strftime('%Y-%m-%d')}T00:00:00Z"
        to_date = f"{to_date.strftime('%Y-%m-%d')}T23:59:59Z"
        url = f'{self.BASE_URL}/api/classes/{encoded_class}/behave?start={from_date}&end={to_date}'
        res = self._session.get(url, headers={'Referer': self.MAIN_DASHBOARD_PAGE_URL})
        json_res = res.json()
        columns = ['teacher_name', 'subject', 'lesson_date', 'lesson_num', 'student_id', 'student_name', 'class_code',
                   'class_num', 'event_type', 'remark', 'justified_by', 'justification']
        data = [parse_json_res(v) for v in json_res]
        behavior_report_df = pd.DataFrame(data, columns=columns)
        behavior_report_df['lesson_date'] = pd.to_datetime(behavior_report_df['lesson_date'], format=self.DATE_FORMAT)
        return behavior_report_df

    def get_students_phonebook(self, class_code: str) -> pd.DataFrame:
        self.assert_logged_in()

        def parse_json_res(details: dict, extra_data: dict) -> list:
            student_json = details.get('student', {})
            student_info_json = details.get('studentInfo', {})
            contacts_list = details.get('contacts', [])
            parent1_json = contacts_list[0] if len(contacts_list) > 0 else {}
            parent2_json = contacts_list[1] if len(contacts_list) > 1 else {}
            student_id = str(student_json.get('studentId', ''))
            student_extra_data_json = extra_data.get(student_id, {})
            date_format = '%Y-%m-%dT%H:%M:%S'
            student_birthdate = student_json.get('birthDate', '')
            try:
                student_birthdate = datetime.strptime(student_birthdate, date_format).strftime('%d/%m/%Y')
            except:
                if student_birthdate:
                    msg = f'Warning: Student birthdate format changed! ("{student_birthdate}" instead of {date_format})'
                    print(msg)
            required_data = [
                student_id,
                student_json.get('familyName', ''),
                student_json.get('privateName', ''),
                student_json.get('gender', ''),
                student_json.get('classCode', ''),
                student_json.get('classNum', ''),
                student_birthdate,
                student_json.get('hebrewBirthDate', ''),
                student_json.get('major', ''),
                student_info_json.get('city1', ''),
                student_info_json.get('address1', ''),
                student_info_json.get('city2', ''),
                student_info_json.get('address2', ''),
                student_info_json.get('phone1', ''),
                student_info_json.get('email1', ''),
                student_info_json.get('cellphone1', ''),
                parent1_json.get('contact', {}).get('contactId', ''),
                parent1_json.get('contact', {}).get('privateName', ''),
                parent1_json.get('contactInfo', {}).get('email1', ''),
                parent1_json.get('contactInfo', {}).get('cellphone1', ''),
                parent2_json.get('contact', {}).get('contactId', ''),
                parent2_json.get('contact', {}).get('privateName', ''),
                parent2_json.get('contactInfo', {}).get('email1', ''),
                parent2_json.get('contactInfo', {}).get('cellphone1', ''),
                '',
                student_extra_data_json.get('NoBro', ''),
                student_extra_data_json.get('NoCom', ''),
                student_extra_data_json.get('OrClass', ''),
                student_extra_data_json.get('OrTeacher', ''),
                student_extra_data_json.get('rama', ''),
                student_extra_data_json.get('Sat', ''),
                '',
                ''
            ]
            return required_data

        encoded_class = urllib.parse.quote(class_code)
        details_url = f'{self.BASE_URL}/api/classes/{encoded_class}/students/details'
        details_res = self._session.get(details_url, headers={'Referer': self.MAIN_DASHBOARD_PAGE_URL})
        json_details_res = details_res.json()
        extra_data_url = f'{self.BASE_URL}/api/classes/{encoded_class}/students/extraData'
        extra_data_res = self._session.get(extra_data_url, headers={'Referer': self.MAIN_DASHBOARD_PAGE_URL})
        json_extra_data_res = extra_data_res.json()
        for _id, det_list in json_extra_data_res.items():
            new_dict = dict()
            for el in det_list:
                if 'columnName' in el:
                    new_dict[el['columnName']] = el.get('val', '')
            json_extra_data_res[_id] = new_dict
        columns = ['student_id', 'family_name', 'private_name', 'gender', 'class_code', 'class_num',
                   'birthdate', 'heb_birthdate', 'study_trend', 'main_city', 'main_address', 'sec_city', 'sec_address',
                   'home_phone', 'student_mail', 'student_phone_num', 'parent1_id', 'parent1_name', 'parent1_mail',
                   'parent1_phone_num', 'parent2_id', 'parent2_name', 'parent2_mail', 'parent2_phone_num',
                   'edge_means', 'num_brothers', 'num_computers', 'original_class', 'original_teacher', 'level',
                   'saturday_practitioner', 'material_help', 'home_visits']
        data = [parse_json_res(v, json_extra_data_res) for v in json_details_res]
        phonebook_df = pd.DataFrame(data, columns=columns)
        return phonebook_df

    def get_grades_report(self, from_date: date, to_date: date, class_code: str, exam_type: int) -> pd.DataFrame:
        self.assert_logged_in()

        def parse_semesters_grades_json_res(grades_json: dict) -> pd.DataFrame:
            const_columns = ['school_name', 'student_name', 'class_code', 'class_num', 'level']
            exams_columns = ['end_semester1', 'begin_semester2', 'end_semester2']
            exam_name_to_column_mapper = {name: col for col, name in self.SEMESTER_EXAM_MAPPER.items()}
            phonebook = self.get_students_phonebook(class_code)
            phonebook['student_id'] = pd.to_numeric(phonebook['student_id'])
            phonebook.set_index('student_id', inplace=True)
            phonebook['student_name'] = phonebook['family_name'] + ' ' + phonebook['private_name']
            phonebook['school_name'] = self.school.name
            phonebook['level'] = phonebook.apply(
                lambda row: self.get_class_level(row['class_code'], row['class_num']), axis=1
            )
            df = pd.DataFrame(columns=const_columns + exams_columns, index=pd.Index([], name='student_id'))
            df = df.append(phonebook[const_columns])
            for grade in grades_json:
                exam_type = grade.get('gradeType', {}).get('name', '')
                if exam_type == self.EXAM_TYPE_WORD:
                    exam_name = grade.get('gradingEvent', {}).get('name', '')
                    exam_column = exam_name_to_column_mapper.get(exam_name)
                    if exam_column:
                        student_json = grade.get('student', {})
                        student_id = student_json.get('studentId', np.NaN)
                        if student_id and student_id not in df.index:
                            family_name = student_json.get('familyName', '')
                            private_name = student_json.get('privateName', '')
                            student_class_code = student_json.get('classCode', '')
                            student_class_num = student_json.get('classNum', '')
                            new_student_data = {
                                'student_name': f"{family_name} {private_name}",
                                'class_code': student_class_code,
                                'class_num': student_class_num,
                                'school_name': self.school.name,
                                'level': self.get_class_level(student_class_code, student_class_num)
                            }
                            df = df.append(pd.DataFrame(new_student_data, index=[student_id]))
                        df.loc[student_id, exam_column] = grade.get('grade', {}).get('grade', np.NaN)
            df.replace('', np.NaN, inplace=True)
            df.sort_index(inplace=True)
            df.rename(columns={df.index.name: 'student_id'}, inplace=True)
            no_archives_filter = df['level'] != self.ClassLevel.ARCHIVES
            df = df.loc[no_archives_filter]
            return df

        def parse_all_grades_json_res(grades_json: dict) -> pd.DataFrame:
            const_columns = ['school_name', 'student_id', 'student_name', 'class_code', 'class_num', 'level',
                             'exam_date', 'exam_grade', 'exam_type', 'exam_name', 'exam_subject']
            df = pd.DataFrame(columns=const_columns)
            for grade in grades_json:
                current_exam_type = grade.get('gradeType', {}).get('name', '')
                exam_date = grade.get('gradingEvent', {}).get('eDate', '')
                date_format = '%Y-%m-%dT%H:%M:%S'
                try:
                    exam_date = datetime.strptime(exam_date, date_format).strftime('%d/%m/%Y')
                except:
                    if exam_date:
                        print(f'Warning: Exam date format changed! ("{exam_date}" instead of {date_format})')
                student_json = grade.get('student', {})
                student_id = student_json.get('studentId', pd.NA)
                family_name = student_json.get('familyName', '')
                private_name = student_json.get('privateName', '')
                student_class_code = student_json.get('classCode', '')
                student_class_num = student_json.get('classNum', '')
                exam_grade = grade.get('grade', {}).get('grade', pd.NA)
                exam_name = grade.get('gradingEvent', {}).get('name', pd.NA)
                exam_subject = grade.get('group', {}).get('subjectName', pd.NA)
                new_grade_data = {
                    'school_name': self.school.name,
                    'student_id': student_id,
                    'student_name': f"{family_name} {private_name}",
                    'class_code': student_class_code,
                    'class_num': student_class_num,
                    'level': self.get_class_level(student_class_code, student_class_num),
                    'exam_date': exam_date,
                    'exam_grade': exam_grade,
                    'exam_type': current_exam_type,
                    'exam_name': exam_name,
                    'exam_subject': exam_subject
                }
                df = df.append(new_grade_data, ignore_index=True)
            df.replace('', pd.NA, inplace=True)
            no_archives_filter = df['level'] != self.ClassLevel.ARCHIVES
            df = df.loc[no_archives_filter]
            df.reset_index(inplace=True, drop=True)
            df['exam_date'] = pd.to_datetime(df['exam_date'], format='%d/%m/%Y')
            return df

        parser_mapper = {
            self.ExamType.SEMESTER_EXAM: parse_semesters_grades_json_res,
            self.ExamType.ALL: parse_all_grades_json_res
        }

        encoded_class = urllib.parse.quote(class_code)
        from_date = f"{from_date.strftime('%Y-%m-%d')}T00:00:00Z"
        to_date = f"{to_date.strftime('%Y-%m-%d')}T23:59:59Z"
        grades_url = f'{self.BASE_URL}/api/classes/{encoded_class}/grades?start={from_date}&end={to_date}'
        grades_res = self._session.get(grades_url, headers={'Referer': self.MAIN_DASHBOARD_PAGE_URL})
        json_grades_res = grades_res.json()
        grades_df = parser_mapper[exam_type](json_grades_res)
        return grades_df

    def get_classes_details(self):
        self.assert_logged_in()
        classes_url = f'{self.BASE_URL}/api/classes'
        classes_res = self._session.get(classes_url, headers={'Referer': self.MAIN_DASHBOARD_PAGE_URL})
        json_classes_res = classes_res.json()
        for json_class in json_classes_res:
            class_code = json_class.get('classCode')
            class_num = json_class.get('classNum')
            class_name_attr = json_class.get('className', '')
            level_regex_pattern = re.compile(r'(?P<practitioner>[א-ת ]+?) *- *(?P<level>.?\d.? יח\"?ל)')
            regex_search_res = level_regex_pattern.search(class_name_attr)
            if regex_search_res:
                practitioner = regex_search_res.group('practitioner')
                level = regex_search_res.group('level')
            else:
                practitioner = class_name_attr
                level = self.ClassLevel.NO_LEVEL
            cur_class = Class(class_code=class_code, class_num=class_num, practitioner=practitioner, level=level)
            if class_code not in self.classes_details.keys():
                self.classes_details[class_code] = dict()
            self.classes_details[class_code][class_num] = cur_class
        for class_code, class_numbers_dict in self.classes_details.items():
            archives_class_num = max(class_numbers_dict.keys())
            class_numbers_dict[archives_class_num].practitioner = ''
            class_numbers_dict[archives_class_num].level = self.ClassLevel.ARCHIVES

    def _get_class_details(self, class_code: str, class_num: int) -> Class:
        self.assert_logged_in()
        assert type(class_code) == str, 'Class code must be a string!'
        assert type(class_num) == int, 'Class number must be an integer!'
        return self.classes_details.get(class_code, {}).get(class_num, None)

    def get_class_level(self, class_code: str, class_num: int) -> str:
        class_details = self._get_class_details(class_code, class_num)
        if class_details is None:
            return self.ClassLevel.NO_LEVEL
        return class_details.level

    def get_class_practitioner(self, class_code: str, class_num: int) -> str:
        class_details = self._get_class_details(class_code, class_num)
        if class_details is None:
            return ''
        return class_details.practitioner

    def get_organic_teacher_name(self, class_code: str, class_num: int) -> str:
        phonebook_df = self.get_students_phonebook(class_code)
        if phonebook_df.empty:
            return ''
        class_code_filter = phonebook_df['class_code'] == class_code
        class_num_filter = phonebook_df['class_num'] == class_num
        teachers_df = phonebook_df.loc[class_code_filter & class_num_filter, 'original_teacher']
        return str(teachers_df.mode().head(1).item())

    def get_num_of_active_classes(self, class_code: str):
        class_details = self.classes_details.get(class_code, {})
        return max(class_details.keys()) - 1  # highest class is archives, not active
