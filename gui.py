from PyQt5.QtCore import QThread, pyqtSignal, pyqtSlot, QObject
from dataframe_to_excel import MashovReportsToExcel
from datetime import date, datetime, timedelta
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import Qt
import traceback
import requests
import json
import sys
import os

DEBUG = False


# class for scrollable label
class ScrollLabel(QtWidgets.QScrollArea):

    # contructor
    def __init__(self, *args, **kwargs):
        # making widget resizable
        super().__init__(*args, **kwargs)
        self.setWidgetResizable(True)

        # making qwidget object
        content = QtWidgets.QWidget(self)
        self.setWidget(content)

        # vertical box layout
        lay = QtWidgets.QVBoxLayout(content)

        # creating label
        self.label = QtWidgets.QLabel(content)

        # setting alignment to the text
        self.label.setAlignment(Qt.AlignHCenter)

        # making label multi-line
        self.label.setWordWrap(True)

        # adding label to the layout
        lay.addWidget(self.label)

        # the setText method

    def setText(self, text):
        # setting text to the label
        self.label.setText(text)


class UiMainWindow:
    HEB_YEARS = (
        'תשפ', 'תשפא', 'תשפב', 'תשפג', 'תשפד', 'תשפה', 'תשפו', 'תשפז', 'תשפח', 'תשפט',
        'תשצ', 'תשצא', 'תשצב', 'תשצג', 'תשצד', 'תשצה', 'תשצו', 'תשצז', 'תשצח', 'תשצט', 'תשק'
    )

    @staticmethod
    def get_first_day_of_prev_week() -> date:
        curr_date = datetime.now().date()
        return curr_date - timedelta(days=(curr_date.weekday() + 1) % 7) - timedelta(days=7)

    @staticmethod
    def get_last_day_of_prev_week():
        return UiMainWindow.get_first_day_of_prev_week() + timedelta(days=6)

    def __init__(self):
        self.central_widget = None
        self.gridLayout = None
        self.header_label = None
        self.fetch_from_server_label = None
        self.class_code_10_checkbox = None
        self.class_code_11_checkbox = None
        self.class_code_12_checkbox = None
        self.year_combobox = None
        self.year_combobox = None
        self.from_date_label = None
        self.to_date_label = None
        self.summary_checkbox = None
        self.summary_from_date_picker = None
        self.summary_to_date_picker = None
        self.mashov_checkbox = None
        self.mashov_from_date_picker = None
        self.mashov_to_date_picker = None
        self.periodical_checkbox = None
        self.periodical_from_date_picker = None
        self.periodical_to_date_picker = None
        self.error_label = None
        self.submit_btn = None
        self.class_code_label = None
        self.year_label = None
        self.summary_button_group = None
        self.summary_week_btn = None
        self.summary_year_btn = None
        self.summary_between_date_btn = None
        self.mashov_button_group = None
        self.mashov_week_btn = None
        self.mashov_year_btn = None
        self.mashov_between_date_btn = None
        self.periodical_button_group = None
        self.periodical_week_btn = None
        self.periodical_year_btn = None
        self.periodical_between_date_btn = None
        self.credentials_edit_checkbox = None
        self.username_edit_text = None
        self.password_edit_text = None
        self.credentials_submit_btn = None
        self.open_destination_folder_btn =None
        self._translate = QtCore.QCoreApplication.translate

        self.destination_folder_path = os.path.expanduser('~')
        self.report_maker_thread = None
        self.report_maker_async = None
        if not os.path.exists(CreateReportsAsync.CREDENTIALS_PATH):
            credentials = {
                'username': '',
                'password': ''
            }
            with open(CreateReportsAsync.CREDENTIALS_PATH, 'w') as fp:
                json.dump(credentials, fp)

    def enable_edit_credentials(self):
        self.username_edit_text.setEnabled(self.credentials_edit_checkbox.isChecked())
        self.password_edit_text.setEnabled(self.credentials_edit_checkbox.isChecked())
        self.credentials_submit_btn.setEnabled(self.credentials_edit_checkbox.isChecked())
        self.credentials_edit_checkbox.setEnabled(not self.credentials_edit_checkbox.isChecked())

    def credentials_submit(self):
        credentials = {
            'username': self.username_edit_text.text(),
            'password': self.password_edit_text.text()
        }
        with open(CreateReportsAsync.CREDENTIALS_PATH, 'w') as fp:
            json.dump(credentials, fp)
        self.credentials_edit_checkbox.setChecked(False)
        self.enable_edit_credentials()

    def setupUi(self, main_window):
        main_window.setObjectName("MainWindow")
        self.setWindowIcon(QtGui.QIcon('mashov_icon.png'))
        main_window.showMaximized()
        font = QtGui.QFont()
        font.setPointSize(15)
        main_window.setFont(font)
        main_window.setLayoutDirection(QtCore.Qt.RightToLeft)
        self.central_widget = QtWidgets.QWidget(main_window)
        self.central_widget.setObjectName("central_widget")
        self.gridLayout = QtWidgets.QGridLayout(self.central_widget)
        self.gridLayout.setContentsMargins(10, 10, 10, 10)
        self.gridLayout.setObjectName("gridLayout")

        self.header_label = QtWidgets.QLabel(self.central_widget)
        self.header_label.setAlignment(QtCore.Qt.AlignCenter)
        self.header_label.setObjectName("header_label")

        self.credentials_edit_checkbox = QtWidgets.QCheckBox(self.central_widget)
        self.credentials_edit_checkbox.setChecked(False)
        self.credentials_edit_checkbox.stateChanged.connect(self.enable_edit_credentials)
        self.username_edit_text = QtWidgets.QLineEdit(self.central_widget)
        self.password_edit_text = QtWidgets.QLineEdit(self.central_widget)
        self.password_edit_text.setEchoMode(QtWidgets.QLineEdit.Password)
        self.credentials_submit_btn = QtWidgets.QPushButton(self.central_widget)
        self.credentials_submit_btn.setFlat(False)
        self.credentials_submit_btn.setObjectName("credentials_submit_btn")
        self.credentials_submit_btn.clicked.connect(self.credentials_submit)
        self.enable_edit_credentials()

        self.fetch_from_server_label = QtWidgets.QLabel(self.central_widget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.fetch_from_server_label.setFont(font)
        self.fetch_from_server_label.setText("")
        self.fetch_from_server_label.setObjectName("fetch_from_server_label")

        self.class_code_label = QtWidgets.QLabel(self.central_widget)
        self.class_code_10_checkbox = QtWidgets.QCheckBox(self.central_widget)
        self.class_code_10_checkbox.setObjectName("class_code_10_checkbox")
        self.class_code_11_checkbox = QtWidgets.QCheckBox(self.central_widget)
        self.class_code_11_checkbox.setObjectName("class_code_11_checkbox")
        self.class_code_12_checkbox = QtWidgets.QCheckBox(self.central_widget)
        self.class_code_12_checkbox.setObjectName("class_code_12_checkbox")
        self.year_label = QtWidgets.QLabel(self.central_widget)
        self.year_combobox = QtWidgets.QComboBox(self.central_widget)
        self.year_combobox.setObjectName("year_combobox")
        for _ in self.HEB_YEARS:
            self.year_combobox.addItem("")

        self.from_date_label = QtWidgets.QLabel(self.central_widget)
        self.from_date_label.setObjectName("from_date_label")
        self.to_date_label = QtWidgets.QLabel(self.central_widget)
        self.to_date_label.setObjectName("to_date_label")

        date_now = datetime.now().date()
        self.summary_checkbox = QtWidgets.QCheckBox(self.central_widget)
        self.summary_checkbox.setObjectName("summary_check_btn")
        self.summary_button_group = QtWidgets.QButtonGroup()
        self.summary_week_btn = QtWidgets.QRadioButton(self.central_widget)
        self.summary_week_btn.clicked.connect(self.summary_radio_clicked)
        self.summary_week_btn.setChecked(True)
        self.summary_button_group.addButton(self.summary_week_btn)
        self.summary_year_btn = QtWidgets.QRadioButton(self.central_widget)
        self.summary_year_btn.clicked.connect(self.summary_radio_clicked)
        self.summary_button_group.addButton(self.summary_year_btn)
        self.summary_between_date_btn = QtWidgets.QRadioButton(self.central_widget)
        self.summary_between_date_btn.clicked.connect(self.summary_radio_clicked)
        self.summary_button_group.addButton(self.summary_between_date_btn)
        self.summary_from_date_picker = QtWidgets.QDateEdit(self.central_widget)
        self.summary_from_date_picker.setDisplayFormat('yyyy/MM/dd')
        self.summary_from_date_picker.setMinimumDate(QtCore.QDate(2019, 1, 1))
        self.summary_from_date_picker.setDate(QtCore.QDate(date_now.year, date_now.month, 1))
        self.summary_from_date_picker.setCalendarPopup(True)
        self.summary_from_date_picker.setObjectName("summary_from_date_picker")
        self.summary_to_date_picker = QtWidgets.QDateEdit(self.central_widget)
        self.summary_to_date_picker.setDisplayFormat('yyyy/MM/dd')
        self.summary_to_date_picker.setMinimumDate(QtCore.QDate(2019, 1, 1))
        self.summary_to_date_picker.setDate(QtCore.QDate(date_now.year, date_now.month, 10))
        self.summary_to_date_picker.setCalendarPopup(True)
        self.summary_to_date_picker.setObjectName("summary_to_date_picker")
        self.summary_radio_clicked()

        self.mashov_checkbox = QtWidgets.QCheckBox(self.central_widget)
        self.mashov_checkbox.setObjectName("mashov_check_btn")
        self.mashov_button_group = QtWidgets.QButtonGroup()
        self.mashov_week_btn = QtWidgets.QRadioButton(self.central_widget)
        self.mashov_week_btn.clicked.connect(self.mashov_radio_clicked)
        self.mashov_week_btn.setChecked(True)
        self.mashov_button_group.addButton(self.mashov_week_btn)
        self.mashov_year_btn = QtWidgets.QRadioButton(self.central_widget)
        self.mashov_year_btn.clicked.connect(self.mashov_radio_clicked)
        self.mashov_button_group.addButton(self.mashov_year_btn)
        self.mashov_between_date_btn = QtWidgets.QRadioButton(self.central_widget)
        self.mashov_between_date_btn.clicked.connect(self.mashov_radio_clicked)
        self.mashov_button_group.addButton(self.mashov_between_date_btn)
        self.mashov_from_date_picker = QtWidgets.QDateEdit(self.central_widget)
        self.mashov_from_date_picker.setDisplayFormat('yyyy/MM/dd')
        self.mashov_from_date_picker.setMinimumDate(QtCore.QDate(2019, 1, 1))
        self.mashov_from_date_picker.setDate(QtCore.QDate(date_now.year, date_now.month, 1))
        self.mashov_from_date_picker.setCalendarPopup(True)
        self.mashov_from_date_picker.setObjectName("mashov_from_date_picker")
        self.mashov_to_date_picker = QtWidgets.QDateEdit(self.central_widget)
        self.mashov_to_date_picker.setDisplayFormat('yyyy/MM/dd')
        self.mashov_to_date_picker.setMinimumDate(QtCore.QDate(2019, 1, 1))
        self.mashov_to_date_picker.setDate(QtCore.QDate(date_now.year, date_now.month, 10))
        self.mashov_to_date_picker.setCalendarPopup(True)
        self.mashov_to_date_picker.setObjectName("mashov_to_date_picker")
        self.mashov_radio_clicked()

        self.periodical_checkbox = QtWidgets.QCheckBox(self.central_widget)
        self.periodical_checkbox.setObjectName("periodical_check_btn")
        self.periodical_button_group = QtWidgets.QButtonGroup()
        self.periodical_week_btn = QtWidgets.QRadioButton(self.central_widget)
        self.periodical_week_btn.clicked.connect(self.periodical_radio_clicked)
        self.periodical_week_btn.setChecked(True)
        self.periodical_button_group.addButton(self.periodical_week_btn)
        self.periodical_year_btn = QtWidgets.QRadioButton(self.central_widget)
        self.periodical_year_btn.clicked.connect(self.periodical_radio_clicked)
        self.periodical_button_group.addButton(self.periodical_year_btn)
        self.periodical_between_date_btn = QtWidgets.QRadioButton(self.central_widget)
        self.periodical_between_date_btn.clicked.connect(self.periodical_radio_clicked)
        self.periodical_button_group.addButton(self.periodical_between_date_btn)
        self.periodical_from_date_picker = QtWidgets.QDateEdit(self.central_widget)
        self.periodical_from_date_picker.setDisplayFormat('yyyy/MM/dd')
        self.periodical_from_date_picker.setMinimumDate(QtCore.QDate(2019, 1, 1))
        self.periodical_from_date_picker.setDate(QtCore.QDate(date_now.year, date_now.month, 1))
        self.periodical_from_date_picker.setCalendarPopup(True)
        self.periodical_from_date_picker.setObjectName("periodical_from_date_picker")
        self.periodical_to_date_picker = QtWidgets.QDateEdit(self.central_widget)
        self.periodical_to_date_picker.setDisplayFormat('yyyy/MM/dd')
        self.periodical_to_date_picker.setMinimumDate(QtCore.QDate(2019, 1, 1))
        self.periodical_to_date_picker.setDate(QtCore.QDate(date_now.year, date_now.month, 10))
        self.periodical_to_date_picker.setCalendarPopup(True)
        self.periodical_to_date_picker.setObjectName("periodical_to_date_picker")
        self.periodical_radio_clicked()

        self.error_label = ScrollLabel(self.central_widget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.error_label.setFont(font)
        self.error_label.setText("")
        self.error_label.setStyleSheet('color: red')

        self.submit_btn = QtWidgets.QPushButton(self.central_widget)
        self.submit_btn.setObjectName("submit_btn")
        self.submit_btn.clicked.connect(self.submit)

        self.open_destination_folder_btn = QtWidgets.QPushButton(self.central_widget)
        self.open_destination_folder_btn.setObjectName("open_destination_folder_btn")
        self.open_destination_folder_btn.clicked.connect(self.open_destination_folder)

        self.gridLayout.addWidget(self.header_label, 1, 0, 1, 6)
        self.gridLayout.addWidget(self.fetch_from_server_label, 2, 0, 1, 6)

        self.gridLayout.addWidget(self.credentials_edit_checkbox, 3, 0, 1, 2)
        self.gridLayout.addWidget(self.username_edit_text, 3, 2, 1, 1)
        self.gridLayout.addWidget(self.password_edit_text, 3, 3, 1, 1)
        self.gridLayout.addWidget(self.credentials_submit_btn, 3, 4, 1, 1)

        self.gridLayout.addWidget(self.class_code_label, 4, 0, 1, 1)
        self.gridLayout.addWidget(self.class_code_10_checkbox, 4, 1, 1, 1)
        self.gridLayout.addWidget(self.class_code_11_checkbox, 4, 2, 1, 1)
        self.gridLayout.addWidget(self.class_code_12_checkbox, 4, 3, 1, 1)
        self.gridLayout.addWidget(self.year_label, 4, 4, 1, 1)
        self.gridLayout.addWidget(self.year_combobox, 4, 5, 1, 1)
        self.gridLayout.addWidget(self.from_date_label, 5, 4, 1, 1)
        self.gridLayout.addWidget(self.to_date_label, 5, 5, 1, 1)
        self.gridLayout.addWidget(self.summary_checkbox, 6, 0, 1, 1)
        self.gridLayout.addWidget(self.summary_week_btn, 6, 1, 1, 1)
        self.gridLayout.addWidget(self.summary_year_btn, 6, 2, 1, 1)
        self.gridLayout.addWidget(self.summary_between_date_btn, 6, 3, 1, 1)
        self.gridLayout.addWidget(self.summary_from_date_picker, 6, 4, 1, 1)
        self.gridLayout.addWidget(self.summary_to_date_picker, 6, 5, 1, 1)
        self.gridLayout.addWidget(self.mashov_checkbox, 7, 0, 1, 1)
        self.gridLayout.addWidget(self.mashov_week_btn, 7, 1, 1, 1)
        self.gridLayout.addWidget(self.mashov_year_btn, 7, 2, 1, 1)
        self.gridLayout.addWidget(self.mashov_between_date_btn, 7, 3, 1, 1)
        self.gridLayout.addWidget(self.mashov_from_date_picker, 7, 4, 1, 1)
        self.gridLayout.addWidget(self.mashov_to_date_picker, 7, 5, 1, 1)
        self.gridLayout.addWidget(self.periodical_checkbox, 8, 0, 1, 1)
        self.gridLayout.addWidget(self.periodical_week_btn, 8, 1, 1, 1)
        self.gridLayout.addWidget(self.periodical_year_btn, 8, 2, 1, 1)
        self.gridLayout.addWidget(self.periodical_between_date_btn, 8, 3, 1, 1)
        self.gridLayout.addWidget(self.periodical_from_date_picker, 8, 4, 1, 1)
        self.gridLayout.addWidget(self.periodical_to_date_picker, 8, 5, 1, 1)
        self.gridLayout.addWidget(self.error_label, 9, 0, 1, 6)
        self.gridLayout.addWidget(self.submit_btn, 10, 0, 1, 3)
        self.gridLayout.addWidget(self.open_destination_folder_btn, 10, 3, 1, 3)
        main_window.setCentralWidget(self.central_widget)
        self.retranslateUi(main_window)
        QtCore.QMetaObject.connectSlotsByName(main_window)

    def retranslateUi(self, main_window):
        _translate = QtCore.QCoreApplication.translate
        main_window.setWindowTitle(_translate("MainWindow", "בית יציב - דוחות משוב"))
        self.header_label.setText(_translate("MainWindow", "הפקת דוחות משוב בית יציב"))
        self.class_code_10_checkbox.setText(_translate("MainWindow", "י"))
        self.class_code_11_checkbox.setText(_translate("MainWindow", "יא"))
        self.class_code_12_checkbox.setText(_translate("MainWindow", "יב"))
        for i, year in enumerate(self.HEB_YEARS):
            self.year_combobox.setItemText(i, _translate("MainWindow", year))
        self.from_date_label.setText(_translate("MainWindow", "מתאריך"))
        self.to_date_label.setText(_translate("MainWindow", "עד תאריך"))
        self.summary_week_btn.setText(_translate("MainWindow", "שבועי"))
        self.summary_year_btn.setText(_translate("MainWindow", "מתחילת שנה"))
        self.summary_between_date_btn.setText(_translate("MainWindow", "בין תאריכים"))
        self.periodical_week_btn.setText(_translate("MainWindow", "שבועי"))
        self.periodical_year_btn.setText(_translate("MainWindow", "מתחילת שנה"))
        self.periodical_between_date_btn.setText(_translate("MainWindow", "בין תאריכים"))
        self.summary_checkbox.setText(_translate("MainWindow", "דוח סיכום"))
        self.mashov_week_btn.setText(_translate("MainWindow", "שבועי"))
        self.mashov_year_btn.setText(_translate("MainWindow", "מתחילת שנה"))
        self.mashov_between_date_btn.setText(_translate("MainWindow", "בין תאריכים"))
        self.mashov_checkbox.setText(_translate("MainWindow", "דוח משוב"))
        self.periodical_checkbox.setText(_translate("MainWindow", "דוח תקופתי"))
        self.submit_btn.setText(_translate("MainWindow", "הפקת דוחות"))
        self.class_code_label.setText(_translate("MainWindow", "שכבה:"))
        self.year_label.setText(_translate("MainWindow", "שנה:"))
        with open(CreateReportsAsync.CREDENTIALS_PATH) as fp:
            credentials = json.load(fp)
        self.credentials_submit_btn.setText(_translate("MainWindow", "אישור"))
        self.username_edit_text.setText(_translate("MainWindow", credentials.get('username', '')))
        self.password_edit_text.setText(_translate("MainWindow", credentials.get('password', '')))
        self.credentials_edit_checkbox.setText(_translate("MainWindow", 'עדכון פרטי התחברות'))
        self.open_destination_folder_btn.setText(_translate("MainWindow", 'פתיחת תיקיית יעד'))

    def open_destination_folder(self):
        if self.destination_folder_path:
            try:
                os.startfile(os.path.join(self.destination_folder_path, MashovReportsToExcel.DESTINATION_FOLDER_NAME))
            except Exception as e:
                print(e)
                self.error_label.setText(self._translate("MainWindow", 'תיקיית יעד לא קיימת או שאין לך הרשאות גישה '
                                                                       'אליה'))

    def summary_radio_clicked(self):
        self.summary_from_date_picker.setEnabled(self.summary_between_date_btn.isChecked())
        self.summary_to_date_picker.setEnabled(self.summary_between_date_btn.isChecked())

    def mashov_radio_clicked(self):
        self.mashov_from_date_picker.setEnabled(self.mashov_between_date_btn.isChecked())
        self.mashov_to_date_picker.setEnabled(self.mashov_between_date_btn.isChecked())

    def periodical_radio_clicked(self):
        self.periodical_from_date_picker.setEnabled(self.periodical_between_date_btn.isChecked())
        self.periodical_to_date_picker.setEnabled(self.periodical_between_date_btn.isChecked())

    def is_any_report_checked(self):
        summary_checked = self.summary_checkbox.isChecked()
        mashov_checked = self.mashov_checkbox.isChecked()
        periodical_checked = self.periodical_checkbox.isChecked()
        return summary_checked or mashov_checked or periodical_checked

    def submit(self):
        self.fetch_from_server_label.setText(self._translate("MainWindow", ''))
        self.error_label.setText(self._translate("MainWindow", ''))
        class_codes = []
        if self.class_code_10_checkbox.isChecked():
            class_codes.append('י')
        if self.class_code_11_checkbox.isChecked():
            class_codes.append('יא')
        if self.class_code_12_checkbox.isChecked():
            class_codes.append('יב')
        year = self.year_combobox.currentText()
        if not class_codes:
            self.error_label.setText(self._translate("MainWindow", 'נא לבחור שכבה אחת לפחות'))
            return
        if not self.is_any_report_checked():
            self.error_label.setText(self._translate("MainWindow", 'נא לבחור דוח אחד לפחות'))
            return
        self.error_label.setText(self._translate("MainWindow", ''))

        all_from_dates = []
        all_to_dates = []
        summary_from_date, summary_to_date = None, None
        mashov_from_date, mashov_to_date = None, None
        periodical_from_date, periodical_to_date = None, None
        summary, mashov, periodical = False, False, False
        if self.summary_checkbox.isChecked():
            summary = True
            if self.summary_between_date_btn.isChecked():
                summary_from_date = self.summary_from_date_picker.date().toPyDate()
                summary_to_date = self.summary_to_date_picker.date().toPyDate()
            elif self.summary_week_btn.isChecked():
                summary_from_date = self.get_first_day_of_prev_week()
                summary_to_date = self.get_last_day_of_prev_week()
            else:
                summary_from_date, summary_to_date = None, None
        if self.mashov_checkbox.isChecked():
            mashov = True
            if self.mashov_between_date_btn.isChecked():
                mashov_from_date = self.mashov_from_date_picker.date().toPyDate()
                mashov_to_date = self.mashov_to_date_picker.date().toPyDate()
            elif self.mashov_week_btn.isChecked():
                mashov_from_date = self.get_first_day_of_prev_week()
                mashov_to_date = self.get_last_day_of_prev_week()
            else:
                mashov_from_date, mashov_to_date = None, None
        if self.periodical_checkbox.isChecked():
            periodical = True
            if self.periodical_between_date_btn.isChecked():
                periodical_from_date = self.periodical_from_date_picker.date().toPyDate()
                periodical_to_date = self.periodical_to_date_picker.date().toPyDate()
            elif self.periodical_week_btn.isChecked():
                periodical_from_date = self.get_first_day_of_prev_week()
                periodical_to_date = self.get_last_day_of_prev_week()
            else:
                periodical_from_date, periodical_to_date = None, None
        all_from_dates.append(summary_from_date)
        all_to_dates.append(summary_to_date)
        all_from_dates.append(mashov_from_date)
        all_to_dates.append(mashov_to_date)
        all_from_dates.append(periodical_from_date)
        all_to_dates.append(periodical_to_date)
        from_dates_without_none = [d for d in all_from_dates if d]
        to_date_without_none = [d for d in all_to_dates if d]
        min_from_date = min(from_dates_without_none) if from_dates_without_none else None
        max_to_date = max(to_date_without_none) if to_date_without_none else None

        # Create worker and thread to save the data
        self.report_maker_async = CreateReportsAsync(
            year=year,
            class_codes=class_codes,
            destination_folder_path=self.destination_folder_path,
            min_from_date=min_from_date,
            max_to_date=max_to_date,
            summary_from_date=summary_from_date,
            summary_to_date=summary_to_date,
            mashov_from_date=mashov_from_date,
            mashov_to_date=mashov_to_date,
            periodical_from_date=periodical_from_date,
            periodical_to_date=periodical_to_date,
            summary=summary,
            mashov=mashov,
            periodical=periodical
        )
        self.report_maker_thread = QThread()
        self.report_maker_thread.setObjectName('report_maker_thread')
        self.report_maker_async.moveToThread(self.report_maker_thread)

        # Connect signals
        self.report_maker_async.sig_done.connect(self.on_report_creation_done)
        self.report_maker_async.sig_msg.connect(print)
        self.report_maker_async.sig_update_error.connect(self.on_error_occur)
        self.report_maker_async.sig_update_fetch.connect(self.on_fetch_data)
        self.report_maker_thread.started.connect(self.report_maker_async.create_reports)
        self.report_maker_thread.start()
        self.set_enable_ui_buttons(False)

    def set_enable_ui_buttons(self, enable: bool):
        self.class_code_10_checkbox.setEnabled(enable)
        self.class_code_11_checkbox.setEnabled(enable)
        self.class_code_12_checkbox.setEnabled(enable)
        self.year_combobox.setEnabled(enable)
        self.summary_checkbox.setEnabled(enable)
        self.summary_week_btn.setEnabled(enable)
        self.summary_year_btn.setEnabled(enable)
        self.summary_between_date_btn.setEnabled(enable)
        self.summary_from_date_picker.setEnabled(enable)
        self.summary_to_date_picker.setEnabled(enable)
        self.mashov_checkbox.setEnabled(enable)
        self.mashov_week_btn.setEnabled(enable)
        self.mashov_year_btn.setEnabled(enable)
        self.mashov_between_date_btn.setEnabled(enable)
        self.mashov_from_date_picker.setEnabled(enable)
        self.mashov_to_date_picker.setEnabled(enable)
        self.periodical_checkbox.setEnabled(enable)
        self.periodical_week_btn.setEnabled(enable)
        self.periodical_year_btn.setEnabled(enable)
        self.periodical_between_date_btn.setEnabled(enable)
        self.periodical_from_date_picker.setEnabled(enable)
        self.periodical_to_date_picker.setEnabled(enable)
        self.submit_btn.setEnabled(enable)
        self.credentials_edit_checkbox.setEnabled(enable)
        self.username_edit_text.setEnabled(enable)
        self.password_edit_text.setEnabled(enable)
        self.credentials_submit_btn.setEnabled(enable)
        if enable:
            self.summary_radio_clicked()
            self.mashov_radio_clicked()
            self.periodical_radio_clicked()

    @pyqtSlot()
    def on_report_creation_done(self):
        self.set_enable_ui_buttons(True)
        print(f'Finished reports creation')
        if self.report_maker_thread:
            self.report_maker_thread.terminate()
            self.report_maker_thread = None
        self.summary_radio_clicked()
        self.mashov_radio_clicked()
        self.periodical_radio_clicked()
        self.enable_edit_credentials()

    @pyqtSlot(str)
    def on_error_occur(self, error_msg):
        self.error_label.setText(error_msg)

    @pyqtSlot(str)
    def on_fetch_data(self, msg):
        self.fetch_from_server_label.setText(msg)


class CreateReportsAsync(QObject):
    sig_done = pyqtSignal()  # worker id: emitted at end of work()
    sig_msg = pyqtSignal(str)  # message to be shown to user
    sig_update_error = pyqtSignal(str)
    sig_update_fetch = pyqtSignal(str)
    CREDENTIALS_PATH = 'credentials.json'

    @staticmethod
    def get_exception_msg(e_msg):
        msg = f'שגיאה:\n{e_msg}'
        if DEBUG:
            msg = f'{msg}\n{traceback.format_exc()}'
        return msg

    def __init__(self, year: str, class_codes: list, destination_folder_path: str, min_from_date: date,
                 max_to_date: date, summary_from_date: date = None, summary_to_date: date = None,
                 mashov_from_date: date = None, mashov_to_date: date = None, periodical_from_date: date = None,
                 periodical_to_date: date = None, summary=False, mashov=False, periodical=False):
        super().__init__()
        self.max_to_date = max_to_date
        self.min_from_date = min_from_date
        self.periodical_to_date = periodical_to_date
        self.periodical_from_date = periodical_from_date
        self.mashov_to_date = mashov_to_date
        self.mashov_from_date = mashov_from_date
        self.summary_to_date = summary_to_date
        self.summary_from_date = summary_from_date
        self.year = year
        self.class_codes = class_codes
        self.destination_folder_path = destination_folder_path
        self.summary = summary
        self.mashov = mashov
        self.periodical = periodical

        assert os.path.exists(self.CREDENTIALS_PATH), "קובץ קונפיגורציה חסר!"
        with open(self.CREDENTIALS_PATH) as credentials_json:
            credentials = json.load(credentials_json)
        self.username = credentials.get('username', '')
        self.password = credentials.get('password', '')

    @pyqtSlot()
    def create_reports(self):
        thread_name = QThread.currentThread().objectName()
        thread_id = int(QThread.currentThreadId())
        self.sig_msg.emit(f'Fetching data from thread "{thread_name}" (#{thread_id})')
        self.sig_update_error.emit('מוריד מידע מהשרת, נא להמתין...')
        try:
            report_writer = MashovReportsToExcel(
                self.year, self.class_codes, self.username, self.password, self.destination_folder_path,
                from_date=self.min_from_date, to_date=self.max_to_date)
            self.sig_update_error.emit('מפיק דוחות...')
            if self.summary:
                self.sig_msg.emit(f'Create summary report from thread "{thread_name}" (#{thread_id})')
                if self.summary_from_date and self.summary_to_date:
                    report_writer.write_summary_report(self.summary_from_date, self.summary_to_date)
                else:
                    report_writer.write_summary_report(report_writer.from_date, report_writer.to_date)
            if self.mashov:
                self.sig_msg.emit(f'Create Mashov report from thread "{thread_name}" (#{thread_id})')
                if self.mashov_from_date and self.mashov_to_date:
                    report_writer.write_mashov_report(self.mashov_from_date, self.mashov_to_date)
                else:
                    report_writer.write_mashov_report(report_writer.from_date, report_writer.to_date)
            if self.periodical:
                self.sig_msg.emit(f'Create periodical report from thread "{thread_name}" (#{thread_id})')
                if self.periodical_from_date and self.periodical_to_date:
                    report_writer.write_periodical_report(self.periodical_from_date, self.periodical_to_date)
                else:
                    report_writer.write_periodical_report(report_writer.from_date, report_writer.to_date)
            if not self.min_from_date:
                self.min_from_date = report_writer.from_date
            if not self.max_to_date:
                self.max_to_date = report_writer.to_date
            report_writer.write_raw_behavior_report(self.min_from_date, self.max_to_date)
            self.sig_update_error.emit('הדוחות הופקו בהצלחה!')
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout, requests.exceptions.HTTPError):
            self.sig_update_error.emit(self.get_exception_msg('לא קיים חיבור לאינטרנט או שקיימת בעיה באתר משוב'))
        except Exception as e:
            self.sig_update_error.emit(self.get_exception_msg(e))
        self.sig_done.emit()


class MainWindow(QtWidgets.QMainWindow, UiMainWindow):
    def __init__(self, *args, obj=None, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        self.setupUi(self)


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    app.exec()


if __name__ == '__main__':
    main()
