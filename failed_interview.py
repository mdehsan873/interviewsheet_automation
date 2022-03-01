import gspread
import operator
from oauth2client.service_account import ServiceAccountCredentials
import time
import constant

scopes = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
          "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

months = {'Jan': 3, 'Feb': 4, 'Mar': 5, 'Apr': 6, 'May': 7, 'Jun': 8, 'July': 9, 'Aug': 10, 'Sep': 11, 'Oct': 12,
          'Nov': 13, 'Dec': 14}
cred = ServiceAccountCredentials.from_json_keyfile_name("cred.json", scopes)
client = gspread.authorize(cred)
dashboard = client.open(constant.SHEET_NAME).get_worksheet_by_id(constant.DASHBOARD_SHEET_ID)
client_interview = client.open(constant.SHEET_NAME).get_worksheet_by_id(constant.CLIENT_INTERVIEW)


def failed_interview(java, react_native, react_js, android, django):
    print("Failed Interview")
    for mon, row in months.items():
        total_interview = len(java[mon]) + len(react_js[mon]) + len(react_native[mon]) + len(android[mon]) + len(
            django[mon])
        failed_interview = 0
        for failed in java[mon]:
            if failed == 'Rejected':
                failed_interview += 1
        for failed in react_js[mon]:
            if failed == 'Rejected':
                failed_interview += 1
        for failed in android[mon]:
            if failed == 'Rejected':
                failed_interview += 1
        for failed in react_native[mon]:
            if failed == 'Rejected':
                failed_interview += 1
        for failed in django[mon]:
            if failed == 'Rejected':
                failed_interview += 1
        percent=(failed/total_interview)*100
        dashboard.update_cell(months[mon] + 30, 7,percent)


def populate(data, add):
    for mon, row in months.items():
        fail = 0
        total = 0
        cleared = 0
        if data.get(mon):
            list = data.get(mon)
            total = len(list)
            for item in list:
                if item == 'Cleared':
                    cleared += 1
                else:
                    fail += 1
        if add == 0:
            dashboard.update_cell(row + add, 2, 'React')
        if add == 12:
            dashboard.update_cell(row + add, 2, 'Django')
        if add == 24:
            dashboard.update_cell(row + add, 2, 'Android')
        if add == 36:
            dashboard.update_cell(row + add, 2, 'NodeJs')
        if add == 48:
            dashboard.update_cell(row + add, 2, 'Java')
        if add == 60:
            dashboard.update_cell(row + add, 2, 'React Native')
        dashboard.update_cell(row + add, 3, total)
        dashboard.update_cell(row + add, 4, fail)
        dashboard.update_cell(row + add, 5, cleared)
        print('done')
        time.sleep(5)


def interview_clearance():
    tech = client_interview.col_values(6)
    dates = client_interview.col_values(5)
    status = client_interview.col_values(9)
    minimum = min(len(tech), min(len(dates), len(status)))
    react_js = {}
    node = {}
    java = {}
    android = {}
    django = {}
    react_native = {}
    print(dates)
    for i in range(0, minimum):
        if tech[i] == 'React':

            date = dates[i].split('-')
            try:
                print(date[1])
                if react_js.__contains__(date[1]):

                    react_js.get(date[1]).append(status[i])
                else:
                    interview_status = [status[i]]
                    react_js.__setitem__(date[1], interview_status)
            except IndexError:
                react_js.__setitem__(None, status[i])
        elif tech[i] == 'Django':

            date = dates[i].split('-')
            try:
                print(date[1])
                if django.__contains__(date[1]):

                    django.get(date[1]).append(status[i])
                else:
                    interview_status = [status[i]]
                    django.__setitem__(date[1], interview_status)
            except IndexError:
                django.__setitem__(None, status[i])
        elif tech[i] == 'NodeJs':

            date = dates[i].split('-')
            try:
                print(date[1])
                if node.__contains__(date[1]):

                    node.get(date[1]).append(status[i])
                else:
                    interview_status = [status[i]]
                    django.__setitem__(date[1], interview_status)
            except IndexError:
                node.__setitem__(None, status[i])
        elif tech[i] == 'Android':

            date = dates[i].split('-')
            try:
                print(date[1])
                if android.__contains__(date[1]):

                    android.get(date[1]).append(status[i])
                else:
                    interview_status = [status[i]]
                    android.__setitem__(date[1], interview_status)
            except IndexError:
                android.__setitem__(None, status[i])
        elif tech[i] == 'Java':

            date = dates[i].split('-')
            try:
                print(date[1])
                if java.__contains__(date[1]):

                    java.get(date[1]).append(status[i])
                else:
                    interview_status = [status[i]]
                    java.__setitem__(date[1], interview_status)
            except IndexError:
                java.__setitem__(None, status[i])
        elif tech[i] == 'React Native':

            date = dates[i].split('-')
            try:
                print(date[1])
                if react_native.__contains__(date[1]):

                    react_native.get(date[1]).append(status[i])
                else:
                    interview_status = [status[i]]
                    react_native.__setitem__(date[1], interview_status)
            except IndexError:
                react_native.__setitem__(None, status[i])
    print(react_js)
    populate(react_js, 0)
    populate(django, 12)
    populate(android, 24)
    populate(node, 36)
    populate(java, 48)
    populate(react_native, 60)
    failed_interview(java, react_native, react_js, android, django)
