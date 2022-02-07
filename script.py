import gspread
import operator
from oauth2client.service_account import ServiceAccountCredentials
import time
import constant

scopes = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
          "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

month_list = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
cred = ServiceAccountCredentials.from_json_keyfile_name("cred.json", scopes)
client = gspread.authorize(cred)
dashboard = client.open(constant.SHEET_NAME).get_worksheet_by_id(constant.DASHBOARD_SHEET_ID)


def populate_data(worksheet_id, topic_rows, month_rows, dashboard_month_row, dashboard_month_column):
    sheet = client.open(constant.SHEET_NAME).get_worksheet_by_id(worksheet_id)

    topic_count = {}
    month = 'Jan'
    topic_count.__setitem__(month, [])
    months_row = month_rows
    topics_row = topic_rows
    empty_question_cell = 0
    question_row = 2

    while empty_question_cell < 10:

        month_value = sheet.cell(months_row, 2).value
        if month_value:
            empty_question_cell = 0
            month = month_value.split('-')
            month = month[1]

            if topic_count.__contains__(month):
                topic = sheet.cell(topics_row, 4).value
                if topic:
                    # print(topic)
                    topic_count.get(month).append(topic)
            else:
                topic = sheet.cell(topics_row, 4).value
                topic_count.__setitem__(month, [])
                if topic:
                    # print(topic)
                    topic_count.get(month).append(topic)

        else:
            print(topic_count.get(month))
            topic = sheet.cell(topics_row, 4).value
            print(topic)
            question =sheet.cell(question_row, 5).value
            if question is None:
                empty_question_cell = empty_question_cell + 1
                print(empty_question_cell)
            else:
                empty_question_cell=0
                if topic_count.__contains__(month):
                    if topic:
                        topic_count.__getitem__(month).append(topic)
                else:
                    if topic:
                        topic_count.__setitem__(month,topic)



        months_row = months_row + 1
        topics_row = topics_row + 1
        question_row = question_row + 1

        # print(topic_count)
        time.sleep(2)
    col = dashboard_month_column

    row = dashboard_month_row
    for months in month_list:

        topic_counts = {}

        if topic_count.__contains__(months):
            topics = topic_count.get(months)
            print(topics)
            for topic in topics:
                if topic_counts.__contains__(topic):
                    topic_counts.__setitem__(topic, topic_counts.get(topic) + 1)
                else:
                    topic_counts.__setitem__(topic, 1)
            sorted_topic = dict(sorted(topic_counts.items(), key=operator.itemgetter(1), reverse=True))
            print(sorted_topic)
            i = 0

            for k, v in sorted_topic.items():
                dashboard.update_cell(row, col, k)
                i += 1
                col = 1 + col
                if i >= 5:
                    break
        col = dashboard_month_column
        row = row + 1
        print(row)
