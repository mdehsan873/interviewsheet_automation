import constant
import topic_rankings as script
import failed_interview as failed

script.populate_data(constant.REACT_NATIVE_WORKSHEET_ID, constant.TOPIC_ROW, constant.REACT_NATIVE_MONTH_ROW,
                     'React Native')
failed.interview_clearance()
script.populate_data(constant.DJANGO_WORKSHEET_ID, constant.TOPIC_ROW, constant.DJANGO_MONTH_ROW,
                     'Django')
script.populate_data(constant.JAVA_WORKSHEET_ID, constant.TOPIC_ROW, constant.JAVA_MONTH_ROW,
                     'Java')

script.populate_data(constant.NODE_JS_WORKSHEET_ID, constant.TOPIC_ROW, constant.NODE_JS_MONTH_ROW,
                     'NodeJS')
script.populate_data(constant.REACT_WORKSHEET_ID, constant.TOPIC_ROW, constant.REACT_MONTH_ROW,
                     'ReactJS')
