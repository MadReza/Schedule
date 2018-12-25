from datetime import datetime
import calendar
from time import strptime, strftime, mktime
import re   #text splitting
from xml.etree.ElementTree import XML
import zipfile
from teacher import Teacher, Segment


"""
http://officeopenxml.com/WPtable.php
"""

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
BODY = WORD_NAMESPACE + 'body'
TABLE = WORD_NAMESPACE + 'tbl'
ROW = WORD_NAMESPACE + 'tr'
COL = WORD_NAMESPACE + 'tc'
CELL = WORD_NAMESPACE + 'tbl'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'

def extract_text(element):
    text = ""
    for paragraph in element.iter(PARA):
        for i in paragraph.itertext():
            text += i
        text += " "
    return text.strip()

def convert_str_date(txt):
    datetime_object = datetime.strptime("" + str(datetime.now().year) + " " + txt, '%Y %B %d') #https://docs.python.org/3/library/datetime.html#strftime-and-strptime-behavior
    if datetime_object.month < datetime.now().month:
        year = datetime.now().year + 1
        datetime_object = datetime(year=year, month=datetime_object.month, day=datetime_object.day)
    return datetime_object

def extract_date(element):
    dates = []
    for paragraph in element.iter(PARA):
        text = ""
        for i in paragraph.itertext():
            text += i
        split = re.split('; |, | ', text.strip())

        for nums in split[1:]:
            d = split[0] + " " + nums
            dates.append(convert_str_date(d))
            #dates.append(split[0] + " " +  nums)
    return dates

def extract_time(element):
    times = {}
    t = extract_text(element)
    split = re.split('; |to|\xa0|:| ', t.strip())
    split = list(filter(None, split)) #Remove Empty from list
    if len(split) == 0:
        return times

    #TODO: do checks for the time to be more exact. As in 8 pm and 8 am....
    if int(split[0]) < 8 or int(split[0]) == 12:
        start = strptime(split[0] + ":" + split[1] + " PM", "%I:%M %p")
    else:
        start = strptime(split[0] + ":" + split[1] + " AM", "%I:%M %p")
    if int(split[2]) < 8 or int(split[2]) == 12:
        end = strptime(split[2] + ":" + split[3] + " PM", "%I:%M %p")
    else:
        end = strptime(split[2] + ":" + split[3] + " AM", "%I:%M %p")
    times['start'] = start
    times['end'] = end
    #times['length'] = end.time() - start.time()
    #print(times['length'])
    sh = start.tm_hour * 60 + start.tm_min
    eh = end.tm_hour * 60 + end.tm_min
    times['length'] = eh-sh
    return times

def extract_table_data(path):
    """
    Take the path of a docx file as argument, return a list of times in dictionary format. teacher is the key
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)
    table = tree.find(BODY).find(TABLE)
    rows = table.findall(ROW)

    day = ""
    dates = ""
    teachers = {}

    for row in rows[1:]:
        cols = row.findall(COL)

        #Code
        code = extract_text(cols[0])

        #course
        course = extract_text(cols[1])

        #day
        dayTemp = extract_text(cols[2])
        if dayTemp != "":
            day = dayTemp

        #Calendar/Date
        temp_date = extract_date(cols[3])
        if len(temp_date) != 0:
            dates = temp_date

        #hour
        times = extract_time(cols[4])

        #Teachers
        teacher = extract_text(cols[5])

        if teacher == "":
            continue

        if teacher not in teachers:
            teachers[teacher] = Teacher(teacher)

        segment = Segment(code, course, day, dates, times)

        teachers[teacher].add_segment(segment)

    return teachers

def print_all(data):
    for teacher in data.values():
        print("Teacher: ", teacher.name)
        #print("\tCode: ", teacher.segments[0].code)
        #print("\tCourse: ", teacher.course)
        print("\tTotal Time: ", teacher.total_time / 60, "hours")
        print("\tCourses:")
        for course in teacher.courses:
            print("\t\t", course, teacher.course_time[course] / 60, "hours")
            for segment in teacher.courses[course]:
                for date in segment.dates:
                    d = date.strftime("%A, %d. %B %Y ")
                    s = strftime("%I:%M%p", segment.times['start'] )
                    e = strftime("%I:%M%p", segment.times['end'])
                    t = segment.times['length']
                    print("\t\t\t", d, s, "to", e, "total:", t, "minutes")#%I:%M%p

        print("#####")

data = extract_table_data("Schedule.docx"); #Issues with full path .....
print_all(data)
