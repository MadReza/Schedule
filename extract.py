from datetime import datetime
from xml.etree.ElementTree import XML
import zipfile


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

def extract_table_data(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)
    table = tree.find(BODY).find(TABLE)
    rows = table.findall(ROW)

    day = ""
    date = ""

    for row in rows:
        cols = row.findall(COL)

        #Code
        code = extract_text(cols[0])
        #print(code)

        #course
        course = extract_text(cols[1])
        #print(course)

        #day
        dayTemp = extract_text(cols[2])
        if dayTemp != "":
            #datetime_object = datetime.strptime(dayTemp, '%b %d %Y %I:%M%p')
            day = dayTemp
            print(day)
            #print(datetime_object)
        #print(day)

        #Calendar/Date
        dateTemp = extract_text(cols[3])
        if dateTemp != "":
            date = dateTemp
        #print(date)

        #hour
        hour = extract_text(cols[4])
        #print(hour)

        #Teachers
        teacher = extract_text(cols[5])
        #print(teacher)

        if teacher == "":
            continue

        print(code, course, day, date, hour, teacher)

    print(len(rows))
    print("$$$$$$$$$$$")

    return '\n\n'.join("paragraphs")

tree = extract_table_data("Schedule.docx"); #Issues with full path .....
