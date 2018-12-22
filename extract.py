from xml.etree.ElementTree import XML
import zipfile


"""
Module that extract text from MS XML Word document (.docx).
(Inspired by python-docx <https://github.com/mikemaccana/python-docx>)
"""

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'


def get_docx_text(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    startJunk = True
    section = 0 # 0: Code, 1: Course, 2: Day, 3: Calendar/Date, 4: Hour, 5: Teacher
    day = ""    #Cell Merge makes day empty
    date = ""   #Cell Merge makes date empty
    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        #print(type(paragraph))
        value = ""
        for i in paragraph.itertext():
            value += i

        if value == "Teacher":
            startJunk = False
            continue

        if startJunk:   #Skip all Junk
            continue

        #Start Parse
        print(value)

        if section == 0:
            code = value
        elif section == 1:
            course = value
        elif section == 2 and value != "":  #Merge cell show empty
            day = value
        elif section == 3 and value != "":
            date = value
        elif section == 4:
            hour = value
        elif section == 5:
            teacher = value

        section += 1

        if section == 6 and teacher != "":    #ProcessData for none empty lines
            print(code, course, day, date, hour, teacher)

        section = section % 6
        #print(type(texts))
        #print(texts)
        #print("***")


    return '\n\n'.join(paragraphs)

tree = get_docx_text("Schedule.docx"); #Issues with full path .....
