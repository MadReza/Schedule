import sys, getopt
import os
from extract import extract_table_data
from calendarHelper import create_event, create_calendar, get_calendars
from time import strftime

def print_all(data):
    for teacher in data.values():
        teacher.print()
        print("############################")

def get_or_create_calendar(name, description):
    """
    create or get calendar and return ID
    """
    c = get_calendars()

    if "GIM" not in c:
        return create_calendar(name, description)["id"]
    else:
        return c["GIM"]

def create_calendar_events(calendar, group, num, data):

    for course in data.courses:
        summary = group + " " + num + ": " + course
        location = "1001 Rue Sherbrooke E, Montréal, QC H2L 1L3, Canada"
        for segment in data.courses[course]:
            description = segment.code
            for date in segment.dates:
                d = date.strftime("%Y-%m-%dT") #'2015-05-28T09:00:00'
                s = strftime("%H:%M:%S", segment.times['start'] )
                e = strftime("%H:%M:%S", segment.times['end'])
                startTime = d + s
                endTime = d + e
                create_event(calendar, summary, location, description, startTime, endTime)

def run(path):
    data = extract_table_data(path); #Issues with full path .....
    selected_teacher = ''
    while selected_teacher not in data:
        os.system('clear')
        print("Available Teachers:")
        for teacher in data:
            print("\t",teacher)
        selected_teacher = input('Enter Teacher ("All" for all teachers): ')
        if selected_teacher == "All":
            break

    if selected_teacher == "All":
        print_all(data)
    else:
        data[selected_teacher].print()
        answer = input('Would you like to export to google Calendar?(y or n)')

        if answer == 'y':
            print("Exporting to Google Calendar")
            group = input("What group are they(MAD, CST,...):")
            num = input("What group number are they:")
            print("Finding or creating calendar: GIM")
            id = get_or_create_calendar("GIM", "Cegep GIM courses")
            print("Creating Events")
            create_calendar_events(id, group, num, data[selected_teacher])
            print("DONE")

def main(argv):
    path = ''
    try:
        opts, args = getopt.getopt(argv,"hp:",["path="])
    except getopt.GetoptError:
        print ('schedule.py <PATH_OF_DOCX>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print ('schedule.py <PATH_OF_DOCX>')
            sys.exit()
        elif opt in ("-p", "--path"):
            path = arg

    if path != '':
        run(path)

if __name__ == "__main__":
    main(sys.argv[1:])
