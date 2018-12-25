import sys, getopt
from extract import extract_table_data, print_all

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
        data = extract_table_data(path); #Issues with full path .....
        print_all(data)

    print ('Path "', path)

if __name__ == "__main__":
    main(sys.argv[1:])
