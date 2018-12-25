from time import strftime

class Teacher:
    def __init__(self, name):
        self.name = name
        self.courses = {}
        self.course_time = {}
        self.total_time = 0

    def add_segment(self, segment):
        if segment.course not in self.courses:
            self.courses[segment.course] = []
            self.course_time[segment.course] = 0
        self.courses[segment.course].append(segment)
        self.course_time[segment.course] += segment.times['length'] * len(segment.dates)
        self.total_time += segment.times['length'] * len(segment.dates)

    def print(self):
        print("Teacher: ", self.name)
        print("\tTotal Time: ", self.total_time / 60, "hours")
        print("\tCourses:")
        for course in self.courses:
            print("\t\t", course, self.course_time[course] / 60, "hours")
            for segment in self.courses[course]:
                for date in segment.dates:
                    d = date.strftime("%A, %d. %B %Y ")
                    s = strftime("%I:%M%p", segment.times['start'] )
                    e = strftime("%I:%M%p", segment.times['end'])
                    t = segment.times['length']
                    print("\t\t\t", d, s, "to", e, "total:", t, "minutes")#%I:%M%p


class Segment:
    def __init__(self, code, course, day, dates, times):
        self.code = code
        self.course = course
        self.day = day
        self.dates = dates
        self.times = times
