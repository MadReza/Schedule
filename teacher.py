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

class Segment:
    def __init__(self, code, course, day, dates, times):
        self.code = code
        self.course = course
        self.day = day
        self.dates = dates
        self.times = times
