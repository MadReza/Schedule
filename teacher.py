class Teacher:
    def __init__(self, name, course):
        self.name = name
        self.course = course
        self.segments = []
        self.total_time = 0

    def add_segment(self, segment):
        self.total_time += segment.times['length'] * len(segment.dates)
        self.segments.append(segment)

class Segment:
    def __init__(self, code, course, day, dates, times):
        self.code = code
        self.course = course
        self.day = day
        self.dates = dates
        self.times = times
