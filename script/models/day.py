class Day:

    def __init__(self, time, hours, place, week):
        if time == "":
            self.hours = ''
            self.place = ''
            self.time = ''
        else:
            self.time = time
            self.hours = hours
            self.place = place

        self.week = week