class Person:
    """
    How look schedule?
    It is list of diaries, that look like this
    [
        {
            time: "9:00-18:00",
            hours: 9,
            place = "focus"
            week: 1
        },
        ...
    ]
    """
    def __init__(self, name):
        self.name = name
        self.schedule = list()

    def add_day(self, day):
        self.schedule.append(day)

    def get_name(self):
        return self.name

    def get_schedule(self):
        return self.name

    def get_hours_of_day(self, day):
        return self.schedule[day].hours

    def get_time_of_day(self, day):
        time = self.schedule[day].time
        return time.replace(".00", "")

    def get_place_of_day(self, day):
        return self.schedule[day].place

    def count_work_in_place(self, place):
        return sum(1 for d in self.schedule if d.get("place") == place)

    def count_hours(self):
        return sum(day['hours'] for day in self.schedule)

    def __eq__(self, other):
        return self.name == other.name
