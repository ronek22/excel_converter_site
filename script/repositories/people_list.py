class People:
    def __init__(self):
        self.list = list()

    def add_person(self, place):
        self.list.append(place)

    def get_person(self, person):
        for item in self.list:
            if item == person:
                return item
        raise AttributeError("Nie ma takiej osoby.")

    def get_person_by_id(self, idp):
        return self.list[idp]

    def get_person_name_by_id(self, idp):
        return self.list[idp].get_name()

    def get_people_names(self):
        return [p.get_name() for p in self.list]

    def count(self):
        return len(self.list)
