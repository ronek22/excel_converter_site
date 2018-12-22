class Places:
    def __init__(self):
        self.list = list()

    def add_place(self, place):
        self.list.append(place)

    def get_color(self, place):
        for item in self.list:
            if item.name == place:
                return item.color
        raise AttributeError("Nie ma takiego miejsca.")

    def get_place(self, color):
        if color == 65:
            return "biuro"
        for item in self.list:
            if item.color == color:
                return item.name
        return ""

    def get_places_names(self):
        places_list = [p.get_name() for p in self.list]
        return set(places_list)
