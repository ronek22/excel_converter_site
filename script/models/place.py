class Place:
    def __init__(self, name, color):
        self.name = name
        self.color = color

    def get_name(self):
        return self.name

    def get_color(self):
        return self.color

    def __eq__(self, other):
        return self.name == other.name
