
class Student():
    def __init__(self, name, score):
        self.name = name
        self.score = score

    def print_score(self):
        print('%s: %s' % (self.name, self.score))

student1 = Student("linwei", 100)
student1.print_score()