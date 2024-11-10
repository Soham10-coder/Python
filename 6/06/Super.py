class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age

    def display(self):
        print(self.name, self.age)

class Student(Person):
    def __init__(self, name, age, dob):
        super().__init__(name, age)  
        self.dob = dob

    def displayInfo(self):
        print(self.name, self.age, self.dob)

obj = Student("Rohit",21,"2003-02-01")
obj.display()      
obj.displayInfo()  

