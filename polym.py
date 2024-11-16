class Cat:
    def sound(self):
        print("Meow")

class Dog:
    def sound(self):
        print("Bark")

def animal_sound(animal):
    animal.sound()
    
cat=Cat()
dog=Dog()

animal_sound(cat)
animal_sound(dog)

#function polymorphism
def add(a,b):
    return a+b

print(add(3,9))
print(add("Hello ", "World!"))
print(add([2,3],[5,6]))


