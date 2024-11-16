#single inheritance
from abc import ABC, abstractmethod
class Animal:
    def speak(self):
        print("Animal Speak")

    @abstractmethod
    def walk(self):
        pass
    
class Dog(Animal):
    def bark(self):
        print("Dog Bark")
    
    def walk(self):
        print("Dog Walk")
         
dog = Dog()
dog.speak()
dog.bark()
dog.walk()

#multilevel
class Animal:
    def eat(self):
        print("Eating..")
        
class Dog(Animal):
    def bark(self):
        print("Barking...")

class Puppy(Dog):
    def play(self):
        print("Playing...")

puppy=Puppy()
puppy.eat()
puppy.bark()
puppy.play()

#multiple inheritance
class A:
    def method_a(self):
        print("Method of class A")

class B:
    def method_b(self):
        print("Method of class B")

class C(A,B):
    def method_c(self):
        print("Method of class C")
        
obj=C()
obj.method_a()
obj.method_b()
obj.method_c()

#Hierarchical 
class Animal:
    def eat(self):
        print("Eating...")

class Dog(Animal):
    def bark(self):
        print("Barking...")
        
class Cat(Animal):
    def meow(self):
        print("Meowing...")

dog=Dog()
cat=Cat()

dog.eat() 
dog.bark()

cat.eat()
cat.meow()

#super() function
class Parent:
    def show(self):
        print("This is the parent class")
        
class Child(Parent):
    def show(self):
        super().show()
        print("This is the child class")

child=Child()
child.show()   

#super
class Animal:
    def __init__(self,name):
        self.name= name 
    def sound(self):
        print(f"{self.name} makes a sound.")

class Dog(Animal):
    def __init__(self, name,bread):
        super(). __init__(name)
        self.bread=bread
        
    def sound(self):
        super().sound()
        print(f"{self.name} is a {self.bread} and barks.")

dog=Dog("Buddy","Golden Retriever")
dog.sound()

        
        