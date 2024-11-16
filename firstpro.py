# first programme in python language
print("hi!..lets learn python language");
a=30.50
b=45
name="Aniket"
result= a+b
print("your name: ", name,"and result is=", result)

age=float(input("enter your age: "))

if age>18:
    print("you are eligible to vote")
else:
    print("you are not eligible to vote")
    
#if elif else statement

a=int(input("enter any value to store in variable a: "))
b=int(input("enter any value to store in variable b: "))

if a>b:
    print( "value of a is greater than b")
elif a==b:
    print(" both value are equal")
else:
    print(" b is greater than a")