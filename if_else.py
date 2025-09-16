print("Enter first number")
a = int(input())
print("Enter second number")
b = int(input())
print("Enter third number")
c = int(input())

if(a>0 and b>0 and c>0):
    if(a>b and a>c):
        print("a is greater than b and c")
    elif(b>a and b>c):
        print("b is greater than a and c")
    elif(c>a and c>b):
        print("c is greater than a and b")
else:
    print("Number is negative")
