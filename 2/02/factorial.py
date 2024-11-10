num=int(input("Enter Number : "))

if num<=0:
    print("Not found")
else :
    fact=1
    for i in range(1,num+1):
        fact=fact*i
        
    print(fact)
