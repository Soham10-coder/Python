print("Enter the number")
n = int(input())
count = 0
for i in range(1,n):
    if(n%i==0):
       count = count+1      
if(count==2):
    print("Number is not prime")
else:
    print("Number is prime")