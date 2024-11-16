# iterating over array through for loop;
array=[1,2,3,4,5,6,7,8,9]
# for i in array:
#     print(i)

# for i in range(2,9):
#     print(i, i+1)

for i in range(len(array)):
    print(f' element: {array[i]} index: {i}')

# reverse of an array
for i in range(len(array)-1,-1,-1):
   print(f'element: {array[i]}')  
   
for i in range(0,len(array),+2):
   print(f'{array[i]} is odd')
   
# FACTORIAL OF NUMER
num=int(input("enter any nuber to find out its factorial: "))
n=1;
fact=num;
while n!=num:
    fact=fact*n
    n+=1
    
print(f'factorial of {num} is: {fact}')