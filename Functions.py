# Iterative function
# def factorial(n):
#     fact=1;
#     for i in range(1,n+1):
#         fact=fact*i
#     print(fact)
   
# Recurrsion function 
def factorial(n=6):
    fact=1
    if(n==1 or n==0):
        return 1
    else:
        return n*factorial(n-1)
print(factorial())

