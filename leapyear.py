print("Enter the year:")
year = int(input())

if(year%4==0 and year%100!=0):
    print("Year is leap year")
elif(year%100==0 and year%400==0):
    print("Year is leap year")
else:
    print("Year is not a leap year")