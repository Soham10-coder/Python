print("Enter DAA Mark:")
daa = int(input())

print("Enter AT Mark:")
at = int(input())

print("Enter CN Mark:")
cn = int(input())

total = daa + at + cn
percent = (total*100)/300

print("Total percentage: ",int(percent),"%")

if(percent>=90):
    print("Student pass with distinction")
elif(percent>=75 and percent<90):
    print("Student pass with A grade")
elif(percent>=50 and percent<75):
    print("Student pass with B grade")
elif(percent>=35 and percent<50):
    print("Student pass with C grade")
else:
    print("Student failed")