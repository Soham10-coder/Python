# with open("myfile.txt","w") as file1:
file1 = open("myfile.txt","w")
print("File Opened")
    
file1.write("I am Kunal Patil. ")
# file1.close()

file1=open("myfile.txt","at")
file1.write("I am studying at DYPCET")
file1.close()

file1 = open("myfile.txt","rt")
print(file1.read())



