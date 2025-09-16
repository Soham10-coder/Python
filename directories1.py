# list files in a directory
import os
path="."
for item in os.listdir():
    print(item)
    
#//create directory
# import os
# os.mkdir("New_Folder")
# print("Directory created.")

import os
print("String format: ",os.getcwd())
print("Byte string format: ",os.getcwd())
