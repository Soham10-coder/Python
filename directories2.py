import os
os.mkdir("folder")
print("Directory 'folder' created.")

# list file 
import os
for item in os.listdir("."):
    print(item)

# remove directory
import os 
os.mkdir("folder2")
print("Directory 'folder2' removed.")


#check if directory exists
import os
if os.path.exists("folder2"):
    print("Directory 'folder1' exists.")
else:
     print("Directory 'folder1' does not exists.")


#rename 
import os 
os.rename("folder2","renamed_folder")
print("Directory renamed to 'renamed_folder'.")