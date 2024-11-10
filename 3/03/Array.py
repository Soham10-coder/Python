#---------------------------------------
import array as arr
a=arr.array('i',[1,2,3])
for i in range(0,3):
    print(a[i],end=" " )
print("\n")
#---------------------------------------
b=arr.array('d',[2.5,3.2,3.3])
for i in range(0,3):
    print(b[i],end=" ")
print("\n")
#---------------------------------------
c=arr.array('i',[1,2,3,4,5])
c.insert(1,4)
for i in range(0,6):
    print(c[i],end=" ")
print("\n")
#---------------------------------------   
print(c.count(4))
print("\n")
#---------------------------------------
print(c.index(4))
#---------------------------------------
c.pop()
for i in range(0,len(c)):
    print(c[i],end=" ")
print("\n")
#---------------------------------------
c.remove(4)
for i in range(0,len(c)):
    print(c[i],end=" ")
print("\n")
#---------------------------------------
c.reverse()
for i in range(0,len(c)):
    print(c[i],end=" ")
print("\n")
#---------------------------------------
c.extend([4,3,4,5,6])
for i in range(0,len(c)):
    print(c[i],end=" ")
print("\n")
#---------------------------------------
c.append(2)
for i in range(0,len(c)):
    print(c[i],end=" ")
print("\n")
#---------------------------------------
