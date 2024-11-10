import re
txt="the rain in spain"
x=re.search("^the.*spain$",txt)
print(x)

#---------------------------------------
txt2="the rain in Solapur"
x=re.findall("ap",txt2)
print(x)

#---------------------------------------
x=re.search("\s",txt2)
print(x)

#---------------------------------------
x=re.split("\s",txt2)
print(x)

#---------------------------------------
x=re.split("\s",txt2,1)
print(x)

#---------------------------------------
x=re.sub("\s","9",txt2)
print(x)

#---------------------------------------
x=re.sub("\s","9",txt2,2)
print(x)

#---------------------------------------
x=re.search(r"\bS\w+",txt2)
print(x.span())

#---------------------------------------
x=re.search(r"\bS\w+",txt2)
print(x.string)

#---------------------------------------
x=re.search(r"\bS\w+",txt2)
print(x.group())

#---------------------------------------
x=re.findall("\Athe",txt2)
print(x)

#---------------------------------------
txt3="the rain in 5 Solapur"
x=re.findall("\d",txt3)
print(x)

#---------------------------------------
x=re.findall("\D",txt3)
print(x)

#---------------------------------------
x=re.findall("\W",txt3)
print(x)

#---------------------------------------
x=re.findall("\w",txt3)
print(x)

#---------------------------------------
x=re.findall("Solapur\Z",txt3)
print(x)


x2=input()
x2=re.findall("Solapur\Z",txt3)
print(x2)

