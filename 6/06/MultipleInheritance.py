class Mother1:
    mothername1 = ""
    def mother1(self):
        print(self.mothername1)

class Father1:
    fathername1 = ""
    def father1(self):
        print(self.fathername1)

class Son1(Mother1, Father1):
    def parents1(self):
        print ("Father name is :", self.fathername1)
        print ("Mother name is :", self.mothername1)
s1 = Son1()
s1.fathername1 = "Rajesh"
s1.mothername1 = "Shreya"
s1.parents1()