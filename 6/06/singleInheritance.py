class A:
    def func_1(self):
        print("this function is defined inside the parent class")
class B(A):
    def func_2(self):
        print("This function is defined inside the child class")

object=B()
object.func_1()
object.func_2()