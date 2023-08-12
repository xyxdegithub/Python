class Animals:
    def eat(self):
        print("eat food")
    def move(self):
        print("moving")
    def breathe(self):
        print("breath air")
#哺乳动物的父类是动物
class Mammals(Animals):
    def breastfeed(self):
        print("feed young")
#狗的父类是哺乳动物
class Dog(Mammals):
    def __init__(self,age):
        self.age=age
    def bark(self):
        print("bark bark")


dog=Dog(2)
dog.bark()
dog.eat()

