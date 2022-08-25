class Person: ## This is parent class
    def __init__(self, fname, lname): # This is init function where you will give properties
        self.firstname = fname
        self.lastname = lname

    def printname(self):
        print(self.firstname, self.lastname)

x = Person("John","Doe") # Person(fname, lname)

#Now create child class

class Student(Person):
    def __init__(self, fname, lname, year): #use pass to keep parent functionalities and properties.
        super().__init__(fname, lname) #use super() to bring parent class properties to child class
        self.graduation = year #Child class can add more properties

    def welcome(self):
        print(f"Welcome {self.firstname} {self.lastname} to the class of year {self.graduation}")
x = Student(input("First Name :"), input("Last Name :"), input("Age : "))
x.welcome() # This will print printname function from parent class 
# because Student is still using parent properties and method.