''''''''''
def MyDecorator(Somefunction):

    def adding():
        print("This is the added text to the 'actual function'")
        Somefunction()

    return adding()


@MyDecorator
def actualfunction():
    print("This is the actual function.")
'''''''''

'''''

def mathdecorator(mathsfunction):
    def addition(x, y):
        print("First some addition...")
        z = x + y
        print(z)
        mathsfunction(x, y)

    def division(x, y):
        print("First some division...")
        z = float(x//y)
        print(z)
        mathsfunction(x, y)

    return addition and division



@mathdecorator
def multiplication(x, y):
    print("Now for some multiplication...")
    z = x*y
    print(z)


multiplication(5, 10)
'''''

class my_decorator(object):     # Any classes used as decorators must implement the '__call__' function in order that the result the decorator creates is again a callable function

    def __init__(self, f):
        print("inside my_decorator init function")
        self.f = f

    def __call__(self):
        print("Entering self.f.__name__")
        self.f()
        print("Exited")


@my_decorator
def Function1():
    print("inside Function1()")

@my_decorator
def function2():
    print("Inside function2()")


Function1()     # Calling the function after it has been decorated
function2()     # Calling the function after it has been decorated




































