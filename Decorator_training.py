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
