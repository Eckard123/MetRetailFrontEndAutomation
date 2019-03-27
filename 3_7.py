# 4117-833-5


# A function that takes two values and calculates the average between them
def avg(val1, val2):
    average = (val1 + val2)/2
    return average


# The main function that executes the necessary code
def mainfunction():
    value1 = float(input("Please enter your first number: "))
    value2 = float(input("Please enter your second number: "))
    print("The average between the numbers you entered is: ", float(avg(value1, value2)))


mainfunction()
