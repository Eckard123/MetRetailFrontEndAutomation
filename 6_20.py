# 4117-833-5


def sort3(x, y, z):

    if x == y and x == z:
        return print("max value=", x, "middle value=", x, "min value=", x)
    if x == y:
        if x > z:
            return print("max value=", x, "middle value=", y, "min value=", z)
        if x < z:
            return print("max value=", z, "middle value=", x, "min value=", y)
    if x == z:
        if x > y:
            return print("max value=", x, "middle value=", z, "min value=", y)
        if x < y:
            return print("max value=", y, "middle value=", x, "min value=", z)
    if y == z:
        if y > x:
            return print("max value=", y, "middle value=", z, "min value=", x)
        if y < x:
            return print("max value=", x, "middle value=", y, "min value=", z)
    if x > y and x > z:
        if y >= z:
            return print("max value=", x, "middle value=", y, "min value=", z)
        if y <= z:
            return print("max value=", x, "middle value=", z, "min value=", y)
    if y > x and y > z:
        if x >= z:
            return print("max value=", y, "middle value=", x, "min value=", z)
        if x <= z:
            return print("max value=", y, "middle value=", z, "min value=", x)
    if z > x and z > y:
        if x >= y:
            return print("max value=", z, "middle value=", x, "min value=", y)
        if x <= y:
            return print("max value=", z, "middle value=", y, "min value=", x)


def mainfunction():

    x = input("First positive integer: ")
    y = input("Second positive integer: ")
    z = input("Third positive integer: ")
    print("Your values in order from greatest to smallest is:")
    return sort3(x, y, z)


mainfunction()



