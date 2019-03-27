# 4117-833-5


def myfactorial(n):
    fact = 1
    if n == 0 or n == 1:
        fact = fact
        print(n, "factorial is:", fact)
    else:
        for x in range(1, n+1):
            fact *= x
            if x == n:
                print(n, "factorial is:", int(fact))


def factorial_table():

    x = int(input("Type in number from where you want it's factorial number: "))
    y = int(input("Type in number to where you want it's factorial number: "))
    for z in range(x, y+1):
        print()
        myfactorial(z)
        squared = z*z
        power = 2**z
        print(z, "squared is:", squared)
        print("2 to the power", z, "is:", power)


def mainfunction():
    factorial_table()


mainfunction()




# One can clearly see from this program that initially functions like powers or an integer squared has a bigger result than that the factorial of a value.
# However, values from 4 and onward has the exact opposite effect.

