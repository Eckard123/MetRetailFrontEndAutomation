# 4117-833-5


def conversion():

    cents = float(input("Please enter the amount of cents to be converted to a rand and cent amount.  "))
    if cents < 0:
        cents *= -1

    result1 = cents//100
    result2 = (cents/100) - result1
    result3 = round(result2, 2)
    print(int(result1), "Rand", int(result3*100), "cent")


conversion()

