# 4117-833-5
from math import trunc as tr


def heart_rate(y):
    max_heart_rate = float(208-0.7*y)
    return max_heart_rate


def main_function():
    for x in range(20, 62, 2):
        print("Maximum heart rate at age:", x, "is", tr(heart_rate(x)), "beats per minute.")


main_function()
