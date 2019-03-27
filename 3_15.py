# heart.py
# 4117-833-5
from math import trunc as tr


def heart_rate(y):
    max_heart_rate = float(208-0.7*y)
    return max_heart_rate


def main_function():
    age = float(input("Please enter your age in years: "))
    print("Your maximum heart rate is: ", tr(heart_rate(age)), "beats per minute.")


main_function()
