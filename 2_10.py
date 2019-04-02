# Area of a rectangle
# 4117-833-5
from math import trunc as tr

length = float(input("Please enter the length of the rectangle: "))
width = float(input("Please enter the width of the rectangle: "))
area = tr(length*width)
perimeter = tr(2*length + 2*width)

print("The truncated area of the rectangle is: ", area)
print("The truncated perimeter of the rectangle is: ", perimeter)

