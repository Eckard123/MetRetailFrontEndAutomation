from turtle import *
import turtle as t



# Dear professor, i have tried to adapt this code to get it to work (as you can see from below), but no matter what i do i can't get it to work.
def centipede(length, step, life):
    Turtle.penup()
    theta = 0
    dtheta = 1
    for i in range(life):
        Turtle.Pen.forward()
        Turtle.Pen.left(theta)
        theta += dtheta
        Turtle.Pen.stamp()
        if i > length:
            Turtle.Pen.clearstamps(1)
        if theta > 10:
            if theta < -10:
                dtheta = (dtheta*-1)
        if Turtle.Pen.ycor() > 350:
            Turtle.Pen.left(30)


def mainprogram():
    Turtle.Screen().setworldcoordinates(-400, -400, 400, 400)
    centipede(14, 10, 200)
    t.Canvas().exitonclick()


mainprogram()

