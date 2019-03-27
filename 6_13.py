import turtle


def centipede(length, step, life):
    turtle.Pen.penup()
    theta = 0
    dtheta = 1
    for i in range(life):
        turtle.Pen.forward(step)
        turtle.Pen.left(theta)
        theta += dtheta
        turtle.Pen.stamp()
        if i > length:
            turtle.Pen.clearstamps(1)
        if theta > 10 or theta < -10:
            dtheta = -dtheta
        if turtle.Pen.ycor() > 350:
            turtle.Pen.left(30)


def mainprogram():
    turtle.Screen().setworldcoordinates(-400, -400, 400, 400)
    centipede(14, 10, 200)
    turtle.Canvas.exitonclick()


mainprogram()
