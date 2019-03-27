# 4117-833-5

def find_home():

    steps_away = 120

    # From your front door go down the road 20 steps
    for step in range(1, 21):
        steps_away -= 1
    print()
    print("After walking", step, "steps down the road,", "you must turn left or right at the end of the road at the T-junction.")

    # You are now at the T-junction, you can turn left or right, should you turn left or right go 40 steps
    left_or_right = str(input("Will you be turning left or right?  Please type l or r. "))
    flag = False
    option = None
    if left_or_right == 'r':
        flag = True
    if left_or_right == 'l':
        flag = True

    while flag == False:
        print("Invalid option.")
        left_or_right = str(input("Will you be turning left or right?  Please type l or r. "))
        if left_or_right == 'r':
            break
        if left_or_right == 'l':
            break

    if left_or_right == 'l' or 'r':
        if left_or_right == 'l':
            option = "left"
        elif left_or_right == 'r':
            option = "right"
        else:
            option = None

        for steps in range(1, 41):
            steps_away -= 1
        print("You opted to turn", option, "You must now proceed forward fourty steps")

    # If you turned left you must now turn right and walk 20 steps, if you turned right you must now turn left and walk 20 steps
    print()
    if left_or_right == 'l':
        print("After proceeding fourty steps forward You must turn right because you previously turned left.")
        for Steps in range(1, 21):
            steps_away -= 1
        print("Turn right and walk", Steps, "steps", "forward.")
    print()
    if left_or_right == 'r':
        print("After proceeding fourty steps forward You must turn left because you previously turned right.")
        for Steps in range(1, 21):
            steps_away -= 1
        print("Turn left and walk", Steps, "steps", "forward")

    # After turning you have to proceed forward for another 20 steps
    print()
    print("After turning again proceed another twenty steps forward...Almost there.")
    for steps in range(1, 21):
        steps_away -= 1
    print("All in all you still have to go", steps_away, "steps to get home.")

    # Depending whether you turned left or right previously
    if left_or_right == 'l':
        print()
        print("Turn right and proceed forward the last twenty steps.")
        print()
        for steps in range(1, 21):
            steps_away -= 1
        print("You have walked all this way, you walked a total distance of", 120 - steps_away, "steps.")
    if left_or_right == 'r':
        print()
        print("Turn left and proceed forward the last twenty steps.")
        print()
        for steps in range(1, 21):
            steps_away -= 1
        print("You have walked all this way, you walked a total distance of", 120 - steps_away, "steps.")


def mainfunction():
    find_home()


mainfunction()
