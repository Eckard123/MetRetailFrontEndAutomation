# 4117-833-5
import time


def home():
    print("You opted to go home.")
    print("You will be coming along the N1-North.  Follow the following instructions to get home.")
    for x in range(1, 5):
        print("Keep heading North on N1")
        time.sleep(0.5)
        if x == 4:
            print("Slower, i think i see the offramp.")

    offramps = ["Garsfontein", "Lynwood", "Atterbury"]

    print()
    print("Yup, looks like the Lynwood offramp, lets take it.")
    print()
    if offramps.index("Lynwood"):
        print("Took the Lynwood offramp, good stuff")
    else:
        print("Keep heading North.")

    # Now you follow the road along until you see house number 641 on your left
    print()
    print("Keep following the road until you see house number 641.")
    for x in range(1, 4):
        if x == 1:
            print()
            print("Slower, we are almost there...")
            time.sleep(1)
        elif x == 2:
            print()
            time.sleep(1.5)
            print("Slower still, you are almost there...")
        elif x == 3:
            print()
            time.sleep(2)
            print("Stop!! You have arrived...")


def menlyn():
    print("You opted to go to Menlyn to do some shopping.")
    print("You will be coming along the N1-North.  Follow the following instructions to get to Menlyn.")
    for x in range(1, 3):
        print("Keep heading North on N1")
        time.sleep(1)
        if x == 2:
            print("Slower, i think i see the offramp.")

    offramps = ["Garsfontein", "Lynwood", "Atterbury"]

    print()
    print("Yup, looks like the Atterbury offramp, lets take it.")
    print()
    if offramps.index("Atterbury"):
        print("Took the Atterbury offramp, good stuff")
    else:
        print("Keep heading North.")

    print()
    print("Keep following the road until you get too the T-junction.")
    for x in range(1, 4):
        if x == 1:
            print()
            print("Slower, almost at the T-junction...")
            time.sleep(1)
        elif x == 2:
            print()
            time.sleep(1.5)
            print("Slower still, you are almost at the T-junction...")
        elif x == 3:
            print()
            time.sleep(2)
            print("Stop!! Turn right here...")

    print()
    print("Drive forward, Menlyn will be on your right hand side.")
    for x in range(1, 4):
        time.sleep(1)
        print("You will see Menlyn in", 4-x)
        if x == 3:
            print("You have arrived, Menlyn is on your right hand side.")


def mainfunction():
    option = None
    option = str(input("Do you want to go to Menlyn or to Home?  Type only m or h please.  "))

    while option != 'm' or 'h':
        print("Incorrect option.")
        option = str(input("Do you want to go to Menlyn or to Home?  Type only m or h please.  "))

        if option == 'm' or 'h':
            if option == 'm':
                menlyn()
                break
            if option == 'h':
                home()
                break


mainfunction()

