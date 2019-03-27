# madlib.py
# Eckard

adjective = input("Enter an adjective: ")
noun = input("Enter a noun: ")
verb = input("Enter a verb: ")
adverb = input("Enter an adverb: ")
print("A", adjective, noun, "should never", verb, adverb)

# What follows are examples using the 'sep' parameter and changing it
print("A", adjective, noun, "should never", verb, adverb, sep='$$')
print("A", adjective, noun, "should never", verb, adverb, sep='-')

# What follows is an example of using the 'end' parameter
name = input("Please enter your name: ")
surname = input("Please enter your surname: ")
print("Your email address is: ")
print(name, surname, sep='.', end='@unisa.co.za')



