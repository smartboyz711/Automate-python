import sys

while True:
    print('Type exit to exit program')
    response = input()
    if response == 'exit' :
        sys.exit()
    print('You Typed "'+ response +'" Please Type again.' )