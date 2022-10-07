import random
import sys
a = 1
b = 10
t = 1
winner = False

secretNumber = random.randint(a,b)
print('I am thinking of number between '+ str(a) +' and '+ str(b))

for i in range(1,t+1):
    print('Take a guess. if you guess wrong number more than '+str(t)+' times you lost.('+str(i)+' times)')
    guessNumber = int(input('Input Number : '))
    if guessNumber < secretNumber :
        print('you guess is too low.')
    elif guessNumber > secretNumber :
        print('you guess ts too high.')
    else :
        print('Good Job! my number is '+str(secretNumber)+' you guessed in '+str(i)+' guess!')
        winner = True
        break
if(winner != True):
    print('you guess Wrong more than '+str(t)+' times you loose! HAHA')
input('press any key to Exit program.')