import random


secret_number = random.randint(1, 100)
attempts = 0

print("Я загадал число от 1 до 100.")
print("Попробуй угадать его!")

while True:
    user_input = input("Введи число: ")
    attempts += 1

    guess = int(user_input)

    if guess < secret_number:
        print("Моё число больше.")
    elif guess > secret_number:
        print("Моё число меньше.")
    else:
        print(f"Верно! Ты угадал число за {attempts} попыток.")
        break
