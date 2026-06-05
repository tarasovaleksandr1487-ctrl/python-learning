def add(left, right):
    return left + right


def subtract(left, right):
    return left - right


def multiply(left, right):
    return left * right


def divide(left, right):
    if right == 0:
        raise ZeroDivisionError("На ноль делить нельзя.")
    return left / right


OPERATIONS = {
    "+": add,
    "-": subtract,
    "*": multiply,
    "/": divide,
}


def read_number(prompt):
    while True:
        value = input(prompt).replace(",", ".")

        try:
            return float(value)
        except ValueError:
            print("Введите число, например 12 или 3.5.")


def read_operation():
    available = ", ".join(OPERATIONS)

    while True:
        operation = input(f"Выберите операцию ({available}): ").strip()

        if operation in OPERATIONS:
            return operation

        print("Такой операции нет. Используйте +, -, * или /.")


def run_calculator():
    print("Калькулятор")
    print("Введите два числа и выберите действие.")

    while True:
        left = read_number("Первое число: ")
        operation = read_operation()
        right = read_number("Второе число: ")

        try:
            result = OPERATIONS[operation](left, right)
        except ZeroDivisionError as error:
            print(error)
        else:
            print(f"Результат: {result:g}")

        again = input("Посчитать ещё? (да/нет): ").strip().lower()
        if again not in {"да", "д", "yes", "y"}:
            print("Готово.")
            break


if __name__ == "__main__":
    run_calculator()
