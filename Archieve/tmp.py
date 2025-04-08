from num2words import num2words

def convert_to_words(amount):
    try:
        # Проверяем, что число не отрицательное
        if amount < 0:
            raise ValueError("Введите положительное число.")

        # Разделяем рубли и копейки
        rubles = int(amount)
        kopecks = int(round((amount - rubles) * 100))

        # Определяем правильные формы для рублей и копеек
        rubles_words = f"{num2words(rubles, lang='ru')} {get_currency_form(rubles, ['рубль', 'рубля', 'рублей'])}"
        kopecks_words = f"{kopecks:02d} {get_currency_form(kopecks, ['копейка', 'копейки', 'копеек'])}"

        return f"{rubles} ({rubles_words}), {kopecks_words}"
    except ValueError as e:
        return f"Ошибка: {e}"

def get_currency_form(number, forms):
    """
    Определяет правильную форму слова на основании числа.
    :param number: Число, для которого нужно выбрать форму
    :param forms: Список форм: ['единственное', 'множественное 1-4', 'множественное 5-9, 0']
    :return: Правильная форма слова
    """
    if 11 <= number % 100 <= 19:
        return forms[2]
    elif number % 10 == 1:
        return forms[0]
    elif 2 <= number % 10 <= 4:
        return forms[1]
    else:
        return forms[2]

# Ввод пользователя
try:
    user_input = input("Введите сумму в рублях (например, 109.25): ")
    amount = float(user_input)
    result = convert_to_words(amount)
    print(result)
except ValueError:
    print("Ошибка: введено некорректное число.")
