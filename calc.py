import locale

# Устанавливаем локаль на русский язык
locale.setlocale(locale.LC_NUMERIC, 'ru_RU.UTF-8')

number = [2.001, 20.2]
formatted_number = locale.format_string('%.3f', number)
print(formatted_number)