"""
ЗАДАНИЕ 1.1: Работа с переменными и типами данных

ЗАДАЧА:
Создайте переменные для хранения параметров расчета динамической устойчивости:
1. Название проекта (строка)
2. Начальное напряжение (число с плавающей точкой)
3. Количество итераций (целое число)
4. Расчет завершен (булево значение)
5. Допустимое отклонение напряжения (число)

ЗАДАНИЯ:
1. Создайте все переменные с начальными значениями
2. Выведите тип каждой переменной
3. Измените значение напряжения на 10% больше начального
4. Создайте строку с полной информацией о проекте, используя f-строки
5. Проверьте, находится ли новое напряжение в диапазоне: 
   начальное ± допустимое отклонение

ПРИМЕР ВЫВОДА:
Параметр: название_проекта = "Проект_2024" (тип: str)
Параметр: начальное_напряжение = 220.0 В (тип: float)
...

"""

# ===== ВАШ КОД НИЖЕ =====

# 1. Создайте переменные
project_name = "StabLimit"
invalible_voltage = 220.0
iterations = 10
calculation_completed = False
allowed_voltage_deviation = 1.0

# 2. Выведите типы переменных
print("project_name: ", type(project_name))
print("invalible_voltage: ", type(invalible_voltage))
print("iterations: ", type(iterations))
print("calculation_completed: ", type(calculation_completed))
print("allowed_voltage_deviation: ", type(allowed_voltage_deviation))
# 3. Измените напряжение (увеличьте на 10%)
new_voltage = invalible_voltage * 1.1
# 4. Создайте строку с информацией
full_information = f"Project: {project_name}, Initial voltage: {invalible_voltage} V, Iterations: {iterations}, Calculation completed: {calculation_completed}, Allowed voltage deviation: {allowed_voltage_deviation} V"

# 5. Проверьте диапазон напряжения
in_range = new_voltage >= invalible_voltage - allowed_voltage_deviation and new_voltage <= invalible_voltage + allowed_voltage_deviation
print("in_range: ", in_range)

print("full_information: ", full_information)

# ===== ПРОВЕРКА =====
# Запустите этот скрипт и проверьте вывод
# Убедитесь, что все работает корректно

# ЗАДАНИЕ ВЫПОЛНЕНО

"""
задание выполнено
"""

# вдлаваываыв

