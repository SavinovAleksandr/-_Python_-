# Урок 2: Работа с Excel

## Цель урока
Научиться читать и записывать данные в Excel файлы, что необходимо для работы с исходными данными и результатами расчетов динамической устойчивости.

## Содержание урока

### 1. Библиотека openpyxl

Библиотека `openpyxl` используется для работы с файлами Excel (.xlsx).

#### Основные операции:

```python
from openpyxl import Workbook, load_workbook

# Создание нового файла
workbook = Workbook()
worksheet = workbook.active
worksheet.title = "Результаты"

# Открытие существующего файла
workbook = load_workbook("файл.xlsx")
worksheet = workbook["Лист1"]  # или workbook.active
```

#### Работа с ячейками:

```python
# Запись в ячейку
worksheet["A1"] = "Заголовок"
worksheet.cell(row=1, column=1, value="Значение")

# Чтение из ячейки
значение = worksheet["A1"].value
значение = worksheet.cell(row=1, column=1).value
```

### 2. Чтение данных из Excel

#### Чтение конкретных ячеек:

```python
from openpyxl import load_workbook

workbook = load_workbook("данные.xlsx", data_only=True)
worksheet = workbook["Параметры"]

# Чтение значения
напряжение = worksheet["B2"].value
мощность = worksheet.cell(row=3, column=2).value
```

#### Чтение строк и столбцов:

```python
# Чтение всей строки
строка = [cell.value for cell in worksheet[1]]

# Чтение всего столбца
столбец = [cell.value for cell in worksheet["A"]]

# Чтение диапазона
диапазон = []
for row in worksheet["A1:C10"]:
    для_строки = [cell.value for cell in row]
    диапазон.append(для_строки)
```

### 3. Запись данных в Excel

#### Запись одиночных значений:

```python
worksheet["A1"] = "Название параметра"
worksheet["B1"] = 220.5
worksheet.cell(row=2, column=1, value="Напряжение")
```

#### Запись списка данных:

```python
данные = [100, 150, 200, 250, 300]

for индекс, значение in enumerate(данные, start=1):
    worksheet.cell(row=индекс, column=1, value=значение)

# Или проще:
for row_idx, значение in enumerate(данные, start=1):
    worksheet[f"A{row_idx}"] = значение
```

### 4. Форматирование Excel

```python
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Шрифт
cell = worksheet["A1"]
cell.font = Font(bold=True, size=12, name="Times New Roman")

# Выравнивание
cell.alignment = Alignment(horizontal='center', vertical='center')

# Заливка цветом
cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")

# Границы
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
cell.border = thin_border
```

### 5. Работа с несколькими листами

```python
# Создание нового листа
новый_лист = workbook.create_sheet("Результаты")

# Переключение между листами
worksheet = workbook["Параметры"]
worksheet2 = workbook["Результаты"]

# Список всех листов
print(workbook.sheetnames)
```

### 6. Библиотека pandas (альтернативный способ)

Для более сложной работы с данными удобнее использовать pandas:

```python
import pandas as pd

# Чтение Excel
df = pd.read_excel("данные.xlsx", sheet_name="Параметры")

# Чтение конкретного диапазона
df = pd.read_excel("данные.xlsx", sheet_name="Параметры", 
                   usecols="A:C", nrows=10)

# Запись в Excel
df.to_excel("результаты.xlsx", sheet_name="Результаты", index=False)

# Фильтрация данных
отфильтрованные = df[df["Напряжение"] > 200]

# Группировка
сгруппированные = df.groupby("Режим").mean()
```

## Примеры кода

Смотрите файлы в папке:
- `пример_01_чтение_Excel.py` - базовое чтение
- `пример_02_запись_Excel.py` - базовая запись
- `пример_03_форматирование.py` - форматирование
- `пример_04_сложные_данные.py` - работа со сложными структурами

## Задания

Переходите к файлам заданий:
- `Задание_2.1.py` - Чтение параметров из Excel
- `Задание_2.2.py` - Запись результатов в Excel
- `Задание_2.3.py` - Форматирование результатов
- `Задание_2.4.py` - Комплексная задача: чтение → обработка → запись

## Практические советы

1. **Используйте `data_only=True`** при чтении, чтобы получать вычисленные значения формул
2. **Сохраняйте файлы** после изменений: `workbook.save("файл.xlsx")`
3. **Обрабатывайте ошибки**: файл может не существовать или быть открыт в Excel
4. **Используйте pandas** для сложных операций с данными

