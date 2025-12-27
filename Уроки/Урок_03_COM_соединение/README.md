# Урок 3: COM-соединение

## Цель урока
Понять принципы работы с COM (Component Object Model) и научиться подключаться к внешним программам, таким как RasterWin.

## Важно!

**COM работает только на Windows!** На Mac можно изучать логику и писать код, но для тестирования нужен компьютер с Windows и установленным RasterWin.

## Содержание урока

### 1. Основы COM (Component Object Model)

COM - это технология Microsoft для взаимодействия между компонентами программного обеспечения.

**Основные понятия:**
- **COM-объект** - объект, который предоставляет свои методы и свойства через COM
- **Интерфейс** - набор методов и свойств, которые доступны у объекта
- **ProgID** - текстовый идентификатор объекта (например, "Astra.Rastr")
- **GUID** - глобальный уникальный идентификатор (например, "{EFC5E4AD-A3DD-11D3-B73F-00500454CF3F}")

### 2. Библиотека pywin32

Для работы с COM в Python используется библиотека `pywin32` (модуль `win32com.client`).

**Установка (только на Windows):**
```bash
pip install pywin32
```

### 3. Подключение к COM-объекту

#### Способ 1: Использование ProgID (рекомендуется)

```python
import win32com.client

# Подключение к RasterWin
rastr = win32com.client.Dispatch("Astra.Rastr")
```

#### Способ 2: Использование GUID (fallback)

```python
import win32com.client

# Если ProgID не работает, используем GUID
guid = "{EFC5E4AD-A3DD-11D3-B73F-00500454CF3F}"
rastr = win32com.client.Dispatch(guid)
```

#### Обработка ошибок подключения

```python
import win32com.client
import sys

try:
    rastr = win32com.client.Dispatch("Astra.Rastr")
    print("✅ Подключение успешно")
except Exception as e:
    print(f"❌ Ошибка подключения: {e}")
    sys.exit(1)
```

### 4. Работа с COM-объектами

#### Вызов методов

```python
# Вызов метода без параметров
rastr.rgm()  # Расчет установившегося режима

# Вызов метода с параметрами
rastr.rgm("p")  # Расчет с плоского старта

# Вызов метода с сохранением результата
результат = rastr.rgm()
```

#### Получение свойств

```python
# Получение свойства
таблицы = rastr.Tables

# Получение вложенных объектов
таблица = rastr.Tables.Item("Generator")
```

#### Установка свойств

```python
# Через методы SetZ
таблица = rastr.Tables.Item("Generator")
колонка = таблица.Cols.Item("P")
колонка.SetZ(0, 100.0)  # Установить значение 100.0 в строку 0
```

### 5. Работа с таблицами RasterWin

#### Получение таблицы

```python
table = rastr.Tables.Item("Generator")  # Таблица генераторов
table = rastr.Tables.Item("node")       # Таблица узлов
table = rastr.Tables.Item("vetv")       # Таблица ветвей
```

#### Получение колонки

```python
table = rastr.Tables.Item("Generator")
col_p = table.Cols.Item("P")    # Колонка мощности
col_name = table.Cols.Item("Name")  # Колонка названия
```

#### Чтение значений

```python
# Получение значения из таблицы
table = rastr.Tables.Item("Generator")
col_p = table.Cols.Item("P")

# Способ 1: Z(index)
значение = col_p.Z(0)  # Значение в строке 0

# Способ 2: GetZ(index) - если Z() не работает
значение = col_p.GetZ(0)
```

#### Запись значений

```python
# Установка значения
table = rastr.Tables.Item("Generator")
col_p = table.Cols.Item("P")

# Способ 1: SetZ(index, value)
col_p.SetZ(0, 100.0)  # Установить 100.0 в строку 0

# Способ 2: set_Z() - если SetZ() не работает
col_p.set_Z(0, 100.0)
```

#### Работа с выборками

```python
table = rastr.Tables.Item("Generator")

# Установка выборки (условие)
table.SetSel("sta=1")  # Выбрать все строки, где sta=1

# Поиск следующей выбранной строки
idx = table.FindNextSel(-1)  # Начать с начала (-1)
while idx != -1:
    print(f"Найден индекс: {idx}")
    idx = table.FindNextSel(idx)  # Найти следующий
```

### 6. Проверка платформы

Так как COM работает только на Windows, важно проверять платформу:

```python
import sys

if sys.platform == 'win32':
    try:
        import win32com.client
        COM_AVAILABLE = True
    except ImportError:
        print("pywin32 не установлен")
        COM_AVAILABLE = False
else:
    print("COM доступен только на Windows")
    COM_AVAILABLE = False
```

## Примеры кода

Смотрите файлы:
- `пример_01_подключение.py` - подключение к COM-объекту
- `пример_02_работа_с_таблицами.py` - работа с таблицами
- `пример_03_обработка_ошибок.py` - обработка ошибок

## Задания

- `Задание_3.1.py` - Подключение к COM-объекту
- `Задание_3.2.py` - Работа с таблицами
- `Задание_3.3.py` - Извлечение данных из таблиц

## Практические советы

1. **Всегда обрабатывайте ошибки** - COM может не подключиться
2. **Используйте try-except** для проверки методов
3. **Проверяйте наличие методов** через `hasattr()` перед вызовом
4. **Изучайте документацию RasterWin API** для списка доступных методов
5. **Тестируйте на Windows** с установленным RasterWin

