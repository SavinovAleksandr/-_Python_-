"""
Пример 3.1: Подключение к COM-объекту

Демонстрирует различные способы подключения к RasterWin через COM
ВАЖНО: Этот код работает только на Windows!
"""

import sys

# Проверка платформы
if sys.platform == 'win32':
    try:
        import win32com.client
        COM_AVAILABLE = True
    except ImportError:
        print("⚠️ pywin32 не установлен. Установите: pip install pywin32")
        COM_AVAILABLE = False
else:
    print("⚠️ COM работает только на Windows")
    COM_AVAILABLE = False


def подключиться_к_rastrwin():
    """
    Подключается к RasterWin через COM API
    
    Returns:
        Объект RasterWin или None если не удалось подключиться
    """
    if not COM_AVAILABLE:
        print("❌ COM недоступен на этой платформе")
        return None
    
    print("=" * 60)
    print("ПОДКЛЮЧЕНИЕ К RASTERWIN")
    print("=" * 60)
    
    # Способ 1: Использование ProgID (рекомендуется)
    print("\nСпособ 1: Подключение через ProgID ('Astra.Rastr')...")
    try:
        rastr = win32com.client.Dispatch("Astra.Rastr")
        print("✅ Успешное подключение через ProgID")
        return rastr
    except Exception as e:
        print(f"⚠️ Не удалось подключиться через ProgID: {e}")
    
    # Способ 2: Использование GUID (fallback)
    print("\nСпособ 2: Подключение через GUID...")
    try:
        guid = "{EFC5E4AD-A3DD-11D3-B73F-00500454CF3F}"
        rastr = win32com.client.Dispatch(guid)
        print("✅ Успешное подключение через GUID")
        return rastr
    except Exception as e:
        print(f"⚠️ Не удалось подключиться через GUID: {e}")
    
    # Способ 3: EnsureDispatch (с генерацией типов)
    print("\nСпособ 3: Подключение через EnsureDispatch...")
    try:
        guid = "{EFC5E4AD-A3DD-11D3-B73F-00500454CF3F}"
        rastr = win32com.client.gencache.EnsureDispatch(guid)
        print("✅ Успешное подключение через EnsureDispatch")
        return rastr
    except Exception as e:
        print(f"⚠️ Не удалось подключиться через EnsureDispatch: {e}")
    
    print("\n❌ Не удалось подключиться к RasterWin")
    print("   Проверьте:")
    print("   1. Установлен ли RasterWin на этом компьютере")
    print("   2. Правильно ли указан ProgID или GUID")
    print("   3. Запущены ли необходимые службы")
    
    return None


def получить_информацию_о_объекте(rastr):
    """
    Получает базовую информацию о подключенном объекте
    
    Args:
        rastr: COM-объект RasterWin
    """
    if rastr is None:
        print("❌ Объект не подключен")
        return
    
    print("\n" + "=" * 60)
    print("ИНФОРМАЦИЯ О RASTERWIN")
    print("=" * 60)
    
    try:
        # Попытка получить список доступных таблиц
        print("\n✅ Подключение активно")
        
        # Проверяем наличие основных методов
        методы = ["rgm", "Load", "Save", "Tables"]
        print("\nДоступные методы:")
        for метод in методы:
            if hasattr(rastr, метод):
                print(f"  ✓ {метод}")
            else:
                print(f"  ✗ {метод} (не найден)")
        
        # Проверяем наличие свойства Tables
        if hasattr(rastr, "Tables"):
            print("\n✓ Свойство 'Tables' доступно")
            try:
                таблицы = rastr.Tables
                print(f"  Количество доступных методов таблиц: {len(dir(таблицы))}")
            except Exception as e:
                print(f"  ⚠️ Ошибка при доступе к Tables: {e}")
        
    except Exception as e:
        print(f"⚠️ Ошибка при получении информации: {e}")


def безопасное_подключение():
    """
    Безопасное подключение с обработкой всех ошибок
    """
    print("=" * 60)
    print("БЕЗОПАСНОЕ ПОДКЛЮЧЕНИЕ")
    print("=" * 60)
    
    if not COM_AVAILABLE:
        print("\n⚠️ COM недоступен на этой платформе")
        print("   На Mac этот код не выполнится, но логику можно изучить")
        return None
    
    try:
        # Пытаемся подключиться
        rastr = win32com.client.Dispatch("Astra.Rastr")
        print("\n✅ Подключение успешно!")
        
        # Проверяем, что объект действительно подключен
        if rastr is None:
            print("❌ Объект не был создан")
            return None
        
        return rastr
        
    except AttributeError as e:
        print(f"\n❌ Ошибка атрибута: {e}")
        print("   Возможно, RasterWin не установлен или ProgID неверен")
        return None
        
    except Exception as e:
        print(f"\n❌ Неожиданная ошибка: {e}")
        print(f"   Тип ошибки: {type(e).__name__}")
        return None


# Пример использования
if __name__ == "__main__":
    print("=" * 60)
    print("ПРИМЕР: Подключение к RasterWin через COM")
    print("=" * 60)
    
    # Подключение
    rastr = подключиться_к_rastrwin()
    
    if rastr:
        # Получение информации
        получить_информацию_о_объекте(rastr)
        
        print("\n" + "=" * 60)
        print("✅ Пример завершен успешно!")
        print("=" * 60)
    else:
        print("\n⚠️ Пример не выполнен")
        print("   На Mac это нормально - COM работает только на Windows")
        print("   Изучите код, чтобы понять логику подключения")

