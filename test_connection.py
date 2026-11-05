r"""
test_connection.py — Быстрая диагностика возможности прямого COM-подключения к 1С.

Скрипт проверяет:
1. Зарегистрирован ли в Windows компонент V83.COMConnector.
2. Считывает параметры базы из сетевого файла \\puls.local\1c\1CBases\uuv.v8i
   (или локального ibases.v8i).
3. Пробует установить прямое фоновое соединение с базой без открытия графического окна 1С.
"""

import os
import re
import sys

V8I_NETWORK_PATH = r"\\puls.local\1c\1CBases\uuv.v8i"
V8I_LOCAL_PATH = os.path.expandvars(r"%APPDATA%\1C\1CEStart\ibases.v8i")


def check_com_available():
    print("=" * 70)
    print("  🔍 ТЕСТ ПРЯМОГО COM-ПОДКЛЮЧЕНИЯ К 1С:ПРЕДПРИЯТИЕ (БЕЗ GUI)")
    print("=" * 70)

    # 1. Проверка библиотеки pywin32
    try:
        import win32com.client
    except ImportError:
        print("[ОШИБКА] Не установлена библиотека pywin32.")
        print("Выполните: pip install pywin32")
        return

    # 2. Проверка регистрации V83.COMConnector
    print("1. Проверка регистрации COM-компонента V83.COMConnector в Windows...")
    try:
        connector = win32com.client.Dispatch("V83.COMConnector")
        print("   ✓ V83.COMConnector успешно инициализирован в системе!")
    except Exception as e:
        print(f"   ✗ Не удалось создать V83.COMConnector: {e}")
        print("   (Возможно, на ПК установлен только Тонкий клиент без компоненты COM-соединения)")
        return

    # 3. Поиск файла со списком баз (uuv.v8i)
    v8i_file = None
    if os.path.exists(V8I_NETWORK_PATH):
        v8i_file = V8I_NETWORK_PATH
        print(f"2. Найден сетевой файл списка баз: {v8i_file}")
    elif os.path.exists(V8I_LOCAL_PATH):
        v8i_file = V8I_LOCAL_PATH
        print(f"2. Найден локальный файл списка баз: {v8i_file}")
    else:
        print(f"2. ✗ Файлы баз не найдены по путям:")
        print(f"   • {V8I_NETWORK_PATH}")
        print(f"   • {V8I_LOCAL_PATH}")

    # Чтение строк подключения из v8i
    bases = []
    if v8i_file:
        try:
            with open(v8i_file, "r", encoding="utf-8-sig", errors="ignore") as f:
                content = f.read()
            # Ищем секции [ИмяБазы] и Connect=...
            matches = re.findall(r"\[(.*?)\][\s\S]*?Connect=(.*?)\n", content)
            for name, conn in matches:
                name = name.strip()
                conn = conn.strip()
                bases.append((name, conn))
                print(f"   • База: «{name}» ➔ Connect={conn}")
        except Exception as e:
            print(f"   Ошибка чтения {v8i_file}: {e}")

    if not bases:
        print("\n[ИНФО] Список баз пуст. Пробуем получить активное окно 1С...")
        return

    # 4. Проба подключения к первой подходящей базе (ЕИС / МСК / uuv)
    target_base = None
    for name, conn in bases:
        if any(w in name.lower() for w in ["еис", "мск", "корп", "уув", "uuv"]):
            target_base = (name, conn)
            break

    if not target_base and bases:
        target_base = bases[0]

    name, conn = target_base
    print(f"\n3. Тестовая попытка прямого фонового подключения к базе «{name}»...")
    print(f"   Строка подключения: {conn}")

    try:
        connection = connector.Connect(conn)
        print("\n" + "=" * 70)
        print("  🎉 УСПЕХ! ПРЯМОЕ COM-СОЕДИНЕНИЕ С 1С РАБОТАЕТ!")
        print("=" * 70)
        print("Это означает, что мы можем получать любые данные из 1С НАПРЯМУЮ,")
        print("без мыши, без кликов, без открывания окон — за 1-2 секунды!")
        print("=" * 70 + "\n")
    except Exception as e:
        print("\n" + "=" * 70)
        print(f"  ✗ Ошибка прямого подключения к 1С: {e}")
        print("=" * 70)
        print("Причина: в корпоративной базе для COM может требоваться логин/пароль")
        print("или право 'Внешнее соединение' заблокировано политикой безопасности.")
        print("В этом случае надежным решением является UIA-автоматизация на системных командах.")
        print("=" * 70 + "\n")


if __name__ == "__main__":
    check_com_available()
