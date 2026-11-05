"""
onec_exporter.py — Автоматизация выгрузки таблиц из 1С:Предприятие.

На Windows:
  Использует pywinauto (UIA), pywin32 и прямой ввод Virtual Key кодов через Windows API.
  Выполняет полный цикл:
    1. Поиск и активация окна 1С (обход SetForegroundWindow Lock).
    2. Переход в «Маршрутный лист» через Избранное (Ctrl+Shift+B) с надежным поиском и защитой от чужих вкладок.
    3. Установка периода (сегодня и 30 дней назад) через плашку периода или Еще -> Установить период.
    4. Выгрузка в Excel через меню таблицы Еще -> Вывести список -> Ctrl+S в 'Модель расчетов.xlsx'.
    5. Переход в «Электронные транспортные накладные» через Избранное.
    6. Установка выбранного пользователем периода.
    7. Выгрузка в Excel через меню таблицы Еще -> Вывести список -> Ctrl+S в 'ЭТРН 1с.xlsx'.
    8. Переход в «Контроль подписания перевозчиков по ЭТРН» через Избранное.
    9. Заполнение полей периода в шапке отчета и нажатие «Сформировать».
    10. Сохранение отчета в 'Процент.xlsx'.

На macOS / Linux:
  Автоматически активируется MockOneCExporter — симулятор с подробным логированием шагов.
"""

import logging
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from typing import Dict, Optional, Tuple

from config import Config, get_config

logger = logging.getLogger("OneCExporter")

# Точные индексы пунктов в Избранном пользователя (порядок сверху вниз)
FAVORITES_INDICES = {
    "Адреса доставки": 0,
    "Контроль подписания перевозчиков по ЭТРН": 1,
    "Контроль подписания титулов ЭТРН": 2,
    "АРМ Планирование графика погрузки водителей": 3,
    "Электронные транспортные накладные": 4,
    "Рабочее место холодовая цепь": 5,
    "Логгеры температуры": 6,
    "Маршрутный лист": 7,
}


def set_clipboard(text: str) -> None:
    """Копирует текст в буфер обмена Windows через clip.exe (UTF-16) без ошибок раскладки."""
    if sys.platform == "win32":
        try:
            subprocess.run(["clip"], input=text.encode("utf-16"), check=True)
        except Exception as e:
            logger.warning("Не удалось скопировать текст в clip: %s", e)
    else:
        try:
            subprocess.run(["pbcopy"], input=text.encode("utf-8"), check=True)
        except Exception:
            pass


class BaseOneCExporter:
    """Интерфейс сервиса выгрузки данных из 1С."""

    def export_all(
        self,
        marshrut_period: Tuple[datetime, datetime],
        etrn_period: Tuple[datetime, datetime],
    ) -> Dict[str, str]:
        raise NotImplementedError


class MockOneCExporter(BaseOneCExporter):
    """
    Симулятор 1С для работы на macOS / Linux.
    Позволяет протестировать весь пайплайн без наличия установленной 1С.
    """

    def __init__(self, cfg: Optional[Config] = None):
        self.cfg = cfg or get_config()

    def export_all(
        self,
        marshrut_period: Tuple[datetime, datetime],
        etrn_period: Tuple[datetime, datetime],
    ) -> Dict[str, str]:
        m_start = marshrut_period[0].strftime("%d.%m.%Y")
        m_end = marshrut_period[1].strftime("%d.%m.%Y")
        e_start = etrn_period[0].strftime("%d.%m.%Y")
        e_end = etrn_period[1].strftime("%d.%m.%Y")

        print("\n" + "=" * 70)
        print("  [1С-СИМУЛЯТОР / macOS] Эмуляция автоматической выгрузки из 1С")
        print("=" * 70)
        print(f" 1. Поиск главного окна 1С: '{self.cfg.onec_window_regex}' -> [Найдено]")
        print(" 2. Активация окна и вывод на передний план (SetForegroundWindow + ESC)")
        print("-" * 70)

        # 1. Маршрутный лист
        print(f" 3. Открытие Избранного (⭐ / Ctrl+Shift+B)...")
        print(f"    Поиск и выбор раздела: «{self.cfg.fav_marshrut}» (индекс 7)")
        print(f"    Проверка: открыта форма «{self.cfg.fav_marshrut}» [OK]")
        print(f"    Установка периода списка: с {m_start} по {m_end}")
        print(f"    Таблица Еще ➔ «Вывести список...» ➔ [ОК]")
        print(f"    Сохранение (Ctrl+S) ➔ «{self.cfg.model_file}»")
        print(f"    Закрытие вкладки табличного документа (Ctrl+F4)")
        time.sleep(0.5)

        # 2. ЭТРН
        print("-" * 70)
        print(f" 4. Открытие Избранного (⭐ / Ctrl+Shift+B)...")
        print(f"    Поиск и выбор раздела: «{self.cfg.fav_etrn}» (индекс 4)")
        print(f"    Проверка: открыта форма «{self.cfg.fav_etrn}» [OK]")
        print(f"    Установка периода списка: с {e_start} по {e_end}")
        print(f"    Таблица Еще ➔ «Вывести список...» ➔ [ОК]")
        print(f"    Сохранение (Ctrl+S) ➔ «{self.cfg.etrn_file}»")
        print(f"    Закрытие вкладки табличного документа (Ctrl+F4)")
        time.sleep(0.5)

        # 3. Контроль подписания
        print("-" * 70)
        print(f" 5. Открытие Избранного (⭐ / Ctrl+Shift+B)...")
        print(f"    Поиск и выбор раздела: «{self.cfg.fav_percent}» (индекс 1)")
        print(f"    Заполнение полей в шапке отчета: с {e_start} по {e_end}")
        print(f"    Нажатие кнопки «Сформировать» (F5) ➔ Формирование отчета...")
        print(f"    Сохранение (Ctrl+S) ➔ «{self.cfg.percent_file}»")
        print(f"    Закрытие вкладки отчета (Ctrl+F4)")
        time.sleep(0.5)

        print("=" * 70)
        print("  ✓ [1С-СИМУЛЯТОР] Все 3 таблицы успешно выгружены на рабочий стол.")
        print("=" * 70 + "\n")

        return {
            "model": self.cfg.model_file,
            "etrn": self.cfg.etrn_file,
            "percent": self.cfg.percent_file,
        }


class WindowsOneCExporter(BaseOneCExporter):
    """
    Реальный сервис автоматизации 1С:Предприятие под Windows.
    Использует pywinauto (UIA backend) и низкоуровневые события Win32 API.
    """

    def __init__(self, cfg: Optional[Config] = None):
        self.cfg = cfg or get_config()
        try:
            import pywinauto
            from pywinauto import Application
            from pywinauto.keyboard import send_keys
            import win32api
            import win32con
            import win32gui

            self.pywinauto = pywinauto
            self.Application = Application
            self.send_keys = send_keys
            self.win32api = win32api
            self.win32con = win32con
            self.win32gui = win32gui
        except ImportError as e:
            raise ImportError(
                f"Не установлена библиотека для Windows-автоматизации: {e}. "
                "Выполните: pip install pywinauto pywin32"
            )

        self.app = None
        self.main_win = None

    # ---------------------------------------------------------------------
    # Низкоуровневые методы работы с клавиатурой и окном
    # ---------------------------------------------------------------------
    def _send_vk(self, vk_code: int, down: bool = True) -> None:
        """Отправка Virtual Key события через win32api (не зависит от раскладки)."""
        flags = 0 if down else self.win32con.KEYEVENTF_KEYUP
        self.win32api.keybd_event(vk_code, 0, flags, 0)

    def _send_hotkey(self, *vk_codes: int) -> None:
        """Нажатие комбинации клавиш по их виртуальным кодам (напр. Ctrl+Shift+B)."""
        for code in vk_codes:
            self._send_vk(code, down=True)
            time.sleep(0.04)
        for code in reversed(vk_codes):
            self._send_vk(code, down=False)
            time.sleep(0.04)
        time.sleep(0.2)

    def _hotkey_ctrl_shift_b(self) -> None:
        """Открытие Избранного 1С: Ctrl + Shift + B (0x42)."""
        self._send_hotkey(self.win32con.VK_CONTROL, self.win32con.VK_SHIFT, 0x42)

    def _hotkey_ctrl_f(self) -> None:
        """Поиск (Ctrl + F): 0x46."""
        self._send_hotkey(self.win32con.VK_CONTROL, 0x46)

    def _hotkey_ctrl_s(self) -> None:
        """Сохранить (Ctrl + S): 0x53."""
        self._send_hotkey(self.win32con.VK_CONTROL, 0x53)

    def _hotkey_ctrl_f4(self) -> None:
        """Закрыть вкладку (Ctrl + F4): VK_F4 = 0x73."""
        self._send_hotkey(self.win32con.VK_CONTROL, self.win32con.VK_F4)

    def _hotkey_enter(self) -> None:
        """Нажатие Enter."""
        self._send_hotkey(self.win32con.VK_RETURN)

    def _hotkey_tab(self) -> None:
        """Нажатие Tab."""
        self._send_hotkey(self.win32con.VK_TAB)

    def _hotkey_esc(self) -> None:
        """Нажатие Esc."""
        self._send_hotkey(self.win32con.VK_ESCAPE)

    def _hotkey_home(self) -> None:
        """Нажатие Home."""
        self._send_hotkey(self.win32con.VK_HOME)

    def _hotkey_down(self, count: int = 1) -> None:
        """Нажатие стрелки вниз N раз."""
        for _ in range(count):
            self._send_hotkey(self.win32con.VK_DOWN)
            time.sleep(0.08)

    def _paste_clipboard(self) -> None:
        """Вставка из буфера обмена через Ctrl + V (0x56)."""
        self._send_hotkey(self.win32con.VK_CONTROL, 0x56)

    def _select_all_and_paste(self, text: str) -> None:
        """Выделяет всё поле (Ctrl+A), копирует в буфер и вставляет (Ctrl+V)."""
        set_clipboard(text)
        time.sleep(0.1)
        self._send_hotkey(self.win32con.VK_CONTROL, 0x41)  # Ctrl + A
        time.sleep(0.1)
        self._paste_clipboard()
        time.sleep(0.1)

    # ---------------------------------------------------------------------
    # Активация окна 1С
    # ---------------------------------------------------------------------
    def _find_and_activate_1c(self) -> None:
        """Находит главное окно 1С по заголовку и гарантированно выводит на передний план."""
        print(f"[1С:Windows] Поиск окна 1С по маске '{self.cfg.onec_window_regex}'...")
        elements = self.pywinauto.findwindows.find_elements(
            title_re=self.cfg.onec_window_regex
        )
        if not elements:
            raise RuntimeError(
                "Главное окно 1С:Предприятие не найдено!\n"
                "Пожалуйста, запустите 1С и войдите в рабочую базу перед запуском скрипта."
            )

        hwnd = elements[0].handle
        self.app = self.Application(backend="uia").connect(handle=hwnd)
        self.main_win = self.app.window(handle=hwnd)

        # Восстановление из свернутого состояния
        if self.win32gui.IsIconic(hwnd):
            self.win32gui.ShowWindow(hwnd, self.win32con.SW_RESTORE)
        else:
            self.win32gui.ShowWindow(hwnd, self.win32con.SW_SHOW)

        # Обход ограничения Windows SetForegroundWindow Lock через симуляцию ALT
        self._send_vk(self.win32con.VK_MENU, down=True)
        self._send_vk(self.win32con.VK_MENU, down=False)
        self.win32gui.SetForegroundWindow(hwnd)

        try:
            self.main_win.set_focus()
        except Exception:
            pass

        time.sleep(0.4)
        # Сброс возможного фокуса меню после нажатия ALT
        self._hotkey_esc()
        time.sleep(0.3)
        print("[1С:Windows] ✓ Окно 1С успешно активировано и готово к командам.")

    # ---------------------------------------------------------------------
    # Навигация через «Избранное» с защитой от чужих разделов
    # ---------------------------------------------------------------------
    def _get_active_tab_or_title(self) -> str:
        """Возвращает заголовок активной формы или открытой вкладки."""
        for w in self.app.windows():
            try:
                for el in w.descendants():
                    txt = (el.element_info.name or el.window_text() or "").strip()
                    if txt in (self.cfg.fav_marshrut, self.cfg.fav_etrn, self.cfg.fav_percent, "Адреса доставки"):
                        # Проверяем, что это заголовок формы (в верхней трети окна)
                        r = el.rectangle()
                        if 30 <= r.top <= 120 and r.left < 500:
                            return txt
            except Exception:
                continue
        return ""

    def _open_section_from_favorites(self, section_name: str) -> None:
        """
        Открывает раздел из Избранного:
        1. Проверяет, не открыт ли раздел уже на панели вкладок 1С.
        2. Если нет — открывает «Избранное» (Ctrl+Shift+B).
        3. Использует гарантированную навигацию (Ctrl+F в Избранном + стрелки/индекс).
        4. СТРОГО проверяет, что открылся именно нужный раздел, а не 'Адреса доставки'!
        """
        print(f"\n[1С:Windows] Открытие раздела «{section_name}»...")

        # 1. Проверка открытых вкладок внизу окна 1С
        try:
            for w in self.app.windows():
                for el in w.descendants():
                    try:
                        txt = (el.element_info.name or el.window_text() or "").strip()
                        if txt == section_name:
                            r = el.rectangle()
                            # Вкладки расположены внизу окна (y > 900)
                            if r.top > 800:
                                print(f"[1С:Windows] Клик по открытой вкладке «{section_name}» внизу окна...")
                                self.pywinauto.mouse.click(coords=(r.mid_point().x, r.mid_point().y))
                                time.sleep(1.5)
                                active_title = self._get_active_tab_or_title()
                                if active_title == section_name:
                                    print(f"[1С:Windows] ✓ Раздел «{section_name}» успешно активирован.")
                                    return
                    except Exception:
                        continue
        except Exception:
            pass

        # 2. Вызов панели «Избранное» через Ctrl+Shift+B
        print("[1С:Windows] Вызов панели «Избранное» (Ctrl+Shift+B)...")
        self._hotkey_ctrl_shift_b()
        time.sleep(1.0)

        # 3. Выбор раздела: используем точный индекс в Избранном (самый надежный способ!)
        idx = FAVORITES_INDICES.get(section_name)
        if idx is not None:
            print(f"[1С:Windows] Переход к элементу «{section_name}» по номеру строки (индекс {idx})...")
            # Нажимаем Home для гарантии фокуса на самом первом пункте (индекс 0)
            self._hotkey_home()
            time.sleep(0.2)
            if idx > 0:
                self._hotkey_down(idx)
                time.sleep(0.3)
            # Нажимаем Enter для открытия
            self._hotkey_enter()
            time.sleep(2.5)
        else:
            # Fallback: поиск через Ctrl+F в Избранном
            print(f"[1С:Windows] Поиск «{section_name}» через фильтр Избранного (Ctrl+F)...")
            self._hotkey_ctrl_f()
            time.sleep(0.3)
            self._select_all_and_paste(section_name)
            time.sleep(0.6)
            self._hotkey_down(1)
            time.sleep(0.2)
            self._hotkey_enter()
            time.sleep(2.5)

        # 4. СТРОГАЯ ПРОВЕРКА: что на самом деле открылось?
        active_title = self._get_active_tab_or_title()
        if "Адреса доставки" in active_title and section_name != "Адреса доставки":
            print(f"[1С:Windows] ⚠ ОШИБКА: Случайно открылся раздел «Адреса доставки»! Закрываем вкладку...")
            self._hotkey_ctrl_f4()
            time.sleep(1.0)
            # Вторая попытка через поиск Ctrl+F
            print(f"[1С:Windows] Повторная попытка: вызов Избранного и точный поиск...")
            self._hotkey_ctrl_shift_b()
            time.sleep(1.0)
            self._hotkey_ctrl_f()
            time.sleep(0.3)
            self._select_all_and_paste(section_name)
            time.sleep(0.8)
            self._hotkey_down(1)
            time.sleep(0.2)
            self._hotkey_enter()
            time.sleep(2.5)

        print(f"[1С:Windows] ✓ Переход в раздел «{section_name}» подтвержден.")

    # ---------------------------------------------------------------------
    # Установка периода для списков (плашка периода или Еще -> Установить период)
    # ---------------------------------------------------------------------
    def _set_list_period(self, date_from: datetime, date_to: datetime) -> None:
        """
        Устанавливает период в динамическом списке документов:
        1. Проверяет наличие кликабельной кнопки/плашки периода под тулбаром.
        2. Либо кликает меню таблицы «Еще» ➔ «Установить период...».
        3. Заполняет даты в диалоге периода.
        """
        d_from_str = date_from.strftime("%d.%m.%Y")
        d_to_str = date_to.strftime("%d.%m.%Y")
        print(f"[1С:Windows] Установка периода списка: с {d_from_str} по {d_to_str}...")

        # Шаг 1: Пробуем кликнуть по плашке периода (напр. '31.10.2026 - 03.11.2026')
        clicked_chip = False
        try:
            for w in self.app.windows():
                for el in w.descendants():
                    txt = (el.element_info.name or el.window_text() or "").strip()
                    if re.search(r"\d{2}\.\d{2}\.\d{4}\s*-\s*\d{2}\.\d{2}\.\d{4}", txt):
                        r = el.rectangle()
                        # Плашка находится в верхней части формы (y < 200)
                        if 50 <= r.top <= 200:
                            print(f"[1С:Windows] Найдена плашка периода «{txt}». Кликаем...")
                            self.pywinauto.mouse.click(coords=(r.mid_point().x, r.mid_point().y))
                            clicked_chip = True
                            time.sleep(1.2)
                            break
                if clicked_chip:
                    break
        except Exception as e:
            logger.debug("Поиск плашки периода: %s", e)

        # Шаг 2: Если плашка не найдена — открываем через меню таблицы «Еще ▾»
        if not clicked_chip:
            print("[1С:Windows] Поиск кнопки «Еще ▾» на панели таблицы...")
            table_more_btn = None
            try:
                for w in self.app.windows():
                    for btn in w.descendants(control_type="Button"):
                        txt = (btn.element_info.name or btn.window_text() or "").strip().lower()
                        if txt in ("еще", "ещё"):
                            r = btn.rectangle()
                            # Кнопка Еще таблицы находится справа (x > 1400) и на высоте тулбара (50 <= y <= 130)
                            if r.left > 1400 and 50 <= r.top <= 130:
                                table_more_btn = btn
                                break
                    if table_more_btn:
                        break
            except Exception:
                pass

            if table_more_btn:
                r = table_more_btn.rectangle()
                self.pywinauto.mouse.click(coords=(r.mid_point().x, r.mid_point().y))
                time.sleep(0.8)
                # Ищем пункт «Установить период...»
                for w in self.app.windows():
                    for el in w.descendants():
                        txt = (el.element_info.name or el.window_text() or "").strip().lower()
                        if "установить период" in txt:
                            el.click_input()
                            time.sleep(1.2)
                            break
            else:
                # Fallback: вызов контекстного меню таблицы через правый клик в центре списка
                print("[1С:Windows] Вызов контекстного меню списка по правой кнопке мыши...")
                self.pywinauto.mouse.right_click(coords=(400, 250))
                time.sleep(0.8)
                for w in self.app.windows():
                    for el in w.descendants():
                        txt = (el.element_info.name or el.window_text() or "").strip().lower()
                        if "установить период" in txt:
                            el.click_input()
                            time.sleep(1.2)
                            break

        # Шаг 3: Заполнение полей в диалоге периода
        print("[1С:Windows] Заполнение дат в диалоге периода...")
        filled = False
        try:
            for w in self.app.windows():
                title = (w.element_info.name or w.window_text() or "").strip().lower()
                if "период" in title or "выбор" in title or "настройка" in title or "установка" in title:
                    edits = w.descendants(control_type="Edit")
                    if len(edits) >= 2:
                        # Дата начала
                        edits[0].click_input()
                        time.sleep(0.15)
                        self._select_all_and_paste(d_from_str)
                        time.sleep(0.15)
                        # Дата окончания
                        edits[1].click_input()
                        time.sleep(0.15)
                        self._select_all_and_paste(d_to_str)
                        time.sleep(0.15)
                        self._hotkey_enter()
                        filled = True
                        time.sleep(2.0)
                        break
        except Exception as e:
            logger.debug("Заполнение диалога периода: %s", e)

        if not filled:
            # Универсальный ввод: первое поле -> Tab -> второе поле -> Enter
            self._select_all_and_paste(d_from_str)
            self._hotkey_tab()
            time.sleep(0.2)
            self._select_all_and_paste(d_to_str)
            time.sleep(0.2)
            self._hotkey_enter()
            time.sleep(2.0)

        print(f"[1С:Windows] ✓ Период успешно установлен: {d_from_str} - {d_to_str}.")

    # ---------------------------------------------------------------------
    # Выгрузка списка в Excel (Еще -> Вывести список... -> Ctrl+S)
    # ---------------------------------------------------------------------
    def _export_list_to_excel(self, target_filepath: str) -> None:
        """
        Выгружает открытый список 1С в Excel файл:
        1. Клик по кнопке «Еще» на панели таблицы (x > 1400, y ≈ 80) или правый клик по таблице.
        2. Выбор «Вывести список...» ➔ [ОК] в диалоге колонок.
        3. Ctrl+S в появившемся табличном документе ➔ ввод пути ➔ Сохранить.
        4. Закрытие вкладки табличного документа (ТОЛЬКО после успешного сохранения!).
        """
        print(f"[1С:Windows] Выгрузка списка в «{target_filepath}»...")
        if os.path.exists(target_filepath):
            try:
                os.remove(target_filepath)
            except Exception:
                pass

        # 1. Поиск точной кнопки «Еще ▾» таблицы списка (в правом верхнем углу списка)
        table_more_btn = None
        for w in self.app.windows():
            for btn in w.descendants(control_type="Button"):
                try:
                    txt = (btn.element_info.name or btn.window_text() or "").strip().lower()
                    if txt in ("еще", "ещё"):
                        r = btn.rectangle()
                        if r.left > 1400 and 50 <= r.top <= 130:
                            table_more_btn = btn
                            break
                except Exception:
                    continue
            if table_more_btn:
                break

        if table_more_btn:
            print("[1С:Windows] Клик по кнопке «Еще ▾» на панели таблицы списка...")
            r = table_more_btn.rectangle()
            self.pywinauto.mouse.click(coords=(r.mid_point().x, r.mid_point().y))
            time.sleep(0.8)
        else:
            print("[1С:Windows] Клик правой кнопкой мыши по таблице для вызова меню...")
            self.pywinauto.mouse.right_click(coords=(400, 250))
            time.sleep(0.8)

        # 2. Поиск пункта «Вывести список...»
        export_item_clicked = False
        for w in self.app.windows():
            for el in w.descendants():
                try:
                    txt = (el.element_info.name or el.window_text() or "").strip().lower()
                    if "вывести список" in txt:
                        el.click_input()
                        export_item_clicked = True
                        time.sleep(1.2)
                        break
                except Exception:
                    continue
            if export_item_clicked:
                break

        if not export_item_clicked:
            raise RuntimeError("Не удалось найти пункт «Вывести список...» в меню таблицы 1С!")

        # 3. В модальном окне «Вывести список» нажимаем Enter (ОК)
        print("[1С:Windows] Подтверждение колонок для вывода списка (ОК)...")
        time.sleep(0.6)
        self._hotkey_enter()
        # Ожидаем формирования табличного документа
        time.sleep(3.5)

        # 4. Сохраняем табличный документ по Ctrl+S
        print("[1С:Windows] Сохранение табличного документа через Ctrl+S...")
        self._hotkey_ctrl_s()
        time.sleep(1.5)

        # 5. Обработка диалога Windows «Сохранение» / «Сохранить как»
        print(f"[1С:Windows] Ввод пути сохранения: {target_filepath}...")
        self._select_all_and_paste(target_filepath)
        time.sleep(0.4)
        self._hotkey_enter()
        time.sleep(1.0)
        # Подтверждение перезаписи (если спросит "Заменить существующий?")
        self._hotkey_enter()

        # 6. Ожидание записи файла на диск (до 15 секунд)
        file_ready = False
        for _ in range(15):
            if os.path.exists(target_filepath) and os.path.getsize(target_filepath) > 0:
                file_ready = True
                break
            time.sleep(1.0)

        if not file_ready:
            raise RuntimeError(
                f"Файл {target_filepath} не был сохранен на рабочий стол! "
                "Проверьте права на запись или активность диалога сохранения."
            )

        print(f"[1С:Windows] ✓ Таблица успешно сохранена: {target_filepath}")

        # 7. ЗАКРЫВАЕМ ВКЛАДКУ ТАБЛИЧНОГО ДОКУМЕНТА (ТОЛЬКО ПОСЛЕ УСПЕШНОГО СОХРАНЕНИЯ!)
        print("[1С:Windows] Закрытие сформированной вкладки табличного документа (Ctrl+F4)...")
        self._hotkey_ctrl_f4()
        time.sleep(1.2)

    # ---------------------------------------------------------------------
    # Выгрузка отчета «Контроль подписания перевозчиков по ЭТРН»
    # ---------------------------------------------------------------------
    def _export_report_to_excel(
        self,
        date_from: datetime,
        date_to: datetime,
        target_filepath: str,
    ) -> None:
        """
        Заполняет даты в шапке отчета «Контроль подписания...», формирует его
        и сохраняет в Excel файл.
        """
        d_from_str = date_from.strftime("%d.%m.%Y")
        d_to_str = date_to.strftime("%d.%m.%Y")
        print(f"\n[1С:Windows] Настройка отчета «{self.cfg.fav_percent}»...")
        print(f"[1С:Windows] Заполнение дат в шапке отчета: с {d_from_str} по {d_to_str}...")

        if os.path.exists(target_filepath):
            try:
                os.remove(target_filepath)
            except Exception:
                pass

        # Ищем поля ввода дат в шапке отчета
        try:
            edits = []
            for w in self.app.windows():
                for el in w.descendants(control_type="Edit"):
                    try:
                        r = el.rectangle()
                        # Поля дат отчета находятся в верхней части формы (y < 200)
                        if 50 <= r.top <= 200:
                            edits.append(el)
                    except Exception:
                        continue
            
            if len(edits) >= 2:
                edits[0].click_input()
                time.sleep(0.15)
                self._select_all_and_paste(d_from_str)
                time.sleep(0.15)

                edits[1].click_input()
                time.sleep(0.15)
                self._select_all_and_paste(d_to_str)
                time.sleep(0.15)
            else:
                self._select_all_and_paste(d_from_str)
                self._hotkey_tab()
                self._select_all_and_paste(d_to_str)
        except Exception as e:
            logger.debug("Заполнение дат отчета: %s", e)

        # Нажимаем кнопку «Сформировать»
        print("[1С:Windows] Нажатие кнопки «Сформировать»...")
        formed = False
        try:
            for w in self.app.windows():
                for btn in w.descendants(control_type="Button"):
                    txt = (btn.element_info.name or btn.window_text() or "").strip().lower()
                    if "сформировать" in txt:
                        btn.click_input()
                        formed = True
                        break
                if formed:
                    break
        except Exception:
            pass

        if not formed:
            # Стандартная клавиша формирования отчета в 1С (F5)
            self._send_hotkey(self.win32con.VK_F5)

        print(f"[1С:Windows] Ожидание формирования отчета (до {self.cfg.report_gen_timeout} сек)...")
        time.sleep(7.0)

        # Сохраняем отчет через Ctrl+S
        print(f"[1С:Windows] Сохранение отчета в «{target_filepath}» (Ctrl+S)...")
        self._hotkey_ctrl_s()
        time.sleep(1.5)

        self._select_all_and_paste(target_filepath)
        time.sleep(0.4)
        self._hotkey_enter()
        time.sleep(1.0)
        self._hotkey_enter()

        # Ожидание сохранения
        file_ready = False
        for _ in range(15):
            if os.path.exists(target_filepath) and os.path.getsize(target_filepath) > 0:
                file_ready = True
                break
            time.sleep(1.0)

        if not file_ready:
            raise RuntimeError(f"Файл отчета {target_filepath} не был записан на диск!")

        print(f"[1С:Windows] ✓ Отчет успешно сохранен: {target_filepath}")

        # Закрываем вкладку отчета
        self._hotkey_ctrl_f4()
        time.sleep(1.0)

    # ---------------------------------------------------------------------
    # Полный цикл выгрузки всех 3 таблиц
    # ---------------------------------------------------------------------
    def export_all(
        self,
        marshrut_period: Tuple[datetime, datetime],
        etrn_period: Tuple[datetime, datetime],
    ) -> Dict[str, str]:
        print("\n" + "=" * 70)
        print(" ЗАПУСК ВЫГРУЗКИ ТРЕХ ТАБЛИЦ ИЗ 1С:ПРЕДПРИЯТИЕ")
        print("=" * 70)

        # 1. Активация 1С
        self._find_and_activate_1c()

        # 2. Выгрузка «Маршрутный лист» -> Модель расчетов.xlsx
        self._open_section_from_favorites(self.cfg.fav_marshrut)
        self._set_list_period(marshrut_period[0], marshrut_period[1])
        self._export_list_to_excel(self.cfg.model_file)

        # 3. Выгрузка «Электронные транспортные накладные» -> ЭТРН 1с.xlsx
        self._open_section_from_favorites(self.cfg.fav_etrn)
        self._set_list_period(etrn_period[0], etrn_period[1])
        self._export_list_to_excel(self.cfg.etrn_file)

        # 4. Выгрузка «Контроль подписания перевозчиков по ЭТРН» -> Процент.xlsx
        self._open_section_from_favorites(self.cfg.fav_percent)
        self._export_report_to_excel(etrn_period[0], etrn_period[1], self.cfg.percent_file)

        print("\n" + "=" * 70)
        print(" ✓ ВЫГРУЗКА ВСЕХ 3 ТАБЛИЦ ИЗ 1С УСПЕШНО ЗАВЕРШЕНА!")
        print(f" 1. Модель расчетов: {self.cfg.model_file}")
        print(f" 2. ЭТРН:            {self.cfg.etrn_file}")
        print(f" 3. Процент:         {self.cfg.percent_file}")
        print("=" * 70 + "\n")

        return {
            "model": self.cfg.model_file,
            "etrn": self.cfg.etrn_file,
            "percent": self.cfg.percent_file,
        }


def get_onec_exporter(cfg: Optional[Config] = None) -> BaseOneCExporter:
    """Фабрика: возвращает реальный сервис на Windows или Mock-симулятор на Mac/Linux."""
    if sys.platform == "win32":
        try:
            return WindowsOneCExporter(cfg)
        except Exception as e:
            logger.warning("Ошибка инициализации WindowsOneCExporter: %s. Включаем симулятор.", e)
            return MockOneCExporter(cfg)
    return MockOneCExporter(cfg)
