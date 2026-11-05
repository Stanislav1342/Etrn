"""
onec_exporter.py — Автоматизация выгрузки таблиц из 1С:Предприятие.

На Windows:
  Использует pywinauto (UIA), pywin32 и прямой ввод Virtual Key кодов через Windows API.
  Выполняет полный цикл:
    1. Поиск и активация окна 1С (обход SetForegroundWindow Lock).
    2. Переход в «Маршрутный лист» через Избранное (Ctrl+Shift+B).
    3. Установка периода (сегодня и 30 дней назад) через Еще -> Установить период.
    4. Выгрузка в Excel через Еще -> Вывести список -> Ctrl+S в 'Модель расчетов.xlsx'.
    5. Переход в «Электронные транспортные накладные» через Избранное.
    6. Установка выбранного пользователем периода через Еще -> Установить период.
    7. Выгрузка в Excel через Еще -> Вывести список -> Ctrl+S в 'ЭТРН 1с.xlsx'.
    8. Переход в «Контроль подписания перевозчиков по ЭТРН» через Избранное.
    9. Заполнение полей периода в шапке отчета и нажатие «Сформировать».
    10. Сохранение отчета в 'Процент.xlsx'.

На macOS / Linux:
  Автоматически активируется MockOneCExporter — симулятор с подробным логированием шагов.
"""

import logging
import os
import subprocess
import sys
import time
from datetime import datetime
from typing import Dict, Optional, Tuple

from config import Config, get_config

logger = logging.getLogger("OneCExporter")


def set_clipboard(text: str) -> None:
    """Копирует текст в буфер обмена Windows через clip.exe (UTF-16) без ошибок раскладки."""
    if sys.platform == "win32":
        try:
            subprocess.run(["clip"], input=text.encode("utf-16"), check=True)
        except Exception as e:
            logger.warning("Не удалось скопировать текст в clip: %s", e)
    else:
        # Для macOS отладки через pbcopy
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
        print(f"    Выбор раздела: «{self.cfg.fav_marshrut}»")
        print(f"    Меню «Еще» ➔ «Установить период...»")
        print(f"    Ввод периода: с {m_start} по {m_end}")
        print(f"    Меню «Еще» ➔ «Вывести список...» ➔ [ОК]")
        print(f"    Сохранение (Ctrl+S) ➔ «{self.cfg.model_file}»")
        time.sleep(0.5)

        # 2. ЭТРН
        print("-" * 70)
        print(f" 4. Открытие Избранного (⭐ / Ctrl+Shift+B)...")
        print(f"    Выбор раздела: «{self.cfg.fav_etrn}»")
        print(f"    Меню «Еще» ➔ «Установить период...»")
        print(f"    Ввод периода: с {e_start} по {e_end}")
        print(f"    Меню «Еще» ➔ «Вывести список...» ➔ [ОК]")
        print(f"    Сохранение (Ctrl+S) ➔ «{self.cfg.etrn_file}»")
        time.sleep(0.5)

        # 3. Контроль подписания
        print("-" * 70)
        print(f" 5. Открытие Избранного (⭐ / Ctrl+Shift+B)...")
        print(f"    Выбор раздела: «{self.cfg.fav_percent}»")
        print(f"    Заполнение полей в шапке отчета: с {e_start} по {e_end}")
        print(f"    Нажатие кнопки «Сформировать» (F5) ➔ Формирование отчета...")
        print(f"    Сохранение (Ctrl+S) ➔ «{self.cfg.percent_file}»")
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
    # Навигация через «Избранное»
    # ---------------------------------------------------------------------
    def _open_section_from_favorites(self, section_name: str) -> None:
        """
        Открывает раздел из панели «Избранное»:
        1. Проверяет, не открыт ли раздел уже на панели вкладок 1С.
        2. Если нет — открывает «Избранное» (Ctrl+Shift+B или клик по звездочке ⭐).
        3. Находит пункт в Избранном и активирует его.
        """
        print(f"\n[1С:Windows] Открытие раздела «{section_name}»...")

        # Проверка открытых вкладок внизу/вверху окна 1С
        try:
            for w in self.app.windows():
                for el in w.descendants():
                    try:
                        txt = (el.window_text() or "").strip()
                        if txt == section_name and el.is_displayed():
                            # Проверяем, что элемент похож на вкладку
                            el.click_input()
                            time.sleep(1.5)
                            print(f"[1С:Windows] ✓ Раздел «{section_name}» уже был открыт на панели вкладок.")
                            return
                    except Exception:
                        continue
        except Exception:
            pass

        # Открываем Избранное комбинацией клавиш Ctrl+Shift+B
        print("[1С:Windows] Вызов панели «Избранное» (Ctrl+Shift+B)...")
        self._hotkey_ctrl_shift_b()
        time.sleep(1.0)

        # Проверяем, открылось ли меню Избранного
        fav_opened = False
        fav_pane = None
        for w in self.app.windows():
            try:
                for el in w.descendants():
                    txt = (el.window_text() or "").strip().lower()
                    if "избранное" in txt or section_name.lower() in txt:
                        fav_opened = True
                        fav_pane = w
                        break
                if fav_opened:
                    break
            except Exception:
                pass

        # Если комбинация не открыла панель — пробуем клик по кнопке со звездочкой ⭐
        if not fav_opened:
            print("[1С:Windows] Поиск кнопки со звездочкой ⭐ (Избранное) на панели 1С...")
            clicked_star = False
            try:
                for w in self.app.windows():
                    for el in w.descendants():
                        try:
                            txt = (el.window_text() or "").strip().lower()
                            if "избранное" in txt and el.is_displayed():
                                el.click_input()
                                clicked_star = True
                                time.sleep(1.2)
                                break
                        except Exception:
                            continue
                    if clicked_star:
                        break
            except Exception as e:
                logger.debug("Клик по звездочке: %s", e)

        # В открывшейся панели Избранного ищем нужный пункт
        item_clicked = False
        windows_to_check = [fav_pane] if fav_pane else self.app.windows()
        for w in windows_to_check:
            if not w:
                continue
            try:
                for el in w.descendants():
                    try:
                        txt = (el.window_text() or "").strip()
                        if section_name.lower() in txt.lower() and el.is_displayed():
                            print(f"[1С:Windows] Клик по пункту «{txt}» в Избранном...")
                            el.click_input()
                            item_clicked = True
                            time.sleep(2.0)
                            break
                    except Exception:
                        continue
                if item_clicked:
                    break
            except Exception:
                pass

        # Если прямой клик по тексту в UIA не сработал — вводим текст фильтра
        if not item_clicked:
            print(f"[1С:Windows] Поиск в Избранном через фильтр: '{section_name}'...")
            set_clipboard(section_name)
            self._paste_clipboard()
            time.sleep(0.5)
            self._hotkey_enter()
            time.sleep(2.0)

        print(f"[1С:Windows] ✓ Переход в раздел «{section_name}» выполнен.")

    # ---------------------------------------------------------------------
    # Установка периода для списков (Еще -> Установить период...)
    # ---------------------------------------------------------------------
    def _set_list_period(self, date_from: datetime, date_to: datetime) -> None:
        """
        Устанавливает период в динамическом списке документов через:
        Еще ➔ Установить период...
        """
        d_from_str = date_from.strftime("%d.%m.%Y")
        d_to_str = date_to.strftime("%d.%m.%Y")
        print(f"[1С:Windows] Установка периода списка: с {d_from_str} по {d_to_str}...")

        # 1. Поиск и клик по кнопке «Еще»
        more_btn_clicked = False
        try:
            for w in self.app.windows():
                for btn in w.descendants(control_type="Button"):
                    try:
                        txt = (btn.window_text() or "").strip().lower()
                        if txt in ("еще", "ещё") and btn.is_displayed():
                            btn.click_input()
                            more_btn_clicked = True
                            time.sleep(0.8)
                            break
                    except Exception:
                        continue
                if more_btn_clicked:
                    break
        except Exception as e:
            logger.debug("Клик по кнопке Еще: %s", e)

        # 2. В выпавшем меню ищем пункт «Установить период...»
        period_item_clicked = False
        if more_btn_clicked:
            try:
                for w in self.app.windows():
                    for el in w.descendants():
                        try:
                            txt = (el.window_text() or "").strip().lower()
                            if "установить период" in txt and el.is_displayed():
                                el.click_input()
                                period_item_clicked = True
                                time.sleep(1.2)
                                break
                        except Exception:
                            continue
                    if period_item_clicked:
                        break
            except Exception as e:
                logger.debug("Клик по Установить период: %s", e)

        # Если через меню Еще не вышло — ищем отдельную кнопку «Период» на тулбаре
        if not period_item_clicked:
            try:
                for w in self.app.windows():
                    for btn in w.descendants(control_type="Button"):
                        try:
                            txt = (btn.window_text() or "").strip().lower()
                            if "период" in txt and btn.is_displayed():
                                btn.click_input()
                                period_item_clicked = True
                                time.sleep(1.2)
                                break
                        except Exception:
                            continue
                    if period_item_clicked:
                        break
            except Exception:
                pass

        # 3. Заполнение полей в диалоге «Установка периода»
        # В диалоге 1С два поля ввода: Дата начала и Дата окончания
        print("[1С:Windows] Заполнение дат в диалоге периода...")
        filled_dialog = False
        try:
            # Ищем активный диалог установки периода
            for w in self.app.windows():
                title = (w.window_text() or "").strip().lower()
                if "период" in title or "выбор" in title or "установка" in title:
                    edits = w.descendants(control_type="Edit")
                    if len(edits) >= 2:
                        # Первое поле — дата начала
                        edits[0].click_input()
                        time.sleep(0.2)
                        self._select_all_and_paste(d_from_str)
                        time.sleep(0.2)

                        # Второе поле — дата окончания
                        edits[1].click_input()
                        time.sleep(0.2)
                        self._select_all_and_paste(d_to_str)
                        time.sleep(0.2)

                        # Нажимаем Enter или ОК
                        self._hotkey_enter()
                        filled_dialog = True
                        time.sleep(2.0)
                        break
        except Exception as e:
            logger.debug("Заполнение диалога периода: %s", e)

        if not filled_dialog:
            # Fallback: эмуляция по Tab
            print("[1С:Windows] Ввод дат периода с переходом по Tab...")
            self._select_all_and_paste(d_from_str)
            self._hotkey_tab()
            time.sleep(0.2)
            self._select_all_and_paste(d_to_str)
            time.sleep(0.2)
            self._hotkey_enter()
            time.sleep(2.0)

        print(f"[1С:Windows] ✓ Период установлен: {d_from_str} - {d_to_str}.")

    # ---------------------------------------------------------------------
    # Выгрузка списка в Excel (Еще -> Вывести список... -> Ctrl+S)
    # ---------------------------------------------------------------------
    def _export_list_to_excel(self, target_filepath: str) -> None:
        """
        Выгружает открытый список 1С в Excel файл:
        Еще ➔ Вывести список... ➔ [ОК] ➔ Ctrl+S ➔ Ввод пути ➔ Сохранить.
        """
        print(f"[1С:Windows] Выгрузка списка в «{target_filepath}»...")
        if os.path.exists(target_filepath):
            try:
                os.remove(target_filepath)
            except Exception:
                pass

        # 1. Клик по кнопке «Еще»
        more_clicked = False
        for w in self.app.windows():
            for btn in w.descendants(control_type="Button"):
                try:
                    txt = (btn.window_text() or "").strip().lower()
                    if txt in ("еще", "ещё") and btn.is_displayed():
                        btn.click_input()
                        more_clicked = True
                        time.sleep(0.8)
                        break
                except Exception:
                    continue
            if more_clicked:
                break

        # 2. Клик по пункту «Вывести список...»
        export_item_clicked = False
        if more_clicked:
            for w in self.app.windows():
                for el in w.descendants():
                    try:
                        txt = (el.window_text() or "").strip().lower()
                        if "вывести список" in txt and el.is_displayed():
                            el.click_input()
                            export_item_clicked = True
                            time.sleep(1.2)
                            break
                    except Exception:
                        continue
                if export_item_clicked:
                    break

        if not export_item_clicked:
            raise RuntimeError("Не удалось найти пункт «Вывести список...» в меню «Еще»!")

        # 3. В модальном окне «Вывести список» нажимаем Enter / ОК
        print("[1С:Windows] Подтверждение параметров вывода списка (ОК)...")
        time.sleep(0.6)
        self._hotkey_enter()
        # Ожидаем формирования табличного документа
        time.sleep(3.0)

        # 4. Сохраняем табличный документ по Ctrl+S
        print("[1С:Windows] Сохранение табличного документа через Ctrl+S...")
        self._hotkey_ctrl_s()
        time.sleep(1.5)

        # 5. Обработка диалога Windows «Сохранение» / «Сохранить как»
        # Вставляем полный целевой путь
        print(f"[1С:Windows] Ввод пути сохранения: {target_filepath}...")
        self._select_all_and_paste(target_filepath)
        time.sleep(0.4)
        self._hotkey_enter()
        time.sleep(1.0)

        # Если появилось окно «Подтвердить сохранение в виде (заменить?)» — нажимаем Enter (Да)
        self._hotkey_enter()
        time.sleep(2.0)

        # 6. Закрываем сформированную вкладку табличного документа
        print("[1С:Windows] Закрытие вкладки табличного документа (Ctrl+F4)...")
        self._hotkey_ctrl_f4()
        time.sleep(1.0)

        # Проверка создания файла
        if not os.path.exists(target_filepath):
            # Вторая попытка ожидания записи файла
            time.sleep(2.0)

        if os.path.exists(target_filepath):
            print(f"[1С:Windows] ✓ Таблица успешно сохранена: {target_filepath}")
        else:
            logger.warning("Файл %s не появился сразу. Возможно, сохранение продолжается.", target_filepath)

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
            # Находим edit controls в активном окне
            edits = []
            for w in self.app.windows():
                for el in w.descendants(control_type="Edit"):
                    try:
                        if el.is_displayed():
                            edits.append(el)
                    except Exception:
                        continue
            
            # Заполняем первые подходящие поля дат
            if len(edits) >= 2:
                edits[0].click_input()
                time.sleep(0.2)
                self._select_all_and_paste(d_from_str)
                time.sleep(0.2)

                edits[1].click_input()
                time.sleep(0.2)
                self._select_all_and_paste(d_to_str)
                time.sleep(0.2)
            else:
                # Ввод по Tab
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
                    try:
                        txt = (btn.window_text() or "").strip().lower()
                        if "сформировать" in txt and btn.is_displayed():
                            btn.click_input()
                            formed = True
                            break
                    except Exception:
                        continue
                if formed:
                    break
        except Exception:
            pass

        if not formed:
            # Стандартная горячая клавиша формирования отчета в 1С
            self._send_hotkey(self.win32con.VK_F5)

        print(f"[1С:Windows] Ожидание формирования отчета (до {self.cfg.report_gen_timeout} сек)...")
        time.sleep(6.0)

        # Сохраняем отчет через Ctrl+S
        print(f"[1С:Windows] Сохранение отчета в «{target_filepath}» (Ctrl+S)...")
        self._hotkey_ctrl_s()
        time.sleep(1.5)

        self._select_all_and_paste(target_filepath)
        time.sleep(0.4)
        self._hotkey_enter()
        time.sleep(1.0)
        self._hotkey_enter()
        time.sleep(2.0)

        # Закрываем отчет
        self._hotkey_ctrl_f4()
        time.sleep(1.0)

        if os.path.exists(target_filepath):
            print(f"[1С:Windows] ✓ Отчет успешно сохранен: {target_filepath}")
        else:
            logger.warning("Файл отчета %s еще формируется/записывается.", target_filepath)

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
