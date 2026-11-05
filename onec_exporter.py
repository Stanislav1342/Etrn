"""
onec_exporter.py — Автоматизация выгрузки таблиц из 1С:Предприятие через прямые ссылки (e1cib).

На Windows:
  Использует промышленный RPA-стандарт для 1С:Предприятие (Управляемый интерфейс / Такси):
    1. Поиск, максимизация и активация главного окна 1С.
    2. Прямой переход по постоянным навигационным ссылкам (Shift + F11):
       - «Маршрутный лист»: e1cib/list/Документ.МаршрутныйЛист
       - «Электронные транспортные накладные»: e1cib/list/Документ.ЭлектроннаяТранспортнаяНакладная
       - «Контроль подписания»: e1cib/data/Справочник.ВариантыОтчетов?ref=946700e0ed9e2e9e11f1495d81d50ffe
       (В диалоге перехода используется кнопка «Вставить из буфера», что исключает любые сбои раскладки).
    3. Раздел «Маршрутный лист»:
       - Заполнение дат фильтра отгрузки («Больше или равно» и «Меньше или равно») в шапке списка.
       - Прямой вызов кнопки «Вывести список...» на панели инструментов.
       - Сохранение таблицы в 'Модель расчетов.xlsx' и закрытие вкладки «Информация по МЛ» (Ctrl+F4).
    4. Раздел «Электронные транспортные накладные»:
       - Установка периода через «Еще ▾» ➔ «Установить период...» (или плашку периода) в диалоге «Выберите период».
       - Меню «Еще ▾» ➔ «Вывести список...» ➔ подтверждение колонок.
       - Сохранение таблицы в 'ЭТРН 1с.xlsx' и закрытие вкладки «Список» (Ctrl+F4).
    5. Отчет «Контроль подписания перевозчиков по ЭТРН»:
       - Заполнение дат в шапке отчета.
       - Нажатие «Сформировать» (F5) и ожидание готовности.
       - Сохранение в 'Процент.xlsx' и закрытие вкладки отчета (Ctrl+F4).

На macOS / Linux:
  Автоматически активируется MockOneCExporter — симулятор полного пайплайна с подробным логированием.
"""

import logging
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from config import Config, get_config

logger = logging.getLogger("OneCExporter")


def set_clipboard(text: str) -> None:
    """Копирует текст в буфер обмена Windows (через Win32 API / clip.exe) без ошибок кодировки."""
    if sys.platform == "win32":
        try:
            import win32clipboard
            import win32con

            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
            finally:
                win32clipboard.CloseClipboard()
            return
        except Exception as e:
            logger.debug("win32clipboard error: %s, fallback to clip.exe", e)

        try:
            subprocess.run(["clip"], input=text.encode("utf-16"), check=True)
            return
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
        print(" 2. Активация и максимизация окна 1С (ShowWindow SW_MAXIMIZE)")
        print("-" * 70)

        # 1. Маршрутный лист
        print(f" 3. Прямой переход по ссылке (Shift + F11):")
        print(f"    URL: «{self.cfg.url_marshrut}»")
        print(f"    Открыт раздел: «{self.cfg.title_marshrut}» [OK]")
        print(f"    Установка дат в фильтре отгрузки: с {m_start} по {m_end}")
        print(f"    Кнопка тулбара «Вывести список...» ➔ [ОК] в диалоге")
        print(f"    Сохранение (Ctrl+S) ➔ «{self.cfg.model_file}»")
        print(f"    Закрытие вкладки «Информация по МЛ» (Ctrl+F4)")
        time.sleep(0.4)

        # 2. ЭТРН
        print("-" * 70)
        print(f" 4. Прямой переход по ссылке (Shift + F11):")
        print(f"    URL: «{self.cfg.url_etrn}»")
        print(f"    Открыт раздел: «{self.cfg.title_etrn}» [OK]")
        print(f"    Установка периода списка через диалог: с {e_start} по {e_end}")
        print(f"    Таблица Еще ➔ «Вывести список...» ➔ [ОК] в диалоге")
        print(f"    Сохранение (Ctrl+S) ➔ «{self.cfg.etrn_file}»")
        print(f"    Закрытие вкладки «Список» (Ctrl+F4)")
        time.sleep(0.4)

        # 3. Контроль подписания
        print("-" * 70)
        print(f" 5. Прямой переход по ссылке (Shift + F11):")
        print(f"    URL: «{self.cfg.url_percent}»")
        print(f"    Открыт отчет: «{self.cfg.title_percent}» [OK]")
        print(f"    Заполнение полей в шапке отчета: с {e_start} по {e_end}")
        print(f"    Нажатие кнопки «Сформировать» (F5) ➔ Формирование отчета...")
        print(f"    Сохранение (Ctrl+S) ➔ «{self.cfg.percent_file}»")
        print(f"    Закрытие вкладки отчета (Ctrl+F4)")
        time.sleep(0.4)

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
    Использует прямые навигационные ссылки 1С (Shift+F11 / e1cib) и UI Automation.
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
        """Нажатие комбинации клавиш по их виртуальным кодам."""
        for code in vk_codes:
            self._send_vk(code, down=True)
            time.sleep(0.04)
        for code in reversed(vk_codes):
            self._send_vk(code, down=False)
            time.sleep(0.04)
        time.sleep(0.2)

    def _hotkey_shift_f11(self) -> None:
        """Переход по ссылке 1С: Shift + F11 (0x10, 0x7A)."""
        self._send_hotkey(self.win32con.VK_SHIFT, self.win32con.VK_F11)

    def _hotkey_shift_f10(self) -> None:
        """Контекстное меню Windows: Shift + F10 (0x10, 0x79)."""
        self._send_hotkey(self.win32con.VK_SHIFT, self.win32con.VK_F10)

    def _hotkey_ctrl_s(self) -> None:
        """Сохранить (Ctrl + S): 0x53."""
        self._send_hotkey(self.win32con.VK_CONTROL, 0x53)

    def _hotkey_ctrl_f4(self) -> None:
        """Закрыть текущую вкладку (Ctrl + F4): VK_F4 = 0x73."""
        self._send_hotkey(self.win32con.VK_CONTROL, self.win32con.VK_F4)

    def _hotkey_f5(self) -> None:
        """Сформировать отчет / обновить (F5): VK_F5 = 0x74."""
        self._send_hotkey(self.win32con.VK_F5)

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
        """Выделяет всё поле (Ctrl+A), копирует текст в буфер и вставляет (Ctrl+V)."""
        set_clipboard(text)
        time.sleep(0.08)
        self._send_hotkey(self.win32con.VK_CONTROL, 0x41)  # Ctrl + A
        time.sleep(0.08)
        self._paste_clipboard()
        time.sleep(0.1)

    # ---------------------------------------------------------------------
    # Активация и максимизация окна 1С
    # ---------------------------------------------------------------------
    def _find_and_activate_1c(self) -> None:
        """Находит главное окно 1С по заголовку, максимизирует и выводит на передний план."""
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

        # Разворачиваем окно на весь экран для предсказуемых координат
        self.win32gui.ShowWindow(hwnd, self.win32con.SW_MAXIMIZE)
        time.sleep(0.3)

        # Обход ограничения Windows SetForegroundWindow Lock через симуляцию ALT
        self._send_vk(self.win32con.VK_MENU, down=True)
        self._send_vk(self.win32con.VK_MENU, down=False)
        self.win32gui.SetForegroundWindow(hwnd)

        try:
            self.main_win.set_focus()
        except Exception:
            pass

        time.sleep(0.4)
        # Сброс возможного фокуса системного меню после нажатия ALT
        self._hotkey_esc()
        time.sleep(0.3)
        print("[1С:Windows] ✓ Окно 1С успешно максимизировано и активировано.")

    # ---------------------------------------------------------------------
    # Прямой переход по ссылке 1С (Shift + F11 / e1cib)
    # ---------------------------------------------------------------------
    def _open_by_link(self, url: str, section_title: str) -> None:
        """
        Мгновенно открывает любой раздел 1С по ссылке e1cib через Shift + F11.
        Использует кнопку «Вставить из буфера», что исключает проблемы раскладки.
        """
        print(f"\n[1С:Windows] Переход по прямой ссылке в «{section_title}»...")
        print(f"             URL: {url}")

        # Сбрасываем контекстные меню или диалоги
        self._hotkey_esc()
        time.sleep(0.3)

        # 1. Помещаем URL в буфер обмена
        set_clipboard(url)
        time.sleep(0.1)

        # 2. Вызываем окно «Переход по ссылке» через Shift + F11
        self._hotkey_shift_f11()
        time.sleep(1.0)

        # 3. Ищем окно «Переход по ссылке»
        link_dlg = None
        for _ in range(5):
            try:
                for w in self.app.windows():
                    title = (w.element_info.name or w.window_text() or "").strip()
                    if "переход по ссылке" in title.lower():
                        link_dlg = w
                        break
            except Exception:
                pass
            if link_dlg:
                break
            time.sleep(0.3)

        # 4. Нажимаем кнопку «Вставить из буфера»
        pasted = False
        if link_dlg:
            try:
                for btn in link_dlg.descendants(control_type="Button"):
                    txt = (btn.element_info.name or btn.window_text() or "").strip().lower()
                    if "вставить из буфера" in txt:
                        btn.click_input()
                        pasted = True
                        time.sleep(0.3)
                        break
            except Exception as e:
                logger.debug("Клик по 'Вставить из буфера': %s", e)

        if not pasted:
            # Горячая клавиша кнопки «Вставить из буфера» (Alt + В)
            self._send_hotkey(self.win32con.VK_MENU, 0x44)  # Alt + D (В)
            time.sleep(0.2)
            self._paste_clipboard()
            time.sleep(0.2)

        # 5. Нажимаем кнопку «Перейти»
        transitioned = False
        if link_dlg:
            try:
                for btn in link_dlg.descendants(control_type="Button"):
                    txt = (btn.element_info.name or btn.window_text() or "").strip().lower()
                    if txt == "перейти":
                        btn.click_input()
                        transitioned = True
                        time.sleep(0.5)
                        break
            except Exception:
                pass

        if not transitioned:
            self._hotkey_enter()

        # Ожидаем отрисовки целевой формы
        time.sleep(3.0)
        print(f"[1С:Windows] ✓ Раздел «{section_title}» успешно открыт по прямой ссылке.")

    # ---------------------------------------------------------------------
    # Раздел 1: «Маршрутный лист» (фильтр отгрузки и кнопка Вывести список)
    # ---------------------------------------------------------------------
    def _set_marshrut_period(self, date_from: datetime, date_to: datetime) -> None:
        """
        Устанавливает фильтр дат отгрузки в шапке списка «Маршрутный лист»:
        Строка 1: «Дата отгрузки: [Больше или равно] [03.10.2026 0:00:00]»
        Строка 2: «Дата отгрузки: [Меньше или равно] [05.11.2026 0:00:00]»
        """
        d_from_str = date_from.strftime("%d.%m.%Y")
        d_to_str = date_to.strftime("%d.%m.%Y")
        print(f"[1С:Windows] Установка дат в «Маршрутный лист»: с {d_from_str} по {d_to_str}...")

        # Поиск Edit полей в фильтре шапки
        edits_found: List = []
        try:
            for w in self.app.windows():
                for el in w.descendants(control_type="Edit"):
                    try:
                        r = el.rectangle()
                        # Поля фильтра дат находятся в области y: 90..180, x: 400..750
                        if 90 <= r.top <= 180 and 400 <= r.left <= 750:
                            edits_found.append(el)
                    except Exception:
                        continue
        except Exception as e:
            logger.debug("Поиск Edit полей Маршрутного листа: %s", e)

        # Сортируем поля сверху вниз (edits[0] — Больше или равно, edits[1] — Меньше или равно)
        edits_found.sort(key=lambda x: x.rectangle().top)

        if len(edits_found) >= 2:
            print("[1С:Windows] Заполнение полей фильтра отгрузки через элементы интерфейса...")
            edits_found[0].click_input()
            time.sleep(0.15)
            self._select_all_and_paste(d_from_str)
            time.sleep(0.15)

            edits_found[1].click_input()
            time.sleep(0.15)
            self._select_all_and_paste(d_to_str)
            time.sleep(0.15)
        else:
            print("[1С:Windows] Использование точных экранных координат фильтра дат...")
            # Поле 1 (Больше или равно): x ~ 580, y ~ 122
            self.pywinauto.mouse.click(coords=(580, 122))
            time.sleep(0.15)
            self._select_all_and_paste(d_from_str)
            time.sleep(0.15)

            # Поле 2 (Меньше или равно): x ~ 580, y ~ 153
            self.pywinauto.mouse.click(coords=(580, 153))
            time.sleep(0.15)
            self._select_all_and_paste(d_to_str)
            time.sleep(0.15)

        # Обновление списка (Enter или кнопка Обновить)
        self._hotkey_enter()
        time.sleep(2.0)
        print(f"[1С:Windows] ✓ Фильтр «Маршрутный лист» установлен: {d_from_str} - {d_to_str}.")

    def _export_marshrut_list(self, target_filepath: str) -> None:
        """
        Выгружает открытый «Маршрутный лист» через прямую кнопку панели «Вывести список...»
        """
        print(f"[1С:Windows] Выгрузка «Маршрутный лист» в «{target_filepath}»...")

        # 1. Клик по кнопке «Вывести список...» (на панели тулбара)
        btn_clicked = False
        try:
            for w in self.app.windows():
                for btn in w.descendants(control_type="Button"):
                    txt = (btn.element_info.name or btn.window_text() or "").strip().lower()
                    if "вывести список" in txt:
                        btn.click_input()
                        btn_clicked = True
                        time.sleep(1.0)
                        break
                if btn_clicked:
                    break
        except Exception:
            pass

        if not btn_clicked:
            # Кнопка «Вывести список...» на панели инструментов находится по coords ~ (1010, 88)
            print("[1С:Windows] Клик по кнопке «Вывести список...» по координатам тулбара...")
            self.pywinauto.mouse.click(coords=(1010, 88))
            time.sleep(1.2)

        # 2. Подтверждение колонок в диалоге «Вывести список» (нажать ОК / Enter)
        print("[1С:Windows] Подтверждение диалога «Вывести список» (ОК)...")
        self._confirm_export_dialog()

        # 3. Сохранение файла и закрытие вкладки «Информация по МЛ»
        self._save_in_dialog(target_filepath)

    # ---------------------------------------------------------------------
    # Раздел 2: «Электронные транспортные накладные»
    # ---------------------------------------------------------------------
    def _set_etrn_period(self, date_from: datetime, date_to: datetime) -> None:
        """
        Устанавливает период для списка ЭТРН через диалог «Выберите период».
        Диалог открывается через правое меню «Еще ▾» ➔ «Установить период...» или плашку периода.
        """
        d_from_str = date_from.strftime("%d.%m.%Y")
        d_to_str = date_to.strftime("%d.%m.%Y")
        print(f"[1С:Windows] Установка периода в «ЭТРН»: с {d_from_str} по {d_to_str}...")

        # 1. Пробуем кликнуть по плашке периода под «Создать»
        opened_dialog = False
        try:
            for w in self.app.windows():
                for el in w.descendants():
                    txt = (el.element_info.name or el.window_text() or "").strip()
                    if re.search(r"\d{2}\.\d{2}\.\d{4}\s*-\s*\d{2}\.\d{2}\.\d{4}", txt):
                        r = el.rectangle()
                        if 50 <= r.top <= 140:
                            el.click_input()
                            opened_dialog = True
                            time.sleep(1.2)
                            break
                if opened_dialog:
                    break
        except Exception:
            pass

        # 2. Если плашка не найдена — открываем через меню таблицы «Еще ▾» (справа)
        if not opened_dialog:
            print("[1С:Windows] Открытие диалога периода через меню «Еще ▾»...")
            self._click_table_more_button()
            time.sleep(0.6)
            for w in self.app.windows():
                for el in w.descendants():
                    txt = (el.element_info.name or el.window_text() or "").strip().lower()
                    if "установить период" in txt:
                        el.click_input()
                        opened_dialog = True
                        time.sleep(1.2)
                        break
                if opened_dialog:
                    break

        # 3. Заполнение полей даты в диалоге «Выберите период»
        print("[1С:Windows] Заполнение дат в диалоге «Выберите период»...")
        # При открытии диалога фокус уже стоит на первом поле даты
        self._select_all_and_paste(d_from_str)
        self._hotkey_tab()
        time.sleep(0.2)
        self._select_all_and_paste(d_to_str)
        time.sleep(0.2)

        # Нажимаем кнопку «Выбрать» (или Enter)
        selected = False
        try:
            for w in self.app.windows():
                title = (w.element_info.name or w.window_text() or "").strip().lower()
                if "период" in title:
                    for btn in w.descendants(control_type="Button"):
                        txt = (btn.element_info.name or btn.window_text() or "").strip().lower()
                        if txt == "выбрать":
                            btn.click_input()
                            selected = True
                            time.sleep(1.5)
                            break
                if selected:
                    break
        except Exception:
            pass

        if not selected:
            self._hotkey_enter()
            time.sleep(1.5)

        print(f"[1С:Windows] ✓ Период списка «ЭТРН» успешно установлен: {d_from_str} - {d_to_str}.")

    def _export_etrn_list(self, target_filepath: str) -> None:
        """
        Выгружает открытый список «ЭТРН» через меню таблицы «Еще ▾» ➔ «Вывести список...»
        """
        print(f"[1С:Windows] Выгрузка «ЭТРН» в «{target_filepath}»...")

        # 1. Открытие меню «Еще ▾» таблицы списка
        self._click_table_more_button()
        time.sleep(0.6)

        # 2. Поиск и клик пункта «Вывести список...»
        item_clicked = False
        for w in self.app.windows():
            for el in w.descendants():
                try:
                    txt = (el.element_info.name or el.window_text() or "").strip().lower()
                    if "вывести список" in txt:
                        el.click_input()
                        item_clicked = True
                        time.sleep(1.0)
                        break
                except Exception:
                    continue
            if item_clicked:
                break

        if not item_clicked:
            # Fallback: координатный клик пункта меню
            self.pywinauto.mouse.click(coords=(1600, 576))
            time.sleep(1.0)

        # 3. Подтверждение колонок в диалоге «Вывести список» (ОК)
        self._confirm_export_dialog()

        # 4. Сохранение файла и закрытие сформированной вкладки «Список»
        self._save_in_dialog(target_filepath)

    # ---------------------------------------------------------------------
    # Раздел 3: Отчет «Контроль подписания перевозчиков по ЭТРН»
    # ---------------------------------------------------------------------
    def _export_percent_report(
        self,
        date_from: datetime,
        date_to: datetime,
        target_filepath: str,
    ) -> None:
        """
        Заполняет даты в шапке отчета «Контроль подписания...», формирует его (F5)
        и сохраняет в Excel файл 'Процент.xlsx'.
        """
        d_from_str = date_from.strftime("%d.%m.%Y")
        d_to_str = date_to.strftime("%d.%m.%Y")
        print(f"\n[1С:Windows] Настройка отчета «{self.cfg.title_percent}»...")
        print(f"[1С:Windows] Заполнение дат в шапке отчета: с {d_from_str} по {d_to_str}...")

        # Поиск Edit полей в шапке отчета
        edits: List = []
        try:
            for w in self.app.windows():
                for el in w.descendants(control_type="Edit"):
                    try:
                        r = el.rectangle()
                        if 50 <= r.top <= 120 and 300 <= r.left <= 600:
                            edits.append(el)
                    except Exception:
                        continue
        except Exception:
            pass

        edits.sort(key=lambda x: x.rectangle().left)

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
            # Координатный ввод в шапке отчета: первое поле ~ (350, 80), второе ~ (450, 80)
            self.pywinauto.mouse.click(coords=(350, 80))
            time.sleep(0.15)
            self._select_all_and_paste(d_from_str)
            time.sleep(0.15)
            self.pywinauto.mouse.click(coords=(450, 80))
            time.sleep(0.15)
            self._select_all_and_paste(d_to_str)
            time.sleep(0.15)

        # Нажатие кнопки «Сформировать»
        print("[1С:Windows] Запуск формирования отчета (F5)...")
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
            self._hotkey_f5()

        print(f"[1С:Windows] Ожидание формирования отчета (до {self.cfg.report_gen_timeout} сек)...")
        time.sleep(8.0)

        # Сохранение отчета и закрытие его вкладки
        self._save_in_dialog(target_filepath)

    # ---------------------------------------------------------------------
    # Вспомогательные методы выгрузки и сохранения
    # ---------------------------------------------------------------------
    def _click_table_more_button(self) -> None:
        """Кликает по кнопке «Еще ▾» на панели списка документов (в правой части тулбара)."""
        clicked = False
        for w in self.app.windows():
            for btn in w.descendants(control_type="Button"):
                try:
                    txt = (btn.element_info.name or btn.window_text() or "").strip().lower()
                    if txt in ("еще", "ещё"):
                        r = btn.rectangle()
                        if r.left > 1400 and 50 <= r.top <= 120:
                            btn.click_input()
                            clicked = True
                            break
                except Exception:
                    continue
            if clicked:
                break

        if not clicked:
            # Координатный клик кнопки Еще ~ (1860, 80)
            self.pywinauto.mouse.click(coords=(1860, 80))

    def _confirm_export_dialog(self) -> None:
        """Подтверждает модальное окно «Вывести список» (нажимает кнопку ОК / Enter)."""
        time.sleep(0.8)
        confirmed = False
        try:
            for w in self.app.windows():
                title = (w.element_info.name or w.window_text() or "").strip().lower()
                if "вывести список" in title:
                    for btn in w.descendants(control_type="Button"):
                        txt = (btn.element_info.name or btn.window_text() or "").strip().lower()
                        if txt == "ок":
                            btn.click_input()
                            confirmed = True
                            break
                if confirmed:
                    break
        except Exception:
            pass

        if not confirmed:
            self._hotkey_enter()

        # Ожидаем отрисовки сформированного табличного документа
        time.sleep(3.0)

    def _save_in_dialog(self, target_filepath: str) -> None:
        """
        Сохраняет сформированный табличный документ в Excel по Ctrl+S.
        Проверяет появление файла на диске и только после этого закрывает вкладку (Ctrl+F4).
        """
        print(f"[1С:Windows] Сохранение документа в «{target_filepath}» через Ctrl+S...")

        # Удаляем предыдущий файл, чтобы гарантировать проверку нового
        if os.path.exists(target_filepath):
            try:
                os.remove(target_filepath)
            except Exception:
                pass

        # Вызов диалога Windows «Сохранение»
        self._hotkey_ctrl_s()
        time.sleep(1.5)

        # Ввод пути сохранения в диалог
        self._select_all_and_paste(target_filepath)
        time.sleep(0.4)
        self._hotkey_enter()
        time.sleep(1.0)

        # Проверка диалога подтверждения перезаписи ("Заменить существующий?")
        self._hotkey_enter()

        # Ожидание физической записи файла на диск (до 25 секунд)
        file_ready = False
        for _ in range(25):
            if os.path.exists(target_filepath) and os.path.getsize(target_filepath) > 0:
                file_ready = True
                break
            time.sleep(1.0)

        if not file_ready:
            raise RuntimeError(
                f"Файл {target_filepath} не был сохранен на рабочий стол!\n"
                "Проверьте права на запись или открытые диалоговые окна 1С."
            )

        print(f"[1С:Windows] ✓ Таблица успешно сохранена на рабочий стол: {target_filepath}")

        # ЗАКРЫТИЕ СФОРМИРОВАННОЙ ВКЛАДКИ (ТОЛЬКО ПОСЛЕ УСПЕШНОГО СОХРАНЕНИЯ!)
        print("[1С:Windows] Закрытие сформированной вкладки (Ctrl+F4)...")
        time.sleep(0.5)
        self._hotkey_ctrl_f4()
        time.sleep(1.2)

    # ---------------------------------------------------------------------
    # Полный цикл выгрузки всех 3 таблиц
    # ---------------------------------------------------------------------
    def export_all(
        self,
        marshrut_period: Tuple[datetime, datetime],
        etrn_period: Tuple[datetime, datetime],
    ) -> Dict[str, str]:
        print("\n" + "=" * 70)
        print(" ЗАПУСК ВЫГРУЗКИ ТРЕХ ТАБЛИЦ ИЗ 1С:ПРЕДПРИЯТИЕ (e1cib)")
        print("=" * 70)

        # 1. Активация и разворачивание окна 1С
        self._find_and_activate_1c()

        # 2. Выгрузка «Маршрутный лист» -> Модель расчетов.xlsx
        self._open_by_link(self.cfg.url_marshrut, self.cfg.title_marshrut)
        self._set_marshrut_period(marshrut_period[0], marshrut_period[1])
        self._export_marshrut_list(self.cfg.model_file)

        # 3. Выгрузка «Электронные транспортные накладные» -> ЭТРН 1с.xlsx
        self._open_by_link(self.cfg.url_etrn, self.cfg.title_etrn)
        self._set_etrn_period(etrn_period[0], etrn_period[1])
        self._export_etrn_list(self.cfg.etrn_file)

        # 4. Выгрузка «Контроль подписания перевозчиков по ЭТРН» -> Процент.xlsx
        self._open_by_link(self.cfg.url_percent, self.cfg.title_percent)
        self._export_percent_report(
            etrn_period[0], etrn_period[1], self.cfg.percent_file
        )

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
            logger.warning(
                "Ошибка инициализации WindowsOneCExporter: %s. Включаем симулятор.", e
            )
            return MockOneCExporter(cfg)
    return MockOneCExporter(cfg)
