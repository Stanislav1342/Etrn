"""
config.py — Конфигурация путей, разделов 1С и параметров отчета ЭТРН.
"""

import os
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Tuple


@dataclass
class Config:
    """Конфигурация проекта ЭТРН."""

    # Папка на рабочем столе для входных и выходных файлов
    desktop_dir: str = field(
        default_factory=lambda: os.path.join(os.path.expanduser("~"), "Desktop")
    )
    output_dir_name: str = "ЭТРН"

    # Имена выгружаемых из 1С файлов
    input_etrn_name: str = "ЭТРН 1с.xlsx"
    input_model_name: str = "Модель расчетов.xlsx"
    input_percent_name: str = "Процент.xlsx"

    # Точные названия пунктов в панели «Избранное» 1С
    fav_marshrut: str = "Маршрутный лист"
    fav_etrn: str = "Электронные транспортные накладные"
    fav_percent: str = "Контроль подписания перевозчиков по ЭТРН"

    # Заголовок главного окна 1С (регулярное выражение)
    onec_window_regex: str = ".*(1С:Предприятие|ЕИС МСК|Корпоративная информационная система).*"

    # Паузы и таймауты (секунды)
    action_pause: float = 1.0
    report_gen_timeout: int = 40

    @property
    def output_dir(self) -> str:
        return os.path.join(self.desktop_dir, self.output_dir_name)

    @property
    def etrn_file(self) -> str:
        return os.path.join(self.desktop_dir, self.input_etrn_name)

    @property
    def model_file(self) -> str:
        return os.path.join(self.desktop_dir, self.input_model_name)

    @property
    def percent_file(self) -> str:
        return os.path.join(self.desktop_dir, self.input_percent_name)

    def get_output_file(self, date_label: str) -> str:
        return os.path.join(self.output_dir, f"ЭТРН_{date_label}.xlsx")

    @staticmethod
    def get_marshrut_period() -> Tuple[datetime, datetime]:
        """Для «Маршрутного листа» период: сегодня и 30 дней до него."""
        today = datetime.now()
        start = today - timedelta(days=30)
        return start, today


def get_config() -> Config:
    return Config()
