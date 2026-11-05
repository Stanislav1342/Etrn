"""
run.py — Главный запускаемый скрипт полного цикла отчета по ЭТРН:
  1. Автоматическая выгрузка 3 таблиц из 1С:Предприятие:
     - «Маршрутный лист» (период: сегодня и 30 дней назад) -> Модель расчетов.xlsx
     - «Электронные транспортные накладные» (выбранный период) -> ЭТРН 1с.xlsx
     - «Контроль подписания перевозчиков по ЭТРН» (выбранный период) -> Процент.xlsx
  2. Обработка, сведение данных и формирование стилизованного Excel-отчета с двумя листами:
     - 'ЭТРН' — детальная таблица с неподписанными титулами и статусами
     - 'Отчет' — свод по перевозчикам, водителям, моделям расчетов и среднему времени

Запуск:
    python run.py                     # Единый ввод даты (04.11.2026 или 03.11.2026-05.11.2026)
    python run.py --auto              # Автоматический режим (вчерашний день)
    python run.py --skip-1c           # Без 1С, только обработка готовых файлов на рабочем столе
    python run.py --date 04.11.2026   # Конкретная дата
"""

import argparse
from datetime import datetime, timedelta
import os
import sys
from typing import Tuple

from config import Config, get_config
from etrn_optimizer import process_and_style_etrn
from onec_exporter import get_onec_exporter


def parse_date_str(s: str) -> datetime:
    """Парсинг даты из строки формата ДД.ММ.ГГГГ или ДД.ММ."""
    s = s.strip()
    parts = s.split(".")
    if len(parts) == 2:
        s = f"{s}.{datetime.now().year}"
    elif len(parts) == 3 and len(parts[2]) == 2:
        s = f"{parts[0]}.{parts[1]}.20{parts[2]}"
    try:
        return datetime.strptime(s, "%d.%m.%Y")
    except ValueError:
        raise argparse.ArgumentTypeError(
            f"Неверный формат даты '{s}'. Используйте ДД.ММ.ГГГГ (например, 04.11.2026)"
        )


def parse_user_date_input(raw: str) -> Tuple[datetime, datetime, str]:
    """
    Единый парсер даты или диапазона дат:
    - Пустой ввод (Enter) -> вчерашний день
    - Одна дата (например, 04.11.2026) -> один этот день
    - Диапазон через дефис (например, 03.11.2026-05.11.2026) -> диапазон с и по
    """
    raw = raw.strip()
    today = datetime.now()
    yesterday = today - timedelta(days=1)
    yesterday_str = yesterday.strftime("%d.%m.%Y")

    if not raw:
        return yesterday, yesterday, yesterday_str

    if "-" in raw:
        parts = raw.split("-", 1)
        d_from = parse_date_str(parts[0])
        d_to = parse_date_str(parts[1])
        label = f"{d_from.strftime('%d.%m.%Y')}_{d_to.strftime('%d.%m.%Y')}"
        return d_from, d_to, label
    else:
        d = parse_date_str(raw)
        return d, d, d.strftime("%d.%m.%Y")


def prompt_period_from_user() -> Tuple[datetime, datetime, str]:
    """Единый интерактивный запрос даты или диапазона."""
    yesterday = datetime.now() - timedelta(days=1)
    yesterday_str = yesterday.strftime("%d.%m.%Y")

    print("\n" + "=" * 70)
    print("  ВВЕДИТЕ ДАТУ ИЛИ ДИАПАЗОН ДАТ ДЛЯ ЭТРН И ОТЧЕТА ПОДПИСАНИЯ")
    print("=" * 70)
    print(f"Примеры ввода:")
    print(f" • Одна дата:     04.11.2026")
    print(f" • Диапазон дат:  03.11.2026-05.11.2026")
    print(f" • Нажмите Enter для вчерашней даты [{yesterday_str}]")
    print("-" * 70)

    raw = input("Дата или диапазон (Enter для вчера): ").strip()
    return parse_user_date_input(raw)


def main():
    parser = argparse.ArgumentParser(
        description="Автоматическая выгрузка данных из 1С и построение отчета по ЭТРН"
    )
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Запуск без диалога с вчерашней датой (для планировщиков задач)",
    )
    parser.add_argument(
        "--skip-1c",
        action="store_true",
        help="Пропустить выгрузку из 1С и собрать отчет из существующих файлов на рабочем столе",
    )
    parser.add_argument(
        "--date",
        type=str,
        default=None,
        help="Дата или диапазон дат (например, 04.11.2026 или 03.11.2026-05.11.2026)",
    )

    args = parser.parse_args()
    cfg = get_config()

    print("\n" + "=" * 70)
    print("     🚀 АВТОМАТИЗАЦИЯ СБОРКИ СВОДНОГО ОТЧЕТА ПО ЭТРН ИЗ 1С")
    print("=" * 70)

    # 1. Определение периодов
    marshrut_start, marshrut_end = cfg.get_marshrut_period()

    if args.date:
        etrn_start, etrn_end, date_label = parse_user_date_input(args.date)
    elif args.auto:
        yesterday = datetime.now() - timedelta(days=1)
        etrn_start, etrn_end = yesterday, yesterday
        date_label = yesterday.strftime("%d.%m.%Y")
    else:
        etrn_start, etrn_end, date_label = prompt_period_from_user()

    print(f"Период для «Маршрутный лист»: с {marshrut_start.strftime('%d.%m.%Y')} по {marshrut_end.strftime('%d.%m.%Y')}")
    print(f"Период для «ЭТРН» и отчета:    с {etrn_start.strftime('%d.%m.%Y')} по {etrn_end.strftime('%d.%m.%Y')}")
    print(f"Имя итогового отчета:          ЭТРН_{date_label}.xlsx")
    print("-" * 70)

    # 2. Выгрузка из 1С:Предприятие
    if not args.skip_1c:
        exporter = get_onec_exporter(cfg)
        try:
            exporter.export_all(
                marshrut_period=(marshrut_start, marshrut_end),
                etrn_period=(etrn_start, etrn_end),
            )
        except Exception as e:
            print(f"\n[ОШИБКА 1С] Не удалось завершить выгрузку из 1С: {e}")
            print("Убедитесь, что 1С запущена, активна и содержит разделы в «Избранном».")
            if sys.platform == "win32":
                sys.exit(1)

    # 3. Проверка наличия выгруженных файлов
    print("[Обработка] Проверка наличия исходных файлов на рабочем столе...")
    missing_files = []
    for fpath in (cfg.model_file, cfg.etrn_file, cfg.percent_file):
        if not os.path.exists(fpath):
            missing_files.append(fpath)

    if missing_files:
        print(f"\n[ВНИМАНИЕ] Не найдены следующие файлы на рабочем столе:")
        for mf in missing_files:
            print(f"  • {mf}")
        if sys.platform != "win32":
            print("\n[macOS Режим разработки] Вы находитесь на macOS, реальные файлы 1С не созданы.")
            print("На корпоративном ПК под Windows файлы будут автоматически выгружены из 1С.")
            return
        else:
            print("\nПрерывание: отчет не может быть собран без всех 3 файлов.")
            sys.exit(1)

    # 4. Обработка данных и формирование отчета через etrn_optimizer
    os.makedirs(cfg.output_dir, exist_ok=True)
    output_filepath = cfg.get_output_file(date_label)

    print("\n" + "=" * 70)
    print("      📊 ОБРАБОТКА ДАННЫХ И СТРОЕНИЕ СТИЛИЗОВАННОГО ОТЧЕТА")
    print("=" * 70)

    process_and_style_etrn(
        input_filepath=cfg.etrn_file,
        model_filepath=cfg.model_file,
        percent_filepath=cfg.percent_file,
        output_filepath=output_filepath,
    )

    print("=" * 70)
    print(f"  ✅ ИТОГОВЫЙ ОТЧЕТ УСПЕШНО ГОТОВ: {output_filepath}")
    print("=" * 70 + "\n")


if __name__ == "__main__":
    main()
