"""
update.py — Скрипт обновления проекта «Этрн» с GitHub без использования git.exe.

Используется на корпоративных компьютерах, где запуск git.exe
заблокирован групповыми политиками Windows (AppLocker / WinError 1260).

Запуск:
    python update.py
"""

import io
import os
import urllib.request
import zipfile

REPO_ZIP_URL = "https://github.com/Stanislav1342/Etrn/archive/refs/heads/main.zip"
PROTECTED_FILES = {
    ".env",
    "ЭТРН 1с.xlsx",
    "Модель расчетов.xlsx",
    "Процент.xlsx",
}
PROTECTED_EXTENSIONS = {
    ".xlsx",
    ".xls",
    ".csv",
}


def update_from_github() -> None:
    print("=" * 68)
    print("  🔄 ОБНОВЛЕНИЕ ПРОЕКТА «ЭТРН ОПТИМИЗАТОР» С GITHUB (БЕЗ GIT.EXE)")
    print("=" * 68)
    print(f"Репозиторий: {REPO_ZIP_URL}")
    print("Скачивание актуальной версии...")

    req = urllib.request.Request(
        REPO_ZIP_URL,
        headers={"User-Agent": "Etrn-Updater/1.0"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            zip_bytes = resp.read()
    except Exception as e:
        print(f"\n[ОШИБКА] Не удалось скачать обновление: {e}")
        print("Проверьте интернет-соединение и доступность репозитория GitHub.")
        return

    base_dir = os.path.dirname(os.path.abspath(__file__))
    updated_count = 0

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for member in zf.infolist():
            if member.is_dir():
                continue

            # Убираем корневую папку архива (Etrn-main/)
            parts = member.filename.split("/", 1)
            if len(parts) < 2 or not parts[1]:
                continue
            rel_path = parts[1]
            filename = os.path.basename(rel_path)
            _, ext = os.path.splitext(filename)

            if filename in PROTECTED_FILES or ext.lower() in PROTECTED_EXTENSIONS:
                continue

            dest_path = os.path.join(base_dir, rel_path)
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)

            with zf.open(member) as src, open(dest_path, "wb") as dst:
                dst.write(src.read())
            updated_count += 1
            print(f"  [Обновлен] {rel_path}")

    print("=" * 68)
    print(f"  ✅ Обновление успешно завершено! Обновлено файлов: {updated_count}")
    print("  🔒 Ваши локальные отчеты, файлы Excel и .env сохранены без изменений.")
    print("=" * 68)


if __name__ == "__main__":
    update_from_github()
