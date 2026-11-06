"""
web_test.py — Локальный веб-сервер для мгновенной проверки работы модуля ЭТРН в веб-режиме.

Запуск:
    python web_test.py

После запуска откройте в браузере:
    http://127.0.0.1:8000
    или интерактивную документацию Swagger:
    http://127.0.0.1:8000/docs
"""

import os
import uvicorn
from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from api_router import router as etrn_router


app = FastAPI(
    title="Тестовый веб-портал: Модуль ЭТРН",
    description="Демонстрация работы отчетов ЭТРН в веб-сервере FastAPI.",
    version="1.0.0",
)

app.include_router(etrn_router)


@app.get("/", response_class=HTMLResponse)
async def home_page():
    """Простая веб-страница для проверки загрузки и скачивания отчета."""
    return """
    <!DOCTYPE html>
    <html lang="ru">
    <head>
        <meta charset="UTF-8">
        <title>Тест модуля ЭТРН в веб-сервере</title>
        <style>
            body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #f0f2f5; margin: 0; padding: 40px; }
            .card { max-width: 650px; margin: 0 auto; background: white; border-radius: 12px; padding: 30px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
            h1 { font-size: 22px; color: #1a1a1a; margin-top: 0; }
            p { color: #555; font-size: 14px; line-height: 1.5; }
            .form-group { margin-bottom: 20px; }
            label { display: block; font-weight: 600; margin-bottom: 6px; font-size: 13px; color: #333; }
            input[type="file"], input[type="text"] { width: 100%; box-sizing: border-box; padding: 8px 12px; border: 1px solid #ccc; border-radius: 6px; font-size: 13px; }
            .btn { background: #0066cc; color: white; border: none; padding: 12px 20px; border-radius: 6px; font-size: 15px; font-weight: 600; cursor: pointer; width: 100%; transition: background 0.2s; }
            .btn:hover { background: #0052a3; }
            .btn:disabled { background: #999; cursor: not-allowed; }
            .status-box { margin-top: 20px; padding: 12px; border-radius: 6px; display: none; font-size: 14px; }
            .status-box.success { background: #e6f4ea; color: #137333; border: 1px solid #ceead6; }
            .status-box.error { background: #fce8e6; color: #c5221f; border: 1px solid #fad2cf; }
            .docs-link { display: block; text-align: center; margin-top: 20px; color: #0066cc; text-decoration: none; font-size: 13px; }
            .docs-link:hover { text-decoration: underline; }
        </style>
    </head>
    <body>
        <div class="card">
            <h1>📑 Проверка работы отчета ЭТРН в веб-сервере</h1>
            <p>Этот тест демонстрирует, как сервер собирает сводный отчет полностью в памяти без открытия окон 1С и эмуляции пользователя.</p>
            
            <form id="upload-form" onsubmit="handleUpload(event)">
                <div class="form-group">
                    <label>1. Файл «ЭТРН 1с.xlsx»:</label>
                    <input type="file" name="etrn_file" accept=".xlsx,.csv" required>
                </div>
                <div class="form-group">
                    <label>2. Файл «Модель расчетов.xlsx»:</label>
                    <input type="file" name="model_file" accept=".xlsx" required>
                </div>
                <div class="form-group">
                    <label>3. Файл «Процент.xlsx»:</label>
                    <input type="file" name="percent_file" accept=".xlsx" required>
                </div>
                <div class="form-group">
                    <label>Метка даты для отчета (необязательно):</label>
                    <input type="text" name="date_label" placeholder="Например: 04.11.2026">
                </div>
                <button type="submit" class="btn" id="submit-btn">⚡ Собрать и скачать отчет</button>
            </form>

            <div id="status" class="status-box"></div>
            <a class="docs-link" href="/docs" target="_blank">📖 Открыть интерактивную документацию Swagger API (/docs)</a>
        </div>

        <script>
            async function handleUpload(e) {
                e.preventDefault();
                const btn = document.getElementById('submit-btn');
                const statusBox = document.getElementById('status');
                
                btn.disabled = true;
                btn.innerText = '⏳ Обработка в памяти сервера...';
                statusBox.style.display = 'none';

                const formData = new FormData(document.getElementById('upload-form'));

                try {
                    const res = await fetch('/api/etrn/process-files', {
                        method: 'POST',
                        body: formData
                    });

                    if (!res.ok) {
                        const err = await res.json();
                        throw new Error(err.detail || 'Ошибка обработки на сервере');
                    }

                    // Скачивание файла
                    const blob = await res.blob();
                    const disposition = res.headers.get('Content-Disposition') || '';
                    let filename = 'ЭТРН_Отчет.xlsx';
                    const match = disposition.match(/filename="?([^";]+)"?/);
                    if (match && match[1]) {
                        filename = match[1];
                    }

                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = filename;
                    document.body.appendChild(a);
                    a.click();
                    a.remove();

                    statusBox.className = 'status-box success';
                    statusBox.innerText = '✓ Отчет успешно сформирован и скачан: ' + filename;
                    statusBox.style.display = 'block';
                } catch (err) {
                    statusBox.className = 'status-box error';
                    statusBox.innerText = '✗ Ошибка: ' + err.message;
                    statusBox.style.display = 'block';
                } finally {
                    btn.disabled = false;
                    btn.innerText = '⚡ Собрать и скачать отчет';
                }
            }
        </script>
    </body>
    </html>
    """


if __name__ == "__main__":
    port = 8000
    print(f"\n[Веб-тест] Запуск локального сервера на http://127.0.0.1:{port}...")
    print(f"[Веб-тест] Откройте в браузере: http://127.0.0.1:{port}\n")
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="info")
