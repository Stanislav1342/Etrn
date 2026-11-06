"""
api_router.py — Готовый FastAPI Router для интеграции отчетов ЭТРН в веб-сервер.

Подключение в server.py (буквально две строчки):
    from etrn_service import EtrnService
    from api_router import router as etrn_router
    app.include_router(etrn_router)

Предоставляет API:
1. POST /api/etrn/process-files — Мгновенная сборка из 3 загруженных файлов (Drag-and-Drop в браузере).
2. POST /api/etrn/generate-task — Фоновый запуск автовыгрузки из 1С по датам.
3. GET  /api/etrn/task-status/{task_id} — Опрос прогресса формирования (0-100%).
4. GET  /api/etrn/download/{task_id} — Скачивание сформированного Excel-отчета.
"""

from __future__ import annotations

import os
import threading
import uuid
from datetime import datetime
from typing import Any, Dict, Optional

from fastapi import APIRouter, File, Form, HTTPException, Response, UploadFile, status
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from etrn_service import EtrnService


router = APIRouter(prefix="/api/etrn", tags=["ЭТРН"])
service = EtrnService()

# Потокобезопасное хранилище фоновых задач
_TASKS_LOCK = threading.Lock()
_TASKS_STORE: Dict[str, Dict[str, Any]] = {}


class GenerateReportRequest(BaseModel):
    """Тело запроса на фоновое формирование отчета."""
    date_from: str = Field(..., description="Дата начала (ДД.ММ.ГГГГ)", example="03.11.2026")
    date_to: Optional[str] = Field(None, description="Дата окончания (ДД.ММ.ГГГГ). Если не указана — равна date_from", example="05.11.2026")


# =============================================================================
# 1. Сборка отчета из файлов (Multipart Upload)
# =============================================================================
@router.post(
    "/process-files",
    summary="Собрать отчет из 3 файлов Excel",
    description="Принимает файлы «ЭТРН», «Модель расчетов» и «Процент» и возвращает готовый сводный файл XLSX.",
)
async def process_files(
    etrn_file: UploadFile = File(..., description="Файл «ЭТРН 1с.xlsx»"),
    model_file: UploadFile = File(..., description="Файл «Модель расчетов.xlsx»"),
    percent_file: UploadFile = File(..., description="Файл «Процент.xlsx»"),
    date_label: Optional[str] = Form(None, description="Метка даты для имени файла (напр. '04.11.2026')"),
):
    try:
        etrn_bytes = await etrn_file.read()
        model_bytes = await model_file.read()
        percent_bytes = await percent_file.read()

        # Генерация в памяти
        report_bytes = service.build_report_from_sources(
            etrn_source=etrn_bytes,
            model_source=model_bytes,
            percent_source=percent_bytes,
        )

        # Имя файла для скачивания
        if date_label and date_label.strip():
            label = date_label.strip().replace(" ", "_")
            out_filename = f"ЭТРН_{label}.xlsx"
        else:
            today_str = datetime.now().strftime("%d.%m.%Y")
            out_filename = f"ЭТРН_{today_str}.xlsx"

        return Response(
            content=report_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f'attachment; filename="{out_filename}"',
                "Access-Control-Expose-Headers": "Content-Disposition",
            },
        )
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Ошибка при обработке файлов: {str(e)}",
        )


# =============================================================================
# 2. Фоновая автовыгрузка из 1С и сборка отчета
# =============================================================================
def _run_background_etrn_task(task_id: str, d_from: datetime, d_to: datetime) -> None:
    """Воркер фонового потока для выгрузки из 1С."""
    def progress_callback(pct: int, msg: str) -> None:
        with _TASKS_LOCK:
            if task_id in _TASKS_STORE:
                _TASKS_STORE[task_id]["progress"] = pct
                _TASKS_STORE[task_id]["message"] = msg

    try:
        result = service.generate_full_report(
            date_from=d_from,
            date_to=d_to,
            progress_cb=progress_callback,
        )
        with _TASKS_LOCK:
            _TASKS_STORE[task_id].update({
                "status": "completed",
                "progress": 100,
                "message": "Отчет успешно сформирован!",
                "filename": result["filename"],
                "file_path": result["output_filepath"],
                "completed_at": datetime.now().isoformat(),
            })
    except Exception as e:
        with _TASKS_LOCK:
            _TASKS_STORE[task_id].update({
                "status": "failed",
                "progress": 100,
                "error": str(e),
                "message": f"Ошибка формирования: {str(e)}",
                "failed_at": datetime.now().isoformat(),
            })


@router.post(
    "/generate-task",
    summary="Запустить фоновую автовыгрузку отчета по ЭТРН",
    description="Ставит задачу в очередь фонового потока. Возвращает task_id для отслеживания статуса.",
)
async def generate_task(req: GenerateReportRequest):
    try:
        d_from = service.parse_date(req.date_from)
        d_to = service.parse_date(req.date_to) if req.date_to else d_from
        if d_from > d_to:
            d_from, d_to = d_to, d_from
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Ошибка в формате дат: {e}")

    task_id = str(uuid.uuid4())
    task_info = {
        "task_id": task_id,
        "status": "running",
        "progress": 0,
        "message": "Инициализация задачи...",
        "date_from": d_from.strftime("%d.%m.%Y"),
        "date_to": d_to.strftime("%d.%m.%Y"),
        "created_at": datetime.now().isoformat(),
    }

    with _TASKS_LOCK:
        _TASKS_STORE[task_id] = task_info

    # Запуск в фоновом потоке
    th = threading.Thread(
        target=_run_background_etrn_task,
        args=(task_id, d_from, d_to),
        daemon=True,
    )
    th.start()

    return {"status": "accepted", "task_id": task_id, "message": "Задача успешно запущена"}


@router.get(
    "/task-status/{task_id}",
    summary="Получить статус фоновой задачи",
)
async def get_task_status(task_id: str):
    with _TASKS_LOCK:
        task = _TASKS_STORE.get(task_id)

    if not task:
        raise HTTPException(status_code=404, detail="Задача не найдена")

    return task


@router.get(
    "/download/{task_id}",
    summary="Скачать сформированный Excel-отчет",
)
async def download_task_result(task_id: str):
    with _TASKS_LOCK:
        task = _TASKS_STORE.get(task_id)

    if not task:
        raise HTTPException(status_code=404, detail="Задача не найдена")

    if task.get("status") != "completed":
        raise HTTPException(status_code=400, detail="Задача еще не завершена или завершилась с ошибкой")

    file_path = task.get("file_path")
    if not file_path or not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Файл отчета не найден на диске")

    filename = task.get("filename", os.path.basename(file_path))

    return FileResponse(
        path=file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename,
    )
