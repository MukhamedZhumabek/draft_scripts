#!/usr/bin/env python3

import argparse
import json
from pathlib import Path


def iter_json_objects(file_path: Path):
    """
    Поддерживает два формата:
    1. Весь файл — один JSON объект: {...}
    2. Каждая строка — отдельный JSON объект: {...}
    """

    text = file_path.read_text(encoding="utf-8").strip()

    if not text:
        return

    # Пробуем прочитать весь файл как один JSON
    try:
        data = json.loads(text)

        if isinstance(data, dict):
            yield data
        elif isinstance(data, list):
            for item in data:
                if isinstance(item, dict):
                    yield item

        return

    except json.JSONDecodeError:
        pass

    # Если весь файл не JSON — читаем построчно
    with file_path.open("r", encoding="utf-8") as f:
        for line_number, line in enumerate(f, start=1):
            line = line.strip()

            if not line:
                continue

            try:
                data = json.loads(line)

                if isinstance(data, dict):
                    yield data

            except json.JSONDecodeError:
                print(f"[WARN] Невалидный JSON: {file_path}:{line_number}")


def main():
    parser = argparse.ArgumentParser(
        description="Collect request_id and timestamp from .log JSON files"
    )

    parser.add_argument(
        "logs_dir",
        help="Папка с .log файлами"
    )

    parser.add_argument(
        "-o",
        "--output",
        default="result.json",
        help="Имя выходного JSON файла, по умолчанию result.json"
    )

    args = parser.parse_args()

    logs_dir = Path(args.logs_dir)
    output_file = Path(args.output)

    if not logs_dir.exists():
        print(f"[ERROR] Папка не существует: {logs_dir}")
        return

    if not logs_dir.is_dir():
        print(f"[ERROR] Это не папка: {logs_dir}")
        return

    result = []
    seen_request_ids = set()

    total_files = 0
    total_json_objects = 0
    written_count = 0
    duplicate_count = 0
    missing_fields_count = 0

    for log_file in logs_dir.glob("*.log"):
        total_files += 1

        for obj in iter_json_objects(log_file):
            total_json_objects += 1

            request_id = obj.get("request_id")
            timestamp = obj.get("timestamp")

            if not request_id or not timestamp:
                missing_fields_count += 1
                continue

            if request_id in seen_request_ids:
                duplicate_count += 1
                continue

            seen_request_ids.add(request_id)

            result.append({
                "request_id": request_id,
                "timestamp": timestamp
            })

            written_count += 1

    with output_file.open("w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print("Готово")
    print(f"Файлов обработано: {total_files}")
    print(f"JSON объектов найдено: {total_json_objects}")
    print(f"Записано уникальных записей: {written_count}")
    print(f"Дубликатов request_id пропущено: {duplicate_count}")
    print(f"Без request_id или timestamp пропущено: {missing_fields_count}")
    print(f"Результат сохранён в: {output_file}")


if __name__ == "__main__":
    main()