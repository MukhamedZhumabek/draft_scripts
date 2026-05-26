#!/usr/bin/env python3

import json
from pathlib import Path


LOGS_DIR = Path("./logs")          # папка с .log файлами
OUTPUT_FILE = Path("all_logs.json")


def main():
    result = []
    seen_request_ids = set()

    total_files = 0
    total_lines = 0
    written_count = 0
    duplicate_count = 0
    invalid_json_count = 0
    missing_fields_count = 0

    for log_file in LOGS_DIR.glob("*.log"):
        total_files += 1

        with log_file.open("r", encoding="utf-8") as f:
            for line_number, line in enumerate(f, start=1):
                total_lines += 1
                line = line.strip()

                if not line:
                    continue

                try:
                    obj = json.loads(line)
                except json.JSONDecodeError:
                    invalid_json_count += 1
                    print(f"[WARN] invalid json: {log_file}:{line_number}")
                    continue

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

    with OUTPUT_FILE.open("w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print("Готово")
    print(f"Папка: {LOGS_DIR}")
    print(f"Файлов обработано: {total_files}")
    print(f"Строк прочитано: {total_lines}")
    print(f"Записано уникальных request_id: {written_count}")
    print(f"Дубликатов пропущено: {duplicate_count}")
    print(f"Невалидных JSON строк: {invalid_json_count}")
    print(f"Без request_id или timestamp: {missing_fields_count}")
    print(f"Файл результата: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()