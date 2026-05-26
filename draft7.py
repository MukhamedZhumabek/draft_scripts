from __future__ import annotations

from pathlib import Path
from typing import Any
import json
import re


# ============================================================
# CONFIG
# ============================================================

INPUT_BLOCKS_PATH = Path(r"C:\rag_json\all_blocks.jsonl")
OUTPUT_CHUNKS_PATH = Path(r"C:\rag_json\fat_chunks.jsonl")

TARGET_CHARS = 3500
MAX_CHARS = 6000
OVERLAP_CHARS = 500

KEEP_FILE_NAME_IN_TEXT = True


# ============================================================
# CLEAN
# ============================================================

def clean_for_embedding(text: str) -> str:
    text = text.replace("\x00", " ")
    text = text.replace("\r", "\n")

    # Убрать много пробелов
    text = re.sub(r"[ \t]+", " ", text)

    # Убрать пробелы вокруг переносов
    text = re.sub(r" *\n *", "\n", text)

    # Убрать слишком много пустых строк
    text = re.sub(r"\n{3,}", "\n\n", text)

    # Убрать мусорные повторяющиеся символы
    text = re.sub(r"[_]{3,}", " ", text)
    text = re.sub(r"[-]{5,}", " ", text)
    text = re.sub(r"[.]{5,}", " ", text)

    return text.strip()


def normalize_block_text(block: dict[str, Any]) -> str:
    text = clean_for_embedding(block.get("text", ""))

    if not text:
        return ""

    block_type = block.get("type")

    if block_type == "heading":
        return f"\n{text}\n"

    if block_type == "title":
        return f"\n{text}\n"

    if block_type == "list_item":
        # если пункт уже начинается с 1. или -, не добавляем лишнее
        if re.match(r"^([-–—•]|\d+[\.\)])\s+", text):
            return text
        return f"- {text}"

    if block_type == "table":
        return f"Таблица:\n{text}"

    return text


# ============================================================
# LOAD
# ============================================================

def load_blocks(path: Path) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = []

    with open(path, "r", encoding="utf-8") as file:
        for line in file:
            line = line.strip()

            if not line:
                continue

            block = json.loads(line)

            text = block.get("text", "")
            if not text or not text.strip():
                continue

            blocks.append(block)

    return blocks


def group_by_file(blocks: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    grouped: dict[str, list[dict[str, Any]]] = {}

    for block in blocks:
        file_name = block.get("source_file_name") or "unknown"

        if file_name not in grouped:
            grouped[file_name] = []

        grouped[file_name].append(block)

    for file_blocks in grouped.values():
        file_blocks.sort(key=lambda b: b.get("block_index", 0))

    return grouped


# ============================================================
# CHUNKING
# ============================================================

def make_chunk_text(source_file_name: str, parts: list[str]) -> str:
    body = clean_for_embedding("\n\n".join(parts))

    if KEEP_FILE_NAME_IN_TEXT:
        return f"Файл: {source_file_name}\n\n{body}".strip()

    return body


def split_big_text(text: str, max_chars: int) -> list[str]:
    """
    Если один блок сам огромный, режем его по предложениям/строкам.
    """
    if len(text) <= max_chars:
        return [text]

    pieces = re.split(r"(?<=[.!?。！？])\s+|\n+", text)

    result = []
    current = ""

    for piece in pieces:
        piece = piece.strip()

        if not piece:
            continue

        candidate = f"{current} {piece}".strip() if current else piece

        if len(candidate) <= max_chars:
            current = candidate
        else:
            if current:
                result.append(current)

            current = piece

    if current:
        result.append(current)

    return result


def build_chunks_for_file(source_file_name: str, blocks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    chunks: list[dict[str, Any]] = []

    current_parts: list[str] = []

    def current_text() -> str:
        return make_chunk_text(source_file_name, current_parts)

    def flush() -> None:
        nonlocal current_parts

        if not current_parts:
            return

        text = current_text()

        if not text:
            current_parts = []
            return

        chunks.append(
            {
                "chunk_id": f"{source_file_name}::chunk_{len(chunks):06d}",
                "source_file_name": source_file_name,
                "chunk_index": len(chunks),
                "text": text,
            }
        )

        if OVERLAP_CHARS > 0:
            tail = text[-OVERLAP_CHARS:]
            current_parts = [tail]
        else:
            current_parts = []

    for block in blocks:
        block_text = normalize_block_text(block)

        if not block_text:
            continue

        # если блок сам слишком большой
        block_pieces = split_big_text(block_text, MAX_CHARS)

        for piece in block_pieces:
            candidate_parts = current_parts + [piece]
            candidate_text = make_chunk_text(source_file_name, candidate_parts)

            if len(candidate_text) > MAX_CHARS and current_parts:
                flush()

            current_parts.append(piece)

            if len(current_text()) >= TARGET_CHARS:
                flush()

    flush()

    return chunks


# ============================================================
# SAVE
# ============================================================

def save_chunks(chunks: list[dict[str, Any]], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    with open(path, "w", encoding="utf-8") as file:
        for chunk in chunks:
            file.write(json.dumps(chunk, ensure_ascii=False) + "\n")


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    blocks = load_blocks(INPUT_BLOCKS_PATH)
    grouped = group_by_file(blocks)

    all_chunks: list[dict[str, Any]] = []

    for source_file_name, file_blocks in grouped.items():
        chunks = build_chunks_for_file(source_file_name, file_blocks)
        all_chunks.extend(chunks)

        print(
            f"{source_file_name}: "
            f"blocks={len(file_blocks)}, "
            f"chunks={len(chunks)}"
        )

    save_chunks(all_chunks, OUTPUT_CHUNKS_PATH)

    print()
    print(f"SAVED: {OUTPUT_CHUNKS_PATH}")
    print(f"TOTAL CHUNKS: {len(all_chunks)}")


if __name__ == "__main__":
    main()