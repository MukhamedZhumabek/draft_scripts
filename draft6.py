from __future__ import annotations

from pathlib import Path
from typing import Any
import json


# ============================================================
# CONFIG
# ============================================================

INPUT_BLOCKS_PATH = Path(r"C:\rag_json\all_blocks.jsonl")
OUTPUT_CHUNKS_PATH = Path(r"C:\rag_json\chunks.jsonl")

TARGET_CHARS = 2500
MAX_CHARS = 4500
OVERLAP_BLOCKS = 1

SKIP_EMPTY_TEXT = True


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

            if SKIP_EMPTY_TEXT and not block.get("text", "").strip():
                continue

            blocks.append(block)

    return blocks


# ============================================================
# HELPERS
# ============================================================

def get_doc_key(block: dict[str, Any]) -> str:
    return block.get("document_id") or block.get("source_file_name") or block.get("source_file")


def get_document_title(block: dict[str, Any]) -> str:
    return (
        block.get("document_title")
        or block.get("source_file_name")
        or block.get("source_file")
        or "unknown"
    )


def get_heading_text(block: dict[str, Any]) -> str | None:
    heading_path = block.get("heading_path") or []

    if heading_path:
        return " > ".join(str(x) for x in heading_path if x)

    heading = block.get("heading")

    if heading:
        return str(heading)

    return None


def get_location(blocks: list[dict[str, Any]]) -> dict[str, Any]:
    pages = [b.get("page") for b in blocks if b.get("page") is not None]
    paragraphs = [b.get("paragraph") for b in blocks if b.get("paragraph") is not None]
    slides = [b.get("slide") for b in blocks if b.get("slide") is not None]

    sheets = [b.get("sheet") for b in blocks if b.get("sheet")]
    row_starts = [b.get("row_start") for b in blocks if b.get("row_start") is not None]
    row_ends = [b.get("row_end") for b in blocks if b.get("row_end") is not None]

    return {
        "page_start": min(pages) if pages else None,
        "page_end": max(pages) if pages else None,

        "paragraph_start": min(paragraphs) if paragraphs else None,
        "paragraph_end": max(paragraphs) if paragraphs else None,

        "slide_start": min(slides) if slides else None,
        "slide_end": max(slides) if slides else None,

        "sheet": sheets[0] if sheets else None,
        "row_start": min(row_starts) if row_starts else None,
        "row_end": max(row_ends) if row_ends else None,
    }


def block_to_text(block: dict[str, Any]) -> str:
    block_type = block.get("type")
    text = block.get("text", "").strip()

    if not text:
        return ""

    if block_type == "list_item":
        return f"- {text}"

    if block_type == "table":
        return f"Таблица:\n{text}"

    return text


def build_chunk_text(
    document_title: str,
    heading: str | None,
    blocks: list[dict[str, Any]],
) -> str:
    parts = []

    parts.append(f"Документ: {document_title}")

    if heading:
        parts.append(f"Раздел: {heading}")

    body_parts = []

    for block in blocks:
        text = block_to_text(block)

        if text:
            body_parts.append(text)

    parts.append("\n\n".join(body_parts))

    return "\n\n".join(parts).strip()


def make_chunk(
    document_id: str,
    document_title: str,
    chunk_index: int,
    heading: str | None,
    blocks: list[dict[str, Any]],
) -> dict[str, Any]:
    first_block = blocks[0]

    chunk_text = build_chunk_text(
        document_title=document_title,
        heading=heading,
        blocks=blocks,
    )

    location = get_location(blocks)

    return {
        "chunk_id": f"{document_id}::chunk_{chunk_index:06d}",
        "document_id": document_id,
        "document_title": document_title,
        "source_file": first_block.get("source_file"),
        "source_file_name": first_block.get("source_file_name"),
        "file_extension": first_block.get("file_extension"),

        "chunk_index": chunk_index,
        "heading": heading,

        "page_start": location["page_start"],
        "page_end": location["page_end"],
        "paragraph_start": location["paragraph_start"],
        "paragraph_end": location["paragraph_end"],
        "slide_start": location["slide_start"],
        "slide_end": location["slide_end"],
        "sheet": location["sheet"],
        "row_start": location["row_start"],
        "row_end": location["row_end"],

        "block_ids": [b.get("block_id") for b in blocks],
        "block_indexes": [b.get("block_index") for b in blocks],

        "text": chunk_text,
    }


# ============================================================
# CHUNKING
# ============================================================

def build_chunks_for_document(blocks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not blocks:
        return []

    document_id = get_doc_key(blocks[0])
    document_title = get_document_title(blocks[0])

    chunks: list[dict[str, Any]] = []

    current_blocks: list[dict[str, Any]] = []
    current_heading: str | None = None

    def flush() -> None:
        nonlocal current_blocks

        content_blocks = [
            b for b in current_blocks
            if b.get("type") not in {"title", "heading"}
        ]

        if not content_blocks:
            current_blocks = []
            return

        chunk = make_chunk(
            document_id=document_id,
            document_title=document_title,
            chunk_index=len(chunks),
            heading=current_heading,
            blocks=content_blocks,
        )

        chunks.append(chunk)

        if OVERLAP_BLOCKS > 0:
            current_blocks = content_blocks[-OVERLAP_BLOCKS:]
        else:
            current_blocks = []

    for block in blocks:
        block_type = block.get("type")
        text = block.get("text", "").strip()

        if not text:
            continue

        if block_type in {"title", "heading"}:
            new_heading = get_heading_text(block) or text

            if current_blocks:
                flush()

            current_heading = new_heading
            continue

        block_heading = get_heading_text(block)

        if block_heading and block_heading != current_heading:
            if current_blocks:
                flush()

            current_heading = block_heading

        candidate_blocks = current_blocks + [block]

        candidate_text = build_chunk_text(
            document_title=document_title,
            heading=current_heading,
            blocks=candidate_blocks,
        )

        if len(candidate_text) > MAX_CHARS and current_blocks:
            flush()

        current_blocks.append(block)

        current_text = build_chunk_text(
            document_title=document_title,
            heading=current_heading,
            blocks=current_blocks,
        )

        if len(current_text) >= TARGET_CHARS:
            flush()

    if current_blocks:
        flush()

    return chunks


def group_blocks_by_document(blocks: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    grouped: dict[str, list[dict[str, Any]]] = {}

    for block in blocks:
        doc_key = get_doc_key(block)

        if doc_key not in grouped:
            grouped[doc_key] = []

        grouped[doc_key].append(block)

    for doc_blocks in grouped.values():
        doc_blocks.sort(key=lambda b: b.get("block_index", 0))

    return grouped


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
    grouped = group_blocks_by_document(blocks)

    all_chunks: list[dict[str, Any]] = []

    for document_id, document_blocks in grouped.items():
        chunks = build_chunks_for_document(document_blocks)
        all_chunks.extend(chunks)

        print(
            f"{document_id}: "
            f"blocks={len(document_blocks)}, "
            f"chunks={len(chunks)}"
        )

    save_chunks(all_chunks, OUTPUT_CHUNKS_PATH)

    print()
    print(f"SAVED: {OUTPUT_CHUNKS_PATH}")
    print(f"TOTAL CHUNKS: {len(all_chunks)}")


if __name__ == "__main__":
    main()