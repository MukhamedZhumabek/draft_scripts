import json
import sys

from rag import RAG


file_path = sys.argv[1]

blocks = []

with open(file_path, "r", encoding="utf-8") as file:
    for line in file:
        line = line.strip()

        if not line:
            continue

        block = json.loads(line)
        blocks.append(block)


rag = RAG()
rag.save_to_db(blocks)