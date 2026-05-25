import os
import uuid
import requests

from dotenv import load_dotenv
from openai import OpenAI
from qdrant_client import QdrantClient
from qdrant_client.models import Distance, VectorParams, PointStruct


load_dotenv()


class RAG:
    def __init__(self):
        self.collection = os.getenv("QDRANT_COLLECTION", "knowledge_base")
        self.vector_size = int(os.getenv("VECTOR_SIZE", "1024"))

        self.bge_url = os.getenv("BGE_API_URL")
        self.bge_key = os.getenv("BGE_API_KEY")

        self.qdrant = QdrantClient(
            url=os.getenv("QDRANT_URL", "http://localhost:6333")
        )

        self.llm = OpenAI(
            api_key=os.getenv("LLM_API_KEY"),
            base_url=os.getenv("LLM_BASE_URL"),
        )

        self.llm_model = os.getenv("LLM_MODEL")

    def get_embedding(self, text: str):
        headers = {}

        if self.bge_key:
            headers["Authorization"] = f"Bearer {self.bge_key}"

        response = requests.post(
            self.bge_url,
            json={"texts": [text]},
            headers=headers,
            timeout=60,
        )

        response.raise_for_status()

        data = response.json()

        return data["dense_vecs"][0]

    def save_to_db(self, blocks: list[dict]):
        if self.qdrant.collection_exists(self.collection):
            self.qdrant.delete_collection(self.collection)

        self.qdrant.create_collection(
            collection_name=self.collection,
            vectors_config=VectorParams(
                size=self.vector_size,
                distance=Distance.COSINE,
            ),
        )

        points = []

        for block in blocks:
            text = block.get("text", "").strip()

            if not text:
                continue

            block_id = block["block_id"]

            point = PointStruct(
                id=str(uuid.uuid5(uuid.NAMESPACE_DNS, block_id)),
                vector=self.get_embedding(text),
                payload={
                    "block_id": block.get("block_id"),
                    "document_id": block.get("document_id"),
                    "source_file": block.get("source_file"),
                    "source_file_name": block.get("source_file_name"),
                    "file_extension": block.get("file_extension"),
                    "block_index": block.get("block_index"),
                    "type": block.get("type"),
                    "heading": block.get("heading"),
                    "heading_path": block.get("heading_path"),
                    "page": block.get("page"),
                    "paragraph": block.get("paragraph"),
                    "slide": block.get("slide"),
                    "sheet": block.get("sheet"),
                    "row_start": block.get("row_start"),
                    "row_end": block.get("row_end"),
                    "text": text,
                },
            )

            points.append(point)

        self.qdrant.upsert(
            collection_name=self.collection,
            points=points,
        )

        print(f"Saved blocks: {len(points)}")

    def search(self, question: str, limit: int = 5):
        question_vector = self.get_embedding(question)

        result = self.qdrant.query_points(
            collection_name=self.collection,
            query=question_vector,
            limit=limit,
            with_payload=True,
        )

        return [point.payload for point in result.points]

    def answer(self, question: str):
        blocks = self.search(question)

        context = ""

        for block in blocks:
            context += f"""
Файл: {block.get("source_file_name")}
Тип блока: {block.get("type")}
Заголовок: {block.get("heading")}
Страница: {block.get("page")}
Слайд: {block.get("slide")}
Лист Excel: {block.get("sheet")}
Текст:
{block.get("text")}

---
"""

        response = self.llm.chat.completions.create(
            model=self.llm_model,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Ты помощник поддержки. "
                        "Отвечай только на основе контекста. "
                        "Если в контексте нет ответа, скажи, что в базе знаний нет информации."
                    ),
                },
                {
                    "role": "user",
                    "content": f"Контекст:\n{context}\n\nВопрос:\n{question}",
                },
            ],
            temperature=0.2,
        )

        return response.choices[0].message.content