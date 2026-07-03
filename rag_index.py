import logging
import re
from typing import Optional

import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

logger = logging.getLogger(__name__)


class RAGIndex:
    """Lightweight RAG index using TF-IDF + cosine similarity.

    Chunks paper text, indexes with TF-IDF, and retrieves relevant
    chunks for a given query.
    """

    def __init__(self, chunk_size: int = 500, overlap: int = 100):
        self.chunk_size = chunk_size
        self.overlap = overlap
        self.chunks: list[str] = []
        self.vectorizer = TfidfVectorizer(
            max_features=5000,
            stop_words="english",
            ngram_range=(1, 2),
            sublinear_tf=True,
        )
        self.vectors: Optional[np.ndarray] = None
        self._built = False

    def build(self, text: str) -> list[str]:
        """Chunk and index the text."""
        if not text or len(text.strip()) < 50:
            logger.warning("Text too short for RAG indexing")
            self.chunks = [text] if text else []
            self._built = True
            return self.chunks

        self.chunks = self._chunk_text(text)
        if len(self.chunks) == 0:
            self.chunks = [text[:self.chunk_size]] if text else []

        if len(self.chunks) >= 2:
            try:
                self.vectors = self.vectorizer.fit_transform(self.chunks)
            except Exception as e:
                logger.warning(f"RAG vectorization failed: {e}, falling back")
                self.vectors = None
        else:
            self.vectors = None

        self._built = True
        logger.info(f"RAG index built: {len(self.chunks)} chunks")
        return self.chunks

    def retrieve(self, query: str, top_k: int = 3) -> list[dict]:
        """Retrieve top_k most relevant chunks for query.

        Returns sorted list of {chunk, score, index}.
        """
        if not self._built or not self.chunks:
            return []

        if self.vectors is None or len(self.chunks) < 2:
            return self._fallback_retrieve(query, top_k)

        try:
            query_vec = self.vectorizer.transform([query])
            scores = cosine_similarity(query_vec, self.vectors)[0]
        except Exception:
            return self._fallback_retrieve(query, top_k)

        top_indices = scores.argsort()[::-1][:top_k]
        results = []
        for idx in top_indices:
            if scores[idx] > 0.01:
                results.append({
                    "chunk": self.chunks[idx],
                    "score": float(scores[idx]),
                    "index": int(idx),
                })
        return results

    def retrieve_for_slides(self, slide_topics: list[str], top_k: int = 2) -> dict[str, list[dict]]:
        """For each slide topic, retrieve relevant chunks."""
        return {topic: self.retrieve(topic, top_k) for topic in slide_topics}

    def _fallback_retrieve(self, query: str, top_k: int) -> list[dict]:
        """Simple keyword overlap fallback."""
        query_words = set(re.findall(r"\w+", query.lower()))
        scored = []
        for i, chunk in enumerate(self.chunks):
            chunk_words = set(re.findall(r"\w+", chunk.lower()))
            overlap = len(query_words & chunk_words)
            if overlap > 0:
                scored.append({"chunk": chunk, "score": overlap / (len(query_words) + 1), "index": i})
        scored.sort(key=lambda x: x["score"], reverse=True)
        return scored[:top_k]

    def _chunk_text(self, text: str) -> list[str]:
        """Split text into overlapping chunks by paragraph boundaries."""
        paragraphs = re.split(r"\n\s*\n", text)
        paragraphs = [p.strip() for p in paragraphs if len(p.strip()) > 20]

        if not paragraphs:
            paragraphs = [text]

        chunks = []
        current = ""
        for para in paragraphs:
            if len(current) + len(para) < self.chunk_size:
                current += "\n\n" + para if current else para
            else:
                if current:
                    chunks.append(current)
                current = para
        if current:
            chunks.append(current)

        if len(chunks) <= 1:
            return chunks

        if self.overlap > 0 and len(chunks) > 1:
            overlapped = []
            for i, chunk in enumerate(chunks):
                if i > 0:
                    prev_end = chunks[i - 1][-self.overlap:] if len(chunks[i - 1]) > self.overlap else chunks[i - 1]
                    chunk = prev_end + "\n" + chunk
                overlapped.append(chunk)
            chunks = overlapped

        return chunks
