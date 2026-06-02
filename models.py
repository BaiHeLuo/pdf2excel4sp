from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class TableBlock:
    page_number: int
    table_index: int
    text: str


@dataclass
class Block:
    block_type: str
    text: str
    confidence: float = 1.0


@dataclass
class PageResult:
    page_number: int
    raw_text: str
    normalized_text: str
    tables: list[TableBlock] = field(default_factory=list)
    blocks: list[Block] = field(default_factory=list)
    used_ocr: bool = False


@dataclass
class SectionResult:
    section_id: str
    title: str
    page_number: int
    text: str
    page_refs: list[int] = field(default_factory=list)


@dataclass
class ChunkResult:
    chunk_id: str
    section_id: str
    page_start: int
    page_end: int
    context_hint: str
    text: str


@dataclass
class DocumentResult:
    source_file: str
    pages: list[PageResult]
    sections: list[SectionResult]
    chunks: list[ChunkResult]
    metadata: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "source_file": self.source_file,
            "metadata": self.metadata,
            "pages": [
                {
                    "page_number": page.page_number,
                    "raw_text": page.raw_text,
                    "normalized_text": page.normalized_text,
                    "used_ocr": page.used_ocr,
                    "tables": [
                        {
                            "page_number": table.page_number,
                            "table_index": table.table_index,
                            "text": table.text,
                        }
                        for table in page.tables
                    ],
                }
                for page in self.pages
            ],
            "sections": [
                {
                    "section_id": section.section_id,
                    "title": section.title,
                    "page_number": section.page_number,
                    "page_refs": section.page_refs,
                    "text": section.text,
                }
                for section in self.sections
            ],
            "chunks": [
                {
                    "chunk_id": chunk.chunk_id,
                    "section_id": chunk.section_id,
                    "page_start": chunk.page_start,
                    "page_end": chunk.page_end,
                    "context_hint": chunk.context_hint,
                    "text": chunk.text,
                }
                for chunk in self.chunks
            ],
        }
