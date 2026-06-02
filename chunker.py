from __future__ import annotations

import re
from dataclasses import dataclass

from .models import ChunkResult, DocumentResult, SectionResult

SECTION_PATTERN = re.compile(r"(?m)^\s*(\d{2,4}\.\d{2})\s*(.*)$")


@dataclass
class ChunkSettings:
    chunk_size: int = 3000
    overlap: int = 250


def chunk_document(document: DocumentResult, chunk_size: int = 3000, overlap: int = 250) -> DocumentResult:
    settings = ChunkSettings(chunk_size=chunk_size, overlap=overlap)
    chunks: list[ChunkResult] = []
    chunk_index = 1

    for section in document.sections:
        section_chunks = _chunk_section(section, settings)
        for section_chunk in section_chunks:
            page_start = section.page_refs[0] if section.page_refs else section.page_number
            page_end = section.page_refs[-1] if section.page_refs else section.page_number
            chunks.append(
                ChunkResult(
                    chunk_id=f"chunk_{chunk_index:03d}",
                    section_id=section.section_id,
                    page_start=page_start,
                    page_end=page_end,
                    context_hint=f"章节 {section.section_id}：{section.title}",
                    text=section_chunk,
                )
            )
            chunk_index += 1

    if not chunks:
        combined_text = "\n\n".join(page.normalized_text for page in document.pages if page.normalized_text)
        for section_chunk in _split_text_by_size(combined_text, settings):
            chunks.append(
                ChunkResult(
                    chunk_id=f"chunk_{chunk_index:03d}",
                    section_id="document",
                    page_start=1,
                    page_end=max(len(document.pages), 1),
                    context_hint="全文分块",
                    text=section_chunk,
                )
            )
            chunk_index += 1

    document.chunks = chunks
    document.metadata["chunk_count"] = len(chunks)
    document.metadata["chunk_size"] = settings.chunk_size
    document.metadata["chunk_overlap"] = settings.overlap
    return document


def _chunk_section(section: SectionResult, settings: ChunkSettings) -> list[str]:
    text = section.text.strip()
    if not text:
        return []
    if len(text) <= settings.chunk_size:
        return [text]
    return _split_text_by_size(text, settings)


def _split_text_by_size(text: str, settings: ChunkSettings) -> list[str]:
    paragraphs = [paragraph.strip() for paragraph in re.split(r"\n{2,}", text) if paragraph.strip()]
    if not paragraphs:
        return []

    chunks: list[str] = []
    current = ""
    for paragraph in paragraphs:
        candidate = f"{current}\n\n{paragraph}" if current else paragraph
        if len(candidate) <= settings.chunk_size:
            current = candidate
            continue

        if current:
            chunks.append(current.strip())
        if len(paragraph) <= settings.chunk_size:
            current = paragraph
            continue

        chunks.extend(_split_long_paragraph(paragraph, settings))
        current = ""

    if current:
        chunks.append(current.strip())
    return chunks


def _split_long_paragraph(paragraph: str, settings: ChunkSettings) -> list[str]:
    parts: list[str] = []
    start = 0
    length = len(paragraph)
    while start < length:
        end = min(start + settings.chunk_size, length)
        if end < length:
            breakpoint = max(paragraph.rfind("。", start, end), paragraph.rfind("\n", start, end), paragraph.rfind("。", start, end))
            if breakpoint > start:
                end = breakpoint + 1
        part = paragraph[start:end].strip()
        if part:
            parts.append(part)
        if end >= length:
            break
        start = max(end - settings.overlap, end)
    return parts
