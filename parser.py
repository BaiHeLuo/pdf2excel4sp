from __future__ import annotations

import re
from pathlib import Path

import fitz
import pdfplumber

from .models import DocumentResult, PageResult, SectionResult, TableBlock, Block
from .classifier import detect_repeating_headers, classify_blocks

SECTION_PATTERN = re.compile(r"(?m)^\s*(\d{2,4}\.\d{2})\s*(.*)$")
BROKEN_NUMBER_PATTERN = re.compile(r"(\d{2,4})\s*\.\s*(\d{2})")
DOUBLE_SPACE_PATTERN = re.compile(r"[ \t]{2,}")


def parse_pdf(pdf_path: str | Path) -> DocumentResult:
    source_path = Path(pdf_path)
    pages: list[PageResult] = []
    sections: list[SectionResult] = []

    with fitz.open(source_path) as doc, pdfplumber.open(source_path) as plumber_doc:
        # first pass: extract raw text and raw table candidates
        page_texts: list[str] = []
        raw_table_candidates: dict[int, list[TableBlock]] = {}
        for index, page in enumerate(doc, start=1):
            raw_text = page.get_text("text") or ""
            plumber_page = plumber_doc.pages[index - 1]
            table_candidates = _extract_tables(plumber_page, index, as_candidates=True)
            raw_table_candidates[index] = table_candidates
            page_texts.append(raw_text)
            pages.append(
                PageResult(
                    page_number=index,
                    raw_text=raw_text.strip(),
                    normalized_text="",
                    tables=[],
                    blocks=[],
                )
            )

        # detect repeating headers/footers across pages
        headers = detect_repeating_headers(page_texts)

        # second pass: classify blocks and accept high-confidence tables
        for page in pages:
            raw_text = page.raw_text
            classified = classify_blocks(raw_text, headers)
            # convert to Block objects
            page.blocks = [Block(block_type=t, text=b, confidence=1.0) for t, b in classified]

            # keep normalized_text as concatenation of non-header non-table blocks
            text_blocks = [b.text for b in page.blocks if b.block_type in ("text", "title")]
            page.normalized_text = normalize_text("\n\n".join(text_blocks))

            # filter table candidates by confidence and attach to page.tables
            candidates = raw_table_candidates.get(page.page_number, [])
            accepted_tables: list[TableBlock] = []
            for tbl in candidates:
                conf = _assess_table_confidence(tbl)
                if conf >= 0.5:
                    accepted_tables.append(tbl)
                else:
                    # low-confidence table: keep as a low-confidence block for reference
                    page.blocks.append(Block(block_type="table_candidate", text=tbl.text, confidence=conf))
            page.tables = accepted_tables

    full_text, page_spans = build_full_text_and_spans(pages)
    sections = extract_sections(full_text, page_spans)
    return DocumentResult(
        source_file=str(source_path),
        pages=pages,
        sections=sections,
        chunks=[],
        metadata={
            "page_count": len(pages),
            "section_count": len(sections),
        },
    )


def build_full_text_and_spans(pages: list[PageResult]) -> tuple[str, list[tuple[int, int, int]]]:
    parts: list[str] = []
    spans: list[tuple[int, int, int]] = []
    cursor = 0
    for page in pages:
        if not page.normalized_text:
            continue
        if parts:
            parts.append("\n\n")
            cursor += 2
        start = cursor
        parts.append(page.normalized_text)
        cursor += len(page.normalized_text)
        end = cursor
        spans.append((page.page_number, start, end))
    return "".join(parts), spans


def normalize_text(text: str) -> str:
    lines = []
    previous_line = ""
    for line in text.splitlines():
        cleaned = DOUBLE_SPACE_PATTERN.sub(" ", line).strip()
        if not cleaned:
            if lines and lines[-1] != "":
                lines.append("")
            previous_line = ""
            continue

        merged = _merge_broken_number(previous_line, cleaned)
        if merged is not None:
            lines[-1] = merged
            previous_line = merged
            continue

        cleaned = BROKEN_NUMBER_PATTERN.sub(r"\1.\2", cleaned)
        lines.append(cleaned)
        previous_line = cleaned

    text = "\n".join(lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_sections(full_text: str, page_spans: list[tuple[int, int, int]]) -> list[SectionResult]:
    matches = list(SECTION_PATTERN.finditer(full_text))
    if not matches:
        return []

    sections: list[SectionResult] = []
    for index, match in enumerate(matches):
        section_id = match.group(1)
        title = match.group(2).strip() or section_id
        start = match.start()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(full_text)
        text = full_text[start:end].strip()
        page_refs = _page_refs_for_span(page_spans, start, end)
        sections.append(
            SectionResult(
                section_id=section_id,
                title=title,
                page_number=page_refs[0] if page_refs else 1,
                page_refs=page_refs,
                text=text,
            )
        )
    return sections


def _extract_tables(page: pdfplumber.page.Page, page_number: int, as_candidates: bool = False) -> list[TableBlock]:
    table_blocks: list[TableBlock] = []
    tables = page.extract_tables() or []
    for table_index, table in enumerate(tables, start=1):
        rows = []
        max_columns = max((len(row) for row in table if row), default=0)
        non_empty = 0
        total_cells = 0
        for row in table:
            cells = [normalize_text(cell) if cell else "" for cell in row]
            if len(cells) < max_columns:
                cells.extend([""] * (max_columns - len(cells)))
            rows.append(" | ".join(cell or "" for cell in cells).rstrip())
            for cell in cells:
                total_cells += 1
                if (cell or "").strip():
                    non_empty += 1
        table_text = f"表格 {table_index}（第 {page_number} 页）\n" + "\n".join(rows)
        table_blocks.append(
            TableBlock(
                page_number=page_number,
                table_index=table_index,
                text=table_text.strip(),
            )
        )
    return table_blocks


def _assess_table_confidence(table: TableBlock) -> float:
    # crude confidence: count lines and vertical separators
    lines = [ln for ln in table.text.splitlines() if ln.strip()]
    if not lines:
        return 0.0
    pipe_score = any('|' in ln for ln in lines)
    avg_cols = max((ln.count('|') for ln in lines), default=0)
    # non-empty cell ratio heuristic
    nonempty = sum(1 for ln in lines for cell in ln.split('|') if cell.strip())
    total = sum(len(ln.split('|')) for ln in lines)
    nonempty_ratio = nonempty / total if total else 0.0
    score = min(1.0, (0.4 * nonempty_ratio) + (0.4 * (1.0 if pipe_score else 0.0)) + (0.2 * min(1.0, avg_cols / 4)))
    return float(score)


def _merge_broken_number(previous_line: str, current_line: str) -> str | None:
    if not previous_line:
        return None
    if re.fullmatch(r"\d{2,4}\.", previous_line) and re.fullmatch(r"\d{2}.*", current_line):
        return previous_line + current_line
    if re.fullmatch(r"\d{2,4}\.", previous_line.rstrip()) and re.fullmatch(r"\d{2}.*", current_line):
        return previous_line.rstrip() + current_line
    if re.fullmatch(r"\d{2,4}\.\d", previous_line) and re.fullmatch(r"\d.*", current_line):
        return previous_line + current_line
    return None


def _page_refs_for_span(page_spans: list[tuple[int, int, int]], start: int, end: int) -> list[int]:
    refs = [page_number for page_number, page_start, page_end in page_spans if page_start < end and page_end > start]
    return refs or [1]
