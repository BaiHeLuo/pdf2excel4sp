from __future__ import annotations

import re
from collections import Counter
from typing import Iterable, List, Set


def detect_repeating_headers(pages_texts: Iterable[str], top_n: int = 3, min_repeat: int = 3) -> Set[str]:
    """
    Detect repeated header/footer lines across pages.

    Return a set of line texts that appear in the top `top_n` lines of at least `min_repeat` pages.
    """
    counter: Counter[str] = Counter()
    for text in pages_texts:
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        for ln in lines[:top_n]:
            counter[ln] += 1

    return {ln for ln, c in counter.items() if c >= min_repeat}


def split_into_blocks(text: str) -> list[str]:
    # Split by two or more newlines to get paragraph-like blocks
    blocks = [b.strip() for b in re.split(r"\n{2,}", text) if b.strip()]
    return blocks


def looks_like_table_block(block: str) -> bool:
    # Heuristics: contains multiple '|' or multiple consistent columns separated by 2+ spaces
    if '|' in block and block.count('\n') >= 1:
        return True
    # rows with multiple whitespace-separated columns
    lines = [ln for ln in block.splitlines() if ln.strip()]
    if len(lines) < 2:
        return False
    col_counts = [len(re.split(r"\s{2,}|\t", ln)) for ln in lines]
    if max(col_counts, default=0) >= 2 and len(set(col_counts)) <= max(2, len(lines)//2):
        return True
    return False


def is_header_like(block: str, headers_set: Set[str]) -> bool:
    # If the block's first non-empty line is in headers_set or is very short, treat as header/footer
    first_line = next((ln.strip() for ln in block.splitlines() if ln.strip()), '')
    if not first_line:
        return True
    if first_line in headers_set:
        return True
    # very short lines (<=6 chars) likely to be header/footer or page number
    if len(first_line) <= 6 and any(ch.isdigit() for ch in first_line):
        return True
    return False


def classify_blocks(text: str, headers_set: Set[str]) -> list[tuple[str, str]]:
    """Return list of (block_type, block_text). block_type in {'header','footer','table_candidate','text','title'}"""
    blocks = split_into_blocks(text)
    classified: List[tuple[str, str]] = []
    for b in blocks:
        if is_header_like(b, headers_set):
            classified.append(("header", b))
            continue
        if looks_like_table_block(b):
            classified.append(("table_candidate", b))
            continue
        # Heuristic: if block starts with a numbering like 100.01 treat as title/section head
        if re.match(r"^\s*\d{2,4}\.\d{1,2}\b", b):
            classified.append(("title", b))
            continue
        classified.append(("text", b))
    return classified
