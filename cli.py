from __future__ import annotations

import argparse
import json
from pathlib import Path

from .chunker import chunk_document
from .parser import parse_pdf


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="PDF intelligent parser MVP")
    parser.add_argument("input", help="Path to the source PDF")
    parser.add_argument("--output", help="Path to the JSON output file")
    parser.add_argument("--chunk-size", type=int, default=3000, help="Maximum characters per chunk")
    parser.add_argument("--overlap", type=int, default=250, help="Chunk overlap size")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    document = parse_pdf(args.input)
    document = chunk_document(document, chunk_size=args.chunk_size, overlap=args.overlap)
    payload = document.to_dict()

    output_text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(output_text, encoding="utf-8")
    else:
        print(output_text)


if __name__ == "__main__":
    main()
