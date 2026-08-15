# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "docling>=2.0.0",
# ]
# ///

"""
Convert PDF files to structured Markdown using Docling.
Extracts document layout, tables, reading order, and hierarchy for LLM consumption.
"""

import argparse
import os
import sys
import time
from pathlib import Path

from docling.document_converter import DocumentConverter


def pdf_to_markdown_docling(pdf_path: Path, md_path: Path) -> str:
    """
    Parses layout, tables, reading order, and hierarchy of the PDF into structured Markdown.
    """
    print(f"[*] Running Docling layout analysis and OCR extraction on: {pdf_path.resolve()}")
    start_time = time.time()

    converter = DocumentConverter()
    result = converter.convert(str(pdf_path.resolve()))

    # Export clean markdown representation
    markdown_content = result.document.export_to_markdown()

    md_path.write_text(markdown_content, encoding="utf-8")
    elapsed = time.time() - start_time
    print(f"[✓] Markdown saved to: {md_path.resolve()} ({elapsed:.2f}s)")

    return markdown_content


def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF files to structured Markdown using Docling."
    )
    parser.add_argument(
        "pdf_file",
        type=Path,
        help="Path to the input .pdf file."
    )
    parser.add_argument(
        "-o", "--output",
        type=Path,
        default=None,
        help="Output Markdown file path or directory (defaults to input file's directory and stem name)."
    )
    parser.add_argument(
        "-d", "--output-dir",
        type=Path,
        default=None,
        help="Directory to save the output Markdown file."
    )
    parser.add_argument(
        "--print-md",
        action="store_true",
        help="Print the generated Markdown to stdout."
    )

    args = parser.parse_args()

    input_file = args.pdf_file.resolve()
    if not input_file.exists() or input_file.suffix.lower() != ".pdf":
        print(f"Error: '{input_file}' does not exist or is not a .pdf file.", file=sys.stderr)
        sys.exit(1)

    # Determine destination Markdown path
    if args.output_dir:
        out_dir = args.output_dir.resolve()
        md_out = out_dir / f"{input_file.stem}.md"
    elif args.output:
        out_path = args.output.resolve()
        if out_path.suffix.lower() == ".md":
            md_out = out_path
        else:
            md_out = out_path / f"{input_file.stem}.md"
    else:
        md_out = input_file.parent / f"{input_file.stem}.md"

    md_out.parent.mkdir(parents=True, exist_ok=True)

    try:
        md_text = pdf_to_markdown_docling(input_file, md_out)

        if args.print_md:
            print("\n" + "=" * 40 + " MARKDOWN OUTPUT " + "=" * 40)
            print(md_text)

    except Exception as e:
        print(f"\nConversion failed: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
