# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "docling>=2.0.0",
#     "pywin32>=306",
# ]
# ///

"""
Convert Microsoft Publisher (.pub) files to high-fidelity PDF and clean Markdown for LLMs.
Uses local Office 365 Publisher for rendering and Docling for document layout/text extraction.
"""

import argparse
import os
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client
from docling.document_converter import DocumentConverter


def pub_to_high_fidelity_pdf(pub_path: Path, pdf_path: Path) -> None:
    """
    Converts a .pub file to a PDF using local Microsoft 365 Publisher via COM automation.
    Uses commercial print quality (pbPrintQualityCommercial) to ensure all vector shapes,
    embedded fonts, and high-resolution assets are preserved.
    """
    print(f"[*] Opening Microsoft 365 Publisher...")
    pythoncom.CoInitialize()

    # Publisher COM Constants
    pbFixedFormatTypePDF = 2
    pbPrintQualityCommercial = 2  # Max resolution & embedded fonts

    app = None
    try:
        # Launch Office 365 Publisher in the background
        app = win32com.client.Dispatch("Publisher.Application")

        abs_pub = str(pub_path.resolve())
        abs_pdf = str(pdf_path.resolve())

        print(f"[*] Loading document: {abs_pub}")
        # Open(FileName, ReadOnly=True, OpenAndRepair=False)
        doc = app.Open(abs_pub, True, False)

        print(f"[*] Exporting high-fidelity PDF to: {abs_pdf}")
        # ExportAsFixedFormat(Format, FileName, Intent, IncludeDocProps)
        doc.ExportAsFixedFormat(
            pbFixedFormatTypePDF,
            abs_pdf,
            pbPrintQualityCommercial,
            True,
        )
        doc.Close()
        print(f"[✓] PDF successfully generated.")

    except Exception as e:
        print(f"[!] Error during Publisher export: {e}", file=sys.stderr)
        raise
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
            del app
        pythoncom.CoUninitialize()


def pdf_to_markdown_docling(pdf_path: Path, md_path: Path) -> str:
    """
    Parses layout, tables, reading order, and hierarchy of the PDF into structured Markdown.
    """
    print(f"[*] Running Docling layout analysis and OCR extraction on PDF...")
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
        description="Convert Microsoft Publisher (.pub) files to high-fidelity PDF and Markdown for LLMs."
    )
    parser.add_argument("pub_file", type=Path, help="Path to the input .pub file.")
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        default=None,
        help="Directory to save output files (defaults to input file's directory).",
    )
    parser.add_argument(
        "--print-md",
        action="store_true",
        help="Print the generated Markdown to stdout.",
    )

    args = parser.parse_args()

    input_file = args.pub_file.resolve()
    if not input_file.exists() or input_file.suffix.lower() != ".pub":
        print(
            f"Error: '{input_file}' does not exist or is not a .pub file.",
            file=sys.stderr,
        )
        sys.exit(1)

    out_dir = args.output_dir.resolve() if args.output_dir else input_file.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    base_name = input_file.stem
    pdf_out = out_dir / f"{base_name}.pdf"
    md_out = out_dir / f"{base_name}.md"

    try:
        # Step 1: Office 365 Publisher -> High-Fidelity PDF
        pub_to_high_fidelity_pdf(input_file, pdf_out)

        # Step 2: PDF -> Structured LLM Markdown via Docling
        md_text = pdf_to_markdown_docling(pdf_out, md_out)

        if args.print_md:
            print("\n" + "=" * 40 + " MARKDOWN OUTPUT " + "=" * 40)
            print(md_text)

    except Exception as e:
        print(f"\nPipeline failed: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
