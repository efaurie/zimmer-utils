# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "pywin32>=306",
# ]
# ///

"""
Convert Microsoft Publisher (.pub) files to high-fidelity PDF.
Uses local Office 365 Publisher via COM automation for rendering.
"""

import argparse
import os
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client


def pub_to_high_fidelity_pdf(pub_path: Path, pdf_path: Path) -> None:
    """
    Converts a .pub file to a PDF using local Microsoft 365 Publisher via COM automation.
    Uses commercial print quality (pbPrintQualityCommercial) to ensure all vector shapes,
    embedded fonts, and high-resolution assets are preserved.
    """
    print("[*] Opening Microsoft 365 Publisher...")
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


def main():
    parser = argparse.ArgumentParser(
        description="Convert Microsoft Publisher (.pub) files to high-fidelity PDF."
    )
    parser.add_argument("pub_file", type=Path, help="Path to the input .pub file.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output PDF file path or directory (defaults to input file's directory and stem name).",
    )
    parser.add_argument(
        "-d",
        "--output-dir",
        type=Path,
        default=None,
        help="Directory to save the output PDF file.",
    )

    args = parser.parse_args()

    input_file = args.pub_file.resolve()
    if not input_file.exists() or input_file.suffix.lower() != ".pub":
        print(
            f"Error: '{input_file}' does not exist or is not a .pub file.",
            file=sys.stderr,
        )
        sys.exit(1)

    # Determine destination PDF path
    if args.output_dir:
        out_dir = args.output_dir.resolve()
        pdf_out = out_dir / f"{input_file.stem}.pdf"
    elif args.output:
        out_path = args.output.resolve()
        if out_path.suffix.lower() == ".pdf":
            pdf_out = out_path
        else:
            pdf_out = out_path / f"{input_file.stem}.pdf"
    else:
        pdf_out = input_file.parent / f"{input_file.stem}.pdf"

    pdf_out.parent.mkdir(parents=True, exist_ok=True)

    try:
        start_time = time.time()
        pub_to_high_fidelity_pdf(input_file, pdf_out)
        elapsed = time.time() - start_time
        print(f"[✓] Completed in {elapsed:.2f}s -> {pdf_out}")
    except Exception as e:
        print(f"\nConversion failed: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
