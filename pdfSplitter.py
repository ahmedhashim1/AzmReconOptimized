#!/usr/bin/env python3
"""
PDF Page Splitter - Extract all, even, or odd pages from a PDF file
"""

import argparse
import sys
from pathlib import Path

try:
    from PyPDF2 import PdfReader, PdfWriter
except ImportError:
    print("Error: PyPDF2 is not installed.")
    print("Install it using: pip install PyPDF2")
    sys.exit(1)


def split_pdf(input_path, output_path, mode='all'):
    """
    Split PDF pages based on the specified mode.

    Args:
        input_path: Path to input PDF file
        output_path: Path for output PDF file
        mode: 'all', 'even', or 'odd'
    """
    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()
        total_pages = len(reader.pages)

        print(f"Total pages in PDF: {total_pages}")

        pages_to_extract = []

        if mode == 'all':
            pages_to_extract = list(range(total_pages))
        elif mode == 'even':
            # Even pages (2, 4, 6, ...) are at indices 1, 3, 5, ...
            pages_to_extract = list(range(1, total_pages, 2))
        elif mode == 'odd':
            # Odd pages (1, 3, 5, ...) are at indices 0, 2, 4, ...
            pages_to_extract = list(range(0, total_pages, 2))

        # Add selected pages to writer
        for page_num in pages_to_extract:
            writer.add_page(reader.pages[page_num])

        # Write output file
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)

        print(f"Successfully extracted {len(pages_to_extract)} pages ({mode} pages)")
        print(f"Output saved to: {output_path}")

    except FileNotFoundError:
        print(f"Error: File '{input_path}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        sys.exit(1)


def main():
    # ===== CONFIGURATION - Edit these values =====
    input_file = rf"C:\Users\ahmed\Downloads\70 Astagfar\70 Istaghfarat 18-80.pdf"  # Path to input PDF file
    output_folder = rf"C:\Users\ahmed\Downloads\70 Astagfar\Split"  # Output folder path
    split_mode = "odd"  # Options: 'all', 'even', 'odd'
    # ============================================

    # Validate input file
    input_path = Path(input_file)
    if not input_path.exists():
        print(f"Error: Input file '{input_file}' does not exist.")
        sys.exit(1)

    if not input_path.suffix.lower() == '.pdf':
        print("Error: Input file must be a PDF file.")
        sys.exit(1)

    # Validate split mode
    if split_mode not in ['all', 'even', 'odd']:
        print(f"Error: Invalid split mode '{split_mode}'. Use 'all', 'even', or 'odd'.")
        sys.exit(1)

    # Create output folder if it doesn't exist
    output_dir = Path(output_folder)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Determine output file path
    output_filename = f"{input_path.stem}_{split_mode}{input_path.suffix}"
    output_path = output_dir / output_filename

    # Perform splitting
    print(f"Input file: {input_path}")
    print(f"Mode: {split_mode}")
    print(f"Output folder: {output_dir}")
    print(f"Output file: {output_path}")
    print("-" * 50)

    split_pdf(str(input_path), str(output_path), split_mode)


if __name__ == '__main__':
    main()