#!/usr/bin/env python3
"""
Main entry point for PDF extraction and Excel conversion pipeline.
Supports single PDF or batch processing of all PDFs in a folder.
"""

import sys
from pathlib import Path
import argparse
import json

# Add src directory to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

# Import the extraction and Excel creation modules
from extract_pdf_tables import extract_structured_data
from create_excel_file import create_excel_from_json
from process_pdf import process_pdf


def main():
    parser = argparse.ArgumentParser(
        description="Extract pharmaceutical formulary data from PDF(s) and create Excel file(s)."
    )
    parser.add_argument(
        "pdf_path",
        nargs="?",
        help="Path to a single PDF file to extract",
    )
    parser.add_argument(
        "--pdf-dir",
        help="Path to a directory containing PDF files to process",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default="output",
        help="Output directory for extracted data (default: output)",
    )
    parser.add_argument(
        "--json-only",
        action="store_true",
        help="Only extract JSON, skip Excel creation",
    )
    parser.add_argument(
        "--excel-only",
        action="store_true",
        help="Only create Excel from existing JSON (requires --json-path)",
    )
    parser.add_argument(
        "--json-path",
        help="Path to existing JSON file (for --excel-only mode)",
    )

    args = parser.parse_args()

    # Validate input
    if args.excel_only:
        if not args.json_path:
            print("Error: --excel-only requires --json-path")
            return 1
        json_path = Path(args.json_path)
        if not json_path.exists():
            print(f"Error: JSON file not found: {json_path}")
            return 1

        # Excel-only mode
        print("=" * 80)
        print("EXCEL CREATION MODE")
        print("=" * 80)

        excel_path = json_path.parent / f"{json_path.parent.name}.xlsx"

        print(f"Input JSON: {json_path}")
        print(f"Output Excel: {excel_path}")

        create_excel_from_json(json_path, excel_path)
        print("\nExcel creation complete!\n")
        return 0

    # Batch mode: process all PDFs in a folder
    if args.pdf_dir:
        pdf_dir = Path(args.pdf_dir)
        if not pdf_dir.exists() or not pdf_dir.is_dir():
            print(f"Error: Directory not found: {pdf_dir}")
            return 1

        pdf_files = list(pdf_dir.glob("*.pdf"))
        if not pdf_files:
            print(f"No PDF files found in {pdf_dir}")
            return 1

        print("=" * 80)
        print(f"Processing all PDFs in: {pdf_dir}")
        print("=" * 80)

        for pdf_path in pdf_files:
            print(f"\nProcessing {pdf_path.name}...")
            process_pdf(pdf_path, args.output_dir, args.json_only)

        print("\nBatch processing complete!\n")
        return 0

    # Single PDF mode
    if not args.pdf_path:
        print("Error: pdf_path is required (unless using --pdf-dir or --excel-only)")
        return 1

    pdf_path = Path(args.pdf_path)
    if not pdf_path.exists():
        print(f"Error: PDF file not found: {pdf_path}")
        return 1

    print(f"\nProcessing {pdf_path.name}...")
    process_pdf(pdf_path, args.output_dir, args.json_only)
    return 0


if __name__ == "__main__":
    sys.exit(main())
