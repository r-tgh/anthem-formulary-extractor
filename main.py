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


def process_pdf(pdf_path: Path, output_dir: str, json_only: bool):
    """
    Process a single PDF: extract data, save JSON, and optionally create Excel.
    """
    print("=" * 80)
    print(f"STEP 1: PDF EXTRACTION for {pdf_path.name}")
    print("=" * 80)

    # Create output directory: output/PDFFILENAME/
    pdf_filename = pdf_path.stem
    pdf_output_dir = Path(output_dir) / pdf_filename
    pdf_output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Input PDF: {pdf_path}")
    print(f"Output directory: {pdf_output_dir}")
    print("=" * 80)

    # Extract data from PDF
    data = extract_structured_data(str(pdf_path))

    # Save JSON files
    json_output_path = pdf_output_dir / "extracted_data.json"
    warnings_path = pdf_output_dir / "extraction_warnings.json"
    toc_path = pdf_output_dir / "table_of_contents.json"

    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(data["categories"], f, indent=2, ensure_ascii=False)

    with open(warnings_path, "w", encoding="utf-8") as f:
        json.dump(data["warnings"], f, indent=2, ensure_ascii=False)

    with open(toc_path, "w", encoding="utf-8") as f:
        json.dump(data["table_of_contents"], f, indent=2, ensure_ascii=False)

    print("\n" + "=" * 80)
    print(f"Extraction complete!")
    print(f"  - Categories saved to: {json_output_path}")
    print(f"  - Warnings saved to: {warnings_path}")
    print(f"  - Table of Contents saved to: {toc_path}")

    print(f"\nSummary:")
    print(f"  - TOC entries found: {len(data['table_of_contents'])}")
    print(f"  - Categories found: {len(data['categories'])}")
    print(f"  - Warnings (skipped rows): {len(data['warnings'])}")

    # Count total subcategories and rows
    total_subcategories = sum(len(cat["subCategories"]) for cat in data["categories"])
    total_rows = sum(
        len(subcat["rows"])
        for cat in data["categories"]
        for subcat in cat["subCategories"]
    )

    print(f"  - Total subcategories: {total_subcategories}")
    print(f"  - Total rows extracted: {total_rows}")

    # Stop here if --json-only
    if json_only:
        print("\n" + "=" * 80)
        print("JSON-only mode: Skipping Excel creation")
        print("=" * 80)
        return

    # Excel Creation
    print("\n" + "=" * 80)
    print("STEP 2: EXCEL CREATION")
    print("=" * 80)

    excel_output_path = pdf_output_dir / f"{pdf_filename}.xlsx"

    try:
        create_excel_from_json(json_output_path, excel_output_path)

        print("\n" + "=" * 80)
        print("PIPELINE COMPLETE!")
        print("=" * 80)
        print(f"\nOutput files:")
        print(f"  - JSON: {json_output_path}")
        print(f"  - Excel: {excel_output_path}")
        print(f"  - Warnings: {warnings_path}")
        print(f"  - TOC: {toc_path}")

    except ImportError as e:
        print(f"\nWarning: Could not create Excel file")
        print(f"  Error: {e}")
        print(f"  Please install openpyxl: pip install openpyxl")
        print(f"\n  JSON files have been created successfully in: {pdf_output_dir}")
    except Exception as e:
        print(f"\nError creating Excel file: {e}")
        print(f"  JSON files have been created successfully in: {pdf_output_dir}")


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
            print(f"\n--- Processing {pdf_path.name} ---")
            process_pdf(pdf_path, args.output_dir, args.json_only)

        print("\nBatch processing complete!")
        return 0

    # Single PDF mode
    if not args.pdf_path:
        print("Error: pdf_path is required (unless using --pdf-dir or --excel-only)")
        return 1

    pdf_path = Path(args.pdf_path)
    if not pdf_path.exists():
        print(f"Error: PDF file not found: {pdf_path}")
        return 1

    process_pdf(pdf_path, args.output_dir, args.json_only)
    return 0


if __name__ == "__main__":
    sys.exit(main())