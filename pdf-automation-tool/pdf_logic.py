"""
PDF Logic Module for Title Block Automation
Handles resizing customer drawings and overlaying title block templates
"""

import fitz  # PyMuPDF
from pathlib import Path
from typing import Optional, Set


# A3 Landscape dimensions in points (1 inch = 72 pts, 1 mm = 2.834645 pts)
A3_WIDTH_PTS = 1191  # 420mm
A3_HEIGHT_PTS = 842  # 297mm

# Margin definitions (converting mm to points)
MM_TO_PTS = 2.834645
TOP_MARGIN_MM = 40
BOTTOM_MARGIN_MM = 30
LEFT_MARGIN_MM = 20
RIGHT_MARGIN_MM = 20

# Calculate margins in points
TOP_MARGIN_PTS = TOP_MARGIN_MM * MM_TO_PTS      # ~113 pts
BOTTOM_MARGIN_PTS = BOTTOM_MARGIN_MM * MM_TO_PTS  # ~85 pts
LEFT_MARGIN_PTS = LEFT_MARGIN_MM * MM_TO_PTS     # ~57 pts
RIGHT_MARGIN_PTS = RIGHT_MARGIN_MM * MM_TO_PTS   # ~57 pts


def get_safe_zone_rect() -> fitz.Rect:
    """
    Calculate the Safe Zone rectangle for placing customer drawings.
    
    The Safe Zone is the area within the A3 page where the customer
    drawing will be placed, accounting for margins around the title block.
    
    Returns:
        fitz.Rect: The safe zone rectangle in points
    """
    x0 = LEFT_MARGIN_PTS
    y0 = TOP_MARGIN_PTS
    x1 = A3_WIDTH_PTS - RIGHT_MARGIN_PTS
    y1 = A3_HEIGHT_PTS - BOTTOM_MARGIN_PTS
    
    return fitz.Rect(x0, y0, x1, y1)


def process_with_margins(
    input_path: str,
    template_path: str,
    output_path: str,
    project_data: Optional[dict] = None,
    excluded_pages: Optional[Set[int]] = None
) -> None:
    """
    Process a customer PDF by resizing it to fit within A3 safe zone
    and overlaying a title block template.
    
    Args:
        input_path: Path to the input customer PDF file
        template_path: Path to the title block template PDF
        output_path: Path where the processed PDF will be saved
        project_data: Optional dictionary containing project metadata
                      (project_name, client_name, date, drawn_by)
        excluded_pages: Optional set of 0-based page indices to exclude
                        from processing (e.g., {0, 2, 3} skips pages 1, 3, 4)
    
    Raises:
        FileNotFoundError: If input or template file doesn't exist
        Exception: If PDF processing fails
    """
    # Initialize excluded_pages if None
    if excluded_pages is None:
        excluded_pages = set()
    
    # Validate input files exist
    if not Path(input_path).exists():
        raise FileNotFoundError(f"Input PDF not found: {input_path}")
    
    if not Path(template_path).exists():
        raise FileNotFoundError(f"Template PDF not found: {template_path}")
    
    # Open source documents
    source_doc = fitz.open(input_path)
    template_doc = fitz.open(template_path)
    
    # Create output document
    output_doc = fitz.open()
    
    # Define the full A3 landscape page rect
    a3_rect = fitz.Rect(0, 0, A3_WIDTH_PTS, A3_HEIGHT_PTS)
    
    # Get the safe zone for customer drawings
    safe_zone = get_safe_zone_rect()
    
    # Calculate pages to process
    total_pages = source_doc.page_count
    pages_to_process = [p for p in range(total_pages) if p not in excluded_pages]
    
    print(f"Processing {len(pages_to_process)}/{total_pages} page(s) from: {input_path}")
    if excluded_pages:
        excluded_display = [p + 1 for p in sorted(excluded_pages) if p < total_pages]
        if excluded_display:
            print(f"  Excluding pages: {excluded_display}")
    print(f"A3 Dimensions: {A3_WIDTH_PTS} x {A3_HEIGHT_PTS} pts")
    print(f"Safe Zone: {safe_zone}")
    
    try:
        # Process each page of the input PDF (except excluded ones)
        processed_count = 0
        for page_num in range(source_doc.page_count):
            # Skip excluded pages
            if page_num in excluded_pages:
                print(f"  Skipping page {page_num + 1} (excluded)")
                continue
            
            processed_count += 1
            
            # Get the source page
            source_page = source_doc[page_num]
            
            # Clean/flatten the source page contents
            source_page.clean_contents()
            
            # Create a new blank A3 landscape page in output
            output_page = output_doc.new_page(
                width=A3_WIDTH_PTS,
                height=A3_HEIGHT_PTS
            )
            
            # Place the customer drawing within the Safe Zone
            # keep_proportion=True ensures aspect ratio is maintained
            output_page.show_pdf_page(
                safe_zone,                   # Target rectangle (safe zone)
                source_doc,                  # Source document
                page_num,                    # Page number from source
                keep_proportion=True         # Maintain aspect ratio
            )
            
            # Overlay the title block template on top (full A3 rect)
            # The template should have transparency where the drawing shows through
            if template_doc.page_count > 0:
                output_page.show_pdf_page(
                    a3_rect,                 # Full page rect
                    template_doc,            # Template document
                    0,                       # First page of template
                    keep_proportion=True     # Maintain aspect ratio
                )
            
            print(f"  Processed page {page_num + 1}/{source_doc.page_count} -> output page {processed_count}")
        
        # Check if any pages were processed
        if processed_count == 0:
            print(f"  Warning: All pages were excluded. No output generated.")
            return
        
        # Ensure output directory exists
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Save the output document
        output_doc.save(output_path)
        print(f"Saved output to: {output_path} ({processed_count} pages)")
        
    finally:
        # Close all documents
        source_doc.close()
        template_doc.close()
        output_doc.close()


def process_single_page(
    input_path: str,
    template_path: str,
    output_path: str,
    page_number: int = 0
) -> None:
    """
    Process a single page from a customer PDF.
    
    Args:
        input_path: Path to the input customer PDF file
        template_path: Path to the title block template PDF
        output_path: Path where the processed PDF will be saved
        page_number: Which page to process (0-indexed)
    """
    source_doc = fitz.open(input_path)
    
    if page_number >= source_doc.page_count:
        source_doc.close()
        raise ValueError(f"Page {page_number} doesn't exist. Document has {source_doc.page_count} pages.")
    
    template_doc = fitz.open(template_path)
    output_doc = fitz.open()
    
    a3_rect = fitz.Rect(0, 0, A3_WIDTH_PTS, A3_HEIGHT_PTS)
    safe_zone = get_safe_zone_rect()
    
    try:
        source_page = source_doc[page_number]
        source_page.clean_contents()
        
        output_page = output_doc.new_page(
            width=A3_WIDTH_PTS,
            height=A3_HEIGHT_PTS
        )
        
        output_page.show_pdf_page(
            safe_zone,
            source_doc,
            page_number,
            keep_proportion=True
        )
        
        if template_doc.page_count > 0:
            output_page.show_pdf_page(
                a3_rect,
                template_doc,
                0,
                keep_proportion=True
            )
        
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        output_doc.save(output_path)
        
    finally:
        source_doc.close()
        template_doc.close()
        output_doc.close()


if __name__ == "__main__":
    # Test/example usage
    import sys
    
    if len(sys.argv) >= 4:
        input_file = sys.argv[1]
        template_file = sys.argv[2]
        output_file = sys.argv[3]
        
        process_with_margins(input_file, template_file, output_file)
    else:
        print("Usage: python pdf_logic.py <input.pdf> <template.pdf> <output.pdf>")
        print("\nThis module processes customer PDFs by:")
        print("  1. Resizing them to fit within an A3 safe zone")
        print("  2. Overlaying a title block template")
        print(f"\nGeometry:")
        print(f"  A3 Landscape: {A3_WIDTH_PTS} x {A3_HEIGHT_PTS} pts")
        print(f"  Safe Zone: {get_safe_zone_rect()}")
        print(f"  Top Margin: {TOP_MARGIN_MM}mm ({TOP_MARGIN_PTS:.1f} pts)")
        print(f"  Bottom Margin: {BOTTOM_MARGIN_MM}mm ({BOTTOM_MARGIN_PTS:.1f} pts)")
        print(f"  Left/Right Margins: {LEFT_MARGIN_MM}mm ({LEFT_MARGIN_PTS:.1f} pts)")

