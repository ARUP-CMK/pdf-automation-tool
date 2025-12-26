"""
PDF utility functions for preview and processing
"""

from typing import Tuple, Optional
from PIL import Image
import fitz  # PyMuPDF


def get_page_count(file_path: str) -> int:
    """
    Get the total number of pages in a PDF file.
    
    Args:
        file_path: Path to the PDF file
        
    Returns:
        Number of pages in the PDF, or 0 if error
    """
    try:
        doc = fitz.open(file_path)
        page_count = len(doc)
        doc.close()
        return page_count
    except Exception as e:
        print(f"Error getting page count: {e}")
        return 0


def generate_preview_image(
    file_path: str, 
    max_size: Tuple[int, int], 
    page_number: int = 0,
    border_size: int = 4
) -> Optional[Image.Image]:
    """
    Generate a preview image from a specific page of a PDF file.
    
    Args:
        file_path: Path to the PDF file
        max_size: Tuple of (max_width, max_height) for the preview area
        page_number: Which page to render (0-indexed, default: 0)
        border_size: Size of the border in pixels (default: 4)
        
    Returns:
        PIL Image object resized to fit within max_size while maintaining aspect ratio,
        with a light gray border around it, or None if an error occurs
    """
    try:
        # Open the PDF document
        doc = fitz.open(file_path)
        
        if len(doc) == 0:
            doc.close()
            raise ValueError("PDF file is empty or has no pages")
        
        # Validate page number
        if page_number < 0 or page_number >= len(doc):
            doc.close()
            raise ValueError(f"Page {page_number} does not exist. PDF has {len(doc)} pages.")
        
        # Get the specified page
        page = doc[page_number]
        
        # Calculate zoom factor to render at good quality
        # Higher zoom = better quality but larger image
        zoom = 2.0  # 2x zoom for good quality
        mat = fitz.Matrix(zoom, zoom)
        
        # Render page to a pixmap (image)
        pix = page.get_pixmap(matrix=mat)
        
        # Convert pixmap to PIL Image
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # Close the document
        doc.close()
        
        # Calculate the original aspect ratio
        original_width, original_height = img.size
        original_aspect = original_width / original_height
        
        # Calculate the maximum available size (accounting for border)
        # Subtract border from both dimensions (left+right, top+bottom)
        max_width = max_size[0] - (border_size * 2)
        max_height = max_size[1] - (border_size * 2)
        
        # Calculate target size while preserving aspect ratio (contain, not cover)
        max_aspect = max_width / max_height
        
        if original_aspect > max_aspect:
            # Image is wider - fit to width
            target_width = max_width
            target_height = int(max_width / original_aspect)
        else:
            # Image is taller - fit to height
            target_height = max_height
            target_width = int(max_height * original_aspect)
        
        # Resize image using high-quality resampling while maintaining aspect ratio
        resized_img = img.resize((target_width, target_height), Image.Resampling.LANCZOS)
        
        # Create a new image with border: light gray background
        border_color = (220, 220, 220)  # Light gray RGB
        bordered_img = Image.new("RGB", (target_width + border_size * 2, target_height + border_size * 2), border_color)
        
        # Paste the resized image onto the bordered image, centered
        bordered_img.paste(resized_img, (border_size, border_size))
        
        return bordered_img
        
    except Exception as e:
        print(f"Error generating preview image: {e}")
        return None


def get_page_image(
    file_path: str,
    page_number: int,
    max_size: Tuple[int, int]
) -> Optional[Image.Image]:
    """
    Convenience function to get a specific page as an image.
    This is an alias for generate_preview_image with page_number.
    
    Args:
        file_path: Path to the PDF file
        page_number: Which page to render (0-indexed)
        max_size: Tuple of (max_width, max_height) for the preview area
        
    Returns:
        PIL Image object or None if error
    """
    return generate_preview_image(file_path, max_size, page_number=page_number)
