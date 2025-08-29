"""
Utility functions for coordinate system conversions and PDF operations.
"""

import fitz  # PyMuPDF
from typing import Tuple
from .models import PageDimensions, PDF_POINTS_PER_INCH


def get_pdf_dimensions(pdf_path_or_bytes) -> Tuple[PageDimensions, int]:
    """
    Extract page dimensions and page count from a PDF.
    
    Args:
        pdf_path_or_bytes: PDF file path or bytes
        
    Returns:
        Tuple of (first_page_dimensions, total_pages)
        
    Raises:
        ValueError: If PDF is empty or invalid
    """
    try:
        if isinstance(pdf_path_or_bytes, bytes):
            doc = fitz.open(stream=pdf_path_or_bytes, filetype="pdf")
        else:
            doc = fitz.open(pdf_path_or_bytes)
        
        if len(doc) == 0:
            raise ValueError("Empty PDF")
        
        # Get dimensions from first page
        first_page = doc[0]
        rect = first_page.rect
        dimensions = (rect.width, rect.height)
        page_count = len(doc)
        
        doc.close()
        return dimensions, page_count
    
    except Exception as e:
        raise ValueError(f"Failed to process PDF: {str(e)}")


def pixels_to_pdf_points(pixel_x: float, pixel_y: float, dpi: int, 
                        page_width_pts: float, page_height_pts: float) -> Tuple[float, float]:
    """
    Convert pixel coordinates to PDF points.
    
    Args:
        pixel_x: X coordinate in pixels
        pixel_y: Y coordinate in pixels  
        dpi: DPI used for rendering
        page_width_pts: Page width in PDF points
        page_height_pts: Page height in PDF points
        
    Returns:
        Tuple of (x_pts, y_pts) in PDF coordinate system
    """
    # Convert pixels to inches, then to points
    inch_x = pixel_x / dpi
    inch_y = pixel_y / dpi
    
    pts_x = inch_x * PDF_POINTS_PER_INCH
    pts_y = inch_y * PDF_POINTS_PER_INCH
    
    return pts_x, pts_y


def normalize_coordinates(x0: float, y0: float, x1: float, y1: float) -> Tuple[float, float, float, float]:
    """
    Ensure coordinates are in correct order (top-left to bottom-right).
    
    Args:
        x0, y0, x1, y1: Coordinates that may be in any order
        
    Returns:
        Normalized coordinates (min_x, min_y, max_x, max_y)
    """
    min_x = min(x0, x1)
    max_x = max(x0, x1)
    min_y = min(y0, y1)
    max_y = max(y0, y1)
    
    return min_x, min_y, max_x, max_y


def calculate_aspect_ratio(width: float, height: float) -> float:
    """
    Calculate aspect ratio (width/height).
    
    Args:
        width: Width dimension
        height: Height dimension
        
    Returns:
        Aspect ratio
    """
    if height == 0:
        return 1.0
    return width / height


def scale_coordinates(x: float, y: float, scale_x: float, scale_y: float) -> Tuple[float, float]:
    """
    Scale coordinates by given factors.
    
    Args:
        x, y: Original coordinates
        scale_x, scale_y: Scaling factors
        
    Returns:
        Scaled coordinates
    """
    return x * scale_x, y * scale_y


def apply_margin(x0: float, y0: float, x1: float, y1: float, 
                margin_x: float, margin_y: float, 
                max_width: float, max_height: float) -> Tuple[float, float, float, float]:
    """
    Apply margin to coordinates while keeping them within bounds.
    
    Args:
        x0, y0, x1, y1: Original coordinates
        margin_x, margin_y: Margin amounts to apply
        max_width, max_height: Maximum bounds
        
    Returns:
        Adjusted coordinates with margin applied
    """
    # Apply margin
    new_x0 = max(0, x0 + margin_x)
    new_y0 = max(0, y0 + margin_y)
    new_x1 = min(max_width, x1 - margin_x)
    new_y1 = min(max_height, y1 - margin_y)
    
    # Ensure minimum size
    if new_x1 <= new_x0:
        new_x1 = min(max_width, new_x0 + 50)  # Minimum 50 points width
    if new_y1 <= new_y0:
        new_y1 = min(max_height, new_y0 + 20)  # Minimum 20 points height
    
    return new_x0, new_y0, new_x1, new_y1
