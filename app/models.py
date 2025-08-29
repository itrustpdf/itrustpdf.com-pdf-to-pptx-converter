"""
Data models and type definitions for the PDF to PPTX converter.
"""

from typing import Tuple, List

# Core data types
TextBlock = Tuple[float, float, float, float, str]
"""Text block with coordinates and content: (x0, y0, x1, y1, text) in PDF points"""

PageDimensions = Tuple[float, float]
"""Page dimensions: (width_pts, height_pts) in PDF points"""

# Constants
PDF_POINTS_PER_INCH = 72.0
PPTX_EMU_PER_INCH = 914400
WIDESCREEN_ASPECT_RATIO = 16.0 / 9.0
MINIMUM_TEXT_THRESHOLD = 20
DEFAULT_OCR_DPI = 300
SLIDE_MARGIN_FACTOR = 0.02  # 2% margin


class SlideConfig:
    """Configuration for slide dimensions and layout."""
    
    def __init__(self, width_emu: int, height_emu: int, margin_factor: float = SLIDE_MARGIN_FACTOR):
        self.width_emu = width_emu
        self.height_emu = height_emu
        self.margin_factor = margin_factor
    
    @property
    def width_pts(self) -> float:
        """Convert EMU width to PDF points."""
        return self.width_emu / PPTX_EMU_PER_INCH * PDF_POINTS_PER_INCH
    
    @property
    def height_pts(self) -> float:
        """Convert EMU height to PDF points."""
        return self.height_emu / PPTX_EMU_PER_INCH * PDF_POINTS_PER_INCH


def validate_text_block(block: TextBlock) -> bool:
    """
    Validate that a text block has valid coordinates and non-empty text.
    
    Args:
        block: Text block tuple (x0, y0, x1, y1, text)
        
    Returns:
        True if valid, False otherwise
    """
    if len(block) != 5:
        return False
    
    x0, y0, x1, y1, text = block
    
    # Check coordinate validity
    if not all(isinstance(coord, (int, float)) for coord in [x0, y0, x1, y1]):
        return False
    
    if x1 <= x0 or y1 <= y0:
        return False
    
    # Check text validity
    if not isinstance(text, str) or not text.strip():
        return False
    
    return True


def calculate_text_area(block: TextBlock) -> float:
    """
    Calculate the area of a text block in square points.
    
    Args:
        block: Text block tuple (x0, y0, x1, y1, text)
        
    Returns:
        Area in square points
    """
    x0, y0, x1, y1, _ = block
    return (x1 - x0) * (y1 - y0)


def pdf_points_to_emu(points: float) -> int:
    """
    Convert PDF points to PowerPoint EMU units.
    
    Args:
        points: Value in PDF points
        
    Returns:
        Value in EMU units
    """
    return int(points * PPTX_EMU_PER_INCH / PDF_POINTS_PER_INCH)


def emu_to_pdf_points(emu: int) -> float:
    """
    Convert PowerPoint EMU units to PDF points.
    
    Args:
        emu: Value in EMU units
        
    Returns:
        Value in PDF points
    """
    return emu * PDF_POINTS_PER_INCH / PPTX_EMU_PER_INCH
