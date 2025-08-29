"""
OCR processing module using Tesseract for scanned PDF pages.
"""

import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from typing import List, Optional
import io
import logging

from .models import TextBlock, DEFAULT_OCR_DPI
from .utils import pixels_to_pdf_points, normalize_coordinates

logger = logging.getLogger(__name__)


def ocr_page_lines(page: fitz.Page, dpi: int = DEFAULT_OCR_DPI, 
                  langs: str = 'eng') -> List[TextBlock]:
    """
    Perform OCR on a PDF page and return line-level text blocks.
    
    Args:
        page: PyMuPDF page object
        dpi: DPI for page rendering (default 300)
        langs: Tesseract language codes (default 'eng')
        
    Returns:
        List of text blocks with coordinates and text content
        
    Raises:
        Exception: If OCR processing fails
    """
    try:
        # Render page to PNG image
        mat = fitz.Matrix(dpi / 72, dpi / 72)  # Scale matrix for DPI
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        pix = None  # Free memory
        
        # Convert to PIL Image
        image = Image.open(io.BytesIO(img_data))
        
        # Perform OCR with line-level data
        ocr_data = pytesseract.image_to_data(
            image, 
            lang=langs,
            output_type=pytesseract.Output.DICT,
            config='--psm 6'  # Uniform block of text
        )
        
        # Get page dimensions for coordinate conversion
        page_rect = page.rect
        page_width_pts = page_rect.width
        page_height_pts = page_rect.height
        
        # Group words into lines and create text blocks
        text_blocks = _group_words_into_lines(
            ocr_data, dpi, page_width_pts, page_height_pts
        )
        
        logger.info(f"OCR extracted {len(text_blocks)} text blocks from page")
        return text_blocks
        
    except Exception as e:
        logger.error(f"OCR processing failed: {str(e)}")
        raise Exception(f"OCR processing failed: {str(e)}")


def _group_words_into_lines(ocr_data: dict, dpi: int, 
                           page_width_pts: float, page_height_pts: float) -> List[TextBlock]:
    """
    Group OCR words into lines based on Tesseract's line detection.
    
    Args:
        ocr_data: Tesseract OCR output data
        dpi: DPI used for rendering
        page_width_pts: Page width in PDF points
        page_height_pts: Page height in PDF points
        
    Returns:
        List of text blocks grouped by lines
    """
    lines = {}  # line_num -> {words: [], bbox: [min_x, min_y, max_x, max_y]}
    
    # Group words by line
    for i, text in enumerate(ocr_data['text']):
        # Skip empty text and low confidence words
        if not text.strip() or ocr_data['conf'][i] < 30:
            continue
        
        line_num = ocr_data['line_num'][i]
        if line_num not in lines:
            lines[line_num] = {'words': [], 'bbox': [float('inf'), float('inf'), 0, 0]}
        
        # Convert pixel coordinates to PDF points
        x = ocr_data['left'][i]
        y = ocr_data['top'][i]
        w = ocr_data['width'][i]
        h = ocr_data['height'][i]
        
        x_pts, y_pts = pixels_to_pdf_points(x, y, dpi, page_width_pts, page_height_pts)
        x1_pts, y1_pts = pixels_to_pdf_points(x + w, y + h, dpi, page_width_pts, page_height_pts)
        
        # Update line bounding box
        bbox = lines[line_num]['bbox']
        bbox[0] = min(bbox[0], x_pts)      # min_x
        bbox[1] = min(bbox[1], y_pts)      # min_y
        bbox[2] = max(bbox[2], x1_pts)     # max_x
        bbox[3] = max(bbox[3], y1_pts)     # max_y
        
        lines[line_num]['words'].append(text)
    
    # Create text blocks from lines
    text_blocks = []
    for line_num, line_data in lines.items():
        if not line_data['words']:
            continue
        
        # Combine words into line text
        line_text = ' '.join(line_data['words']).strip()
        if not line_text:
            continue
        
        # Get normalized coordinates
        bbox = line_data['bbox']
        if bbox[0] == float('inf'):  # No valid words in line
            continue
        
        x0, y0, x1, y1 = normalize_coordinates(bbox[0], bbox[1], bbox[2], bbox[3])
        
        # Ensure minimum dimensions
        if x1 - x0 < 5:  # Minimum 5 points width
            x1 = x0 + 5
        if y1 - y0 < 5:  # Minimum 5 points height
            y1 = y0 + 5
        
        text_blocks.append((x0, y0, x1, y1, line_text))
    
    return text_blocks


def test_tesseract_installation() -> bool:
    """
    Test if Tesseract is properly installed and accessible.
    
    Returns:
        True if Tesseract is working, False otherwise
    """
    try:
        # Test with a simple image
        test_image = Image.new('RGB', (100, 50), color='white')
        result = pytesseract.image_to_string(test_image)
        return True
    except Exception as e:
        logger.error(f"Tesseract test failed: {str(e)}")
        return False


def get_tesseract_version() -> Optional[str]:
    """
    Get the version of the installed Tesseract.
    
    Returns:
        Version string if available, None otherwise
    """
    try:
        version = pytesseract.get_tesseract_version()
        return str(version)
    except Exception as e:
        logger.error(f"Failed to get Tesseract version: {str(e)}")
        return None
