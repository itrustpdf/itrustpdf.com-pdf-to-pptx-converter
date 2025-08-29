"""
Main PDF to PPTX conversion pipeline.
"""

import fitz  # PyMuPDF
from typing import List, Optional
import logging

from .models import TextBlock, MINIMUM_TEXT_THRESHOLD
from .utils import get_pdf_dimensions
from .text_extraction import extract_text_blocks_pymupdf, has_sufficient_text, normalize_and_group_text_blocks
from .ocr import ocr_page_lines
from .layout import transform_blocks_to_pptx
from .pptx_generator import create_pptx_from_blocks, calculate_optimal_slide_size

logger = logging.getLogger(__name__)


def pdf_to_pptx(pdf_bytes: bytes, 
               ocr_langs: str = 'eng', 
               dehyphenate: bool = True) -> bytes:
    """
    Convert PDF bytes to PPTX bytes.
    
    Args:
        pdf_bytes: PDF file content as bytes
        ocr_langs: Tesseract language codes for OCR
        dehyphenate: Whether to remove end-of-line hyphenation
        
    Returns:
        PPTX file content as bytes
        
    Raises:
        ValueError: If PDF processing fails
        Exception: If conversion fails
    """
    try:
        logger.info("Starting PDF to PPTX conversion")
        
        # Open PDF document
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        if len(doc) == 0:
            raise ValueError("Empty PDF")
        
        # Get PDF dimensions from first page
        first_page = doc[0]
        pdf_width = first_page.rect.width
        pdf_height = first_page.rect.height
        page_count = len(doc)
        
        logger.info(f"Processing PDF: {page_count} pages, {pdf_width}x{pdf_height} points")
        
        # Calculate optimal slide configuration
        slide_config = calculate_optimal_slide_size(pdf_width, pdf_height)
        
        # Process each page
        all_page_blocks = []
        
        for page_num in range(page_count):
            page = doc[page_num]
            logger.info(f"Processing page {page_num + 1}/{page_count}")
            
            # Extract text blocks for this page
            page_blocks = _extract_page_text_blocks(page, ocr_langs)
            
            # Normalize and group text blocks
            normalized_blocks = normalize_and_group_text_blocks(page_blocks, dehyphenate)
            
            # Transform to PPTX coordinates
            transformed_blocks = transform_blocks_to_pptx(
                normalized_blocks, pdf_width, pdf_height, slide_config
            )
            
            all_page_blocks.append(transformed_blocks)
            logger.info(f"Page {page_num + 1}: {len(transformed_blocks)} text blocks")
        
        doc.close()
        
        # Generate PPTX
        pptx_bytes = create_pptx_from_blocks(all_page_blocks, slide_config)
        
        logger.info(f"Conversion completed: {len(pptx_bytes)} bytes")
        return pptx_bytes
        
    except Exception as e:
        logger.error(f"PDF to PPTX conversion failed: {str(e)}")
        raise Exception(f"Conversion failed: {str(e)}")


def _extract_page_text_blocks(page: fitz.Page, ocr_langs: str) -> List[TextBlock]:
    """
    Extract text blocks from a single PDF page using native extraction or OCR.
    
    Args:
        page: PyMuPDF page object
        ocr_langs: Language codes for OCR
        
    Returns:
        List of text blocks
    """
    try:
        # First, try native text extraction
        native_blocks = extract_text_blocks_pymupdf(page)
        
        # Check if we have sufficient text
        if has_sufficient_text(native_blocks):
            logger.debug(f"Using native text extraction: {len(native_blocks)} blocks")
            return native_blocks
        else:
            logger.debug(f"Insufficient native text ({sum(len(b[4]) for b in native_blocks)} chars), using OCR")
            
        # Fall back to OCR
        try:
            ocr_blocks = ocr_page_lines(page, langs=ocr_langs)
            logger.debug(f"OCR extracted {len(ocr_blocks)} blocks")
            return ocr_blocks
            
        except Exception as ocr_error:
            logger.warning(f"OCR failed: {str(ocr_error)}, using native extraction")
            return native_blocks
            
    except Exception as e:
        logger.error(f"Text extraction failed for page: {str(e)}")
        return []  # Return empty blocks for failed pages


def validate_pdf(pdf_bytes: bytes) -> bool:
    """
    Validate that the input is a valid PDF.
    
    Args:
        pdf_bytes: PDF file content as bytes
        
    Returns:
        True if valid PDF, False otherwise
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        is_valid = len(doc) > 0
        doc.close()
        return is_valid
    except Exception as e:
        logger.error(f"PDF validation failed: {str(e)}")
        return False


def get_pdf_info(pdf_bytes: bytes) -> dict:
    """
    Extract basic information from a PDF.
    
    Args:
        pdf_bytes: PDF file content as bytes
        
    Returns:
        Dictionary with PDF information
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        metadata = doc.metadata or {}
        info = {
            'page_count': len(doc),
            'title': metadata.get('title', ''),
            'author': metadata.get('author', ''),
            'subject': metadata.get('subject', ''),
            'creator': metadata.get('creator', ''),
            'producer': metadata.get('producer', ''),
            'creation_date': metadata.get('creationDate', ''),
            'modification_date': metadata.get('modDate', ''),
        }
        
        if len(doc) > 0:
            first_page = doc[0]
            info['page_width'] = first_page.rect.width
            info['page_height'] = first_page.rect.height
            info['page_aspect_ratio'] = first_page.rect.width / first_page.rect.height
        
        doc.close()
        return info
        
    except Exception as e:
        logger.error(f"Failed to get PDF info: {str(e)}")
        return {'error': str(e)}


def estimate_processing_time(pdf_bytes: bytes) -> float:
    """
    Estimate processing time for a PDF based on page count and content complexity.
    
    Args:
        pdf_bytes: PDF file content as bytes
        
    Returns:
        Estimated processing time in seconds
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page_count = len(doc)
        doc.close()
        
        # Base time per page (seconds)
        base_time_per_page = 2.0
        
        # Additional time for OCR-heavy documents
        ocr_time_per_page = 5.0
        
        # Assume 50% of pages might need OCR (conservative estimate)
        estimated_time = (page_count * base_time_per_page) + (page_count * 0.5 * ocr_time_per_page)
        
        return max(5.0, estimated_time)  # Minimum 5 seconds
        
    except Exception as e:
        logger.error(f"Failed to estimate processing time: {str(e)}")
        return 30.0  # Default estimate
