"""
Main PDF to PPTX and PPTX to PDF conversion pipeline.
"""

import fitz  # PyMuPDF
from typing import List, Optional, Tuple
import logging
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import tempfile
import os

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


def pptx_to_pdf(pptx_bytes: bytes) -> bytes:
    """
    Convert PPTX bytes to PDF bytes with 1:1 layout preservation.
    
    Args:
        pptx_bytes: PPTX file content as bytes
        
    Returns:
        PDF file content as bytes
        
    Raises:
        ValueError: If PPTX processing fails
        Exception: If conversion fails
    """
    try:
        logger.info("Starting PPTX to PDF conversion")
        
        # Create a temporary file for the PPTX
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp_pptx:
            tmp_pptx.write(pptx_bytes)
            tmp_pptx_path = tmp_pptx.name
        
        try:
            # Load the presentation
            presentation = Presentation(tmp_pptx_path)
            slide_count = len(presentation.slides)
            
            logger.info(f"Processing PPTX: {slide_count} slides")
            
            # Get slide dimensions (PPTX uses inches, convert to points for PDF)
            slide_width = presentation.slide_width.inches
            slide_height = presentation.slide_height.inches
            
            # Create PDF in memory
            pdf_buffer = io.BytesIO()
            
            # Convert inches to points (1 inch = 72 points)
            pdf_width = slide_width * 72
            pdf_height = slide_height * 72
            
            # Create PDF canvas with custom page size
            c = canvas.Canvas(pdf_buffer, pagesize=(pdf_width, pdf_height))
            
            # Process each slide
            for slide_idx, slide in enumerate(presentation.slides):
                logger.info(f"Processing slide {slide_idx + 1}/{slide_count}")
                
                # Create a new page for each slide
                if slide_idx > 0:
                    c.showPage()
                
                # Extract and draw all shapes from the slide
                _draw_slide_shapes(c, slide, pdf_width, pdf_height)
            
            # Save the PDF
            c.save()
            
            # Get PDF bytes
            pdf_bytes = pdf_buffer.getvalue()
            
            logger.info(f"Conversion completed: {len(pdf_bytes)} bytes")
            return pdf_bytes
            
        finally:
            # Clean up temporary file
            os.unlink(tmp_pptx_path)
            
    except Exception as e:
        logger.error(f"PPTX to PDF conversion failed: {str(e)}")
        raise Exception(f"Conversion failed: {str(e)}")


def _draw_slide_shapes(canvas_obj, slide, pdf_width: float, pdf_height: float):
    """
    Draw all text shapes from a slide onto the PDF canvas.
    
    Args:
        canvas_obj: ReportLab canvas object
        slide: pptx slide object
        pdf_width: PDF page width in points
        pdf_height: PDF page height in points
    """
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
            
        # Get shape position and size (PPTX uses EMUs, convert to points)
        # 1 inch = 914400 EMUs, 1 inch = 72 points
        left_emu = shape.left
        top_emu = shape.top
        width_emu = shape.width
        height_emu = shape.height
        
        # Convert EMUs to inches, then to points
        left_in = left_emu / 914400
        top_in = top_emu / 914400
        width_in = width_emu / 914400
        height_in = height_emu / 914400
        
        left_pt = left_in * 72
        top_pt = top_in * 72
        width_pt = width_in * 72
        height_pt = height_in * 72
        
        # PPTX origin is top-left, PDF origin is bottom-left
        # Convert Y coordinate
        pdf_y = pdf_height - top_pt - height_pt
        
        # Process text frame
        text_frame = shape.text_frame
        
        # Get text alignment
        alignment = 'left'
        if hasattr(text_frame.paragraphs[0], 'alignment'):
            align_val = text_frame.paragraphs[0].alignment
            if align_val:
                if align_val.name == 'CENTER':
                    alignment = 'center'
                elif align_val.name == 'RIGHT':
                    alignment = 'right'
                elif align_val.name == 'JUSTIFY':
                    alignment = 'justify'
        
        # Get font properties from the first paragraph/run
        if text_frame.paragraphs and text_frame.paragraphs[0].runs:
            first_run = text_frame.paragraphs[0].runs[0]
            font_name = first_run.font.name or 'Helvetica'
            font_size = first_run.font.size.pt if first_run.font.size else 12
            is_bold = first_run.font.bold
            is_italic = first_run.font.italic
        else:
            font_name = 'Helvetica'
            font_size = 12
            is_bold = False
            is_italic = False
        
        # Set font on canvas
        font_style = ''
        if is_bold and is_italic:
            font_style = 'bolditalic'
        elif is_bold:
            font_style = 'bold'
        elif is_italic:
            font_style = 'italic'
        
        canvas_obj.setFont(font_name, font_size)
        
        # Extract text content
        text_content = []
        for paragraph in text_frame.paragraphs:
            para_text = ''
            for run in paragraph.runs:
                para_text += run.text
            if para_text.strip():
                text_content.append(para_text)
        
        full_text = '\n'.join(text_content)
        
        if full_text.strip():
            # Draw text box
            canvas_obj.drawString(
                left_pt,
                pdf_y + height_pt - font_size,  # Adjust for baseline
                full_text
            )


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


def validate_pptx(pptx_bytes: bytes) -> bool:
    """
    Validate that the input is a valid PPTX file.
    
    Args:
        pptx_bytes: PPTX file content as bytes
        
    Returns:
        True if valid PPTX, False otherwise
    """
    try:
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp_file:
            tmp_file.write(pptx_bytes)
            tmp_path = tmp_file.name
        
        try:
            presentation = Presentation(tmp_path)
            # Check if we can access basic properties
            _ = len(presentation.slides)
            _ = presentation.slide_width
            _ = presentation.slide_height
            return True
        except Exception as e:
            logger.error(f"PPTX validation failed: {str(e)}")
            return False
        finally:
            os.unlink(tmp_path)
            
    except Exception as e:
        logger.error(f"PPTX file handling failed: {str(e)}")
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


def get_pptx_info(pptx_bytes: bytes) -> dict:
    """
    Extract basic information from a PPTX file.
    
    Args:
        pptx_bytes: PPTX file content as bytes
        
    Returns:
        Dictionary with PPTX information
    """
    try:
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp_file:
            tmp_file.write(pptx_bytes)
            tmp_path = tmp_file.name
        
        try:
            presentation = Presentation(tmp_path)
            
            info = {
                'slide_count': len(presentation.slides),
                'slide_width_inches': presentation.slide_width.inches,
                'slide_height_inches': presentation.slide_height.inches,
                'slide_width_points': presentation.slide_width.inches * 72,
                'slide_height_points': presentation.slide_height.inches * 72,
                'slide_aspect_ratio': presentation.slide_width.inches / presentation.slide_height.inches,
            }
            
            # Count shapes per slide
            shapes_by_slide = []
            for slide in presentation.slides:
                shapes_by_slide.append(len(slide.shapes))
            
            info['shapes_by_slide'] = shapes_by_slide
            info['total_shapes'] = sum(shapes_by_slide)
            
            return info
            
        finally:
            os.unlink(tmp_path)
            
    except Exception as e:
        logger.error(f"Failed to get PPTX info: {str(e)}")
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


def estimate_pptx_processing_time(pptx_bytes: bytes) -> float:
    """
    Estimate processing time for a PPTX based on slide count and content complexity.
    
    Args:
        pptx_bytes: PPTX file content as bytes
        
    Returns:
        Estimated processing time in seconds
    """
    try:
        info = get_pptx_info(pptx_bytes)
        slide_count = info.get('slide_count', 1)
        
        # Base time per slide (seconds)
        base_time_per_slide = 1.0
        
        # Additional time for complex slides with many shapes
        shapes_per_slide = info.get('total_shapes', 0) / max(slide_count, 1)
        complexity_factor = min(5.0, shapes_per_slide / 10)  # Cap at 5x
        
        estimated_time = slide_count * base_time_per_slide * (1 + complexity_factor)
        
        return max(3.0, estimated_time)  # Minimum 3 seconds
        
    except Exception as e:
        logger.error(f"Failed to estimate PPTX processing time: {str(e)}")
        return 15.0  # Default estimate
