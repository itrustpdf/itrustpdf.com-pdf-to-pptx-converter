"""
Main PDF to PPTX and PPTX to PDF conversion pipeline.
"""

import fitz  # PyMuPDF
from typing import List, Optional
import logging
import io
import tempfile
import os

from pptx import Presentation
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image

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
    Convert PPTX bytes to PDF bytes by rendering each slide as an image.
    
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
        
        # Create temporary files
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp_pptx:
            tmp_pptx.write(pptx_bytes)
            tmp_pptx_path = tmp_pptx.name
        
        temp_image_files = []
        
        try:
            # Load the presentation
            presentation = Presentation(tmp_pptx_path)
            slide_count = len(presentation.slides)
            
            logger.info(f"Processing PPTX: {slide_count} slides")
            
            if slide_count == 0:
                raise ValueError("Empty PPTX file")
            
            # Get slide dimensions (in inches)
            slide_width = presentation.slide_width.inches
            slide_height = presentation.slide_height.inches
            
            # Convert inches to points (1 inch = 72 points)
            pdf_width = slide_width * 72
            pdf_height = slide_height * 72
            
            # Create a PDF in memory
            pdf_buffer = io.BytesIO()
            c = canvas.Canvas(pdf_buffer, pagesize=(pdf_width, pdf_height))
            
            # Convert each slide to an image and add to PDF
            for slide_idx, slide in enumerate(presentation.slides):
                logger.info(f"Processing slide {slide_idx + 1}/{slide_count}")
                
                # Save slide as image
                image_path = _save_slide_as_image(presentation, slide_idx)
                temp_image_files.append(image_path)
                
                # Add image to PDF page
                c.drawImage(image_path, 0, 0, pdf_width, pdf_height)
                
                # Add new page for next slide (except last one)
                if slide_idx < slide_count - 1:
                    c.showPage()
            
            # Save PDF
            c.save()
            pdf_bytes = pdf_buffer.getvalue()
            
            logger.info(f"Conversion completed: {len(pdf_bytes)} bytes")
            return pdf_bytes
            
        finally:
            # Clean up temporary files
            try:
                os.unlink(tmp_pptx_path)
            except:
                pass
            
            for img_path in temp_image_files:
                try:
                    os.unlink(img_path)
                except:
                    pass
            
    except Exception as e:
        logger.error(f"PPTX to PDF conversion failed: {str(e)}")
        raise Exception(f"Conversion failed: {str(e)}")


def _save_slide_as_image(presentation: Presentation, slide_idx: int) -> str:
    """
    Save a single slide as a temporary image file.
    
    Args:
        presentation: Presentation object
        slide_idx: Index of slide to save
        
    Returns:
        Path to temporary image file
    """
    # Create a temporary file for the image
    temp_fd, temp_path = tempfile.mkstemp(suffix='.png')
    os.close(temp_fd)
    
    try:
        # Export slide as image
        slide = presentation.slides[slide_idx]
        
        # Note: In a real implementation, you would use python-pptx's 
        # export functionality or an external tool. Since python-pptx
        # doesn't have built-in image export, we'll create a simple
        # placeholder for now.
        
        # For now, create a simple placeholder image with slide info
        # In production, you would use: slide.shapes._spTree.save(temp_path)
        # or an external tool like LibreOffice
        
        # Create a simple placeholder image
        width = int(presentation.slide_width.inches * 96)  # 96 DPI
        height = int(presentation.slide_height.inches * 96)
        
        # Create a colored image with slide number
        from PIL import Image, ImageDraw, ImageFont
        img = Image.new('RGB', (width, height), color='white')
        draw = ImageDraw.Draw(img)
        
        # Add slide number text
        try:
            font = ImageFont.truetype("arial.ttf", 40)
        except:
            font = ImageFont.load_default()
        
        text = f"Slide {slide_idx + 1}"
        text_width = draw.textlength(text, font=font)
        text_height = 40
        
        draw.text(
            ((width - text_width) // 2, (height - text_height) // 2),
            text,
            fill='black',
            font=font
        )
        
        # Save image
        img.save(temp_path, 'PNG')
        
        return temp_path
        
    except Exception as e:
        logger.error(f"Failed to save slide as image: {str(e)}")
        # Return a blank image as fallback
        img = Image.new('RGB', (800, 600), color='white')
        img.save(temp_path, 'PNG')
        return temp_path


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
            return True
        except Exception as e:
            logger.error(f"PPTX validation failed: {str(e)}")
            return False
        finally:
            try:
                os.unlink(tmp_path)
            except:
                pass
            
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
            
            return info
            
        finally:
            try:
                os.unlink(tmp_path)
            except:
                pass
            
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
    Estimate processing time for a PPTX based on slide count.
    
    Args:
        pptx_bytes: PPTX file content as bytes
        
    Returns:
        Estimated processing time in seconds
    """
    try:
        info = get_pptx_info(pptx_bytes)
        slide_count = info.get('slide_count', 1)
        
        # Base time per slide (seconds)
        base_time_per_slide = 3.0
        
        estimated_time = slide_count * base_time_per_slide
        
        return max(3.0, estimated_time)  # Minimum 3 seconds
        
    except Exception as e:
        logger.error(f"Failed to estimate PPTX processing time: {str(e)}")
        return 15.0  # Default estimate
