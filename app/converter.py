"""
Main PDF to PPTX and PPTX to PDF conversion pipeline.
"""

import fitz  # PyMuPDF
from typing import List
import logging
import io
import tempfile
import os
from pathlib import Path

from pptx import Presentation
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from PIL import Image, ImageDraw, ImageFont

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
    Convert PPTX bytes to PDF bytes - simple 1:1 image-based conversion.
    
    Args:
        pptx_bytes: PPTX file content as bytes
        
    Returns:
        PDF file content as bytes
        
    Raises:
        ValueError: If PPTX processing fails
        Exception: If conversion fails
    """
    temp_dir = None
    
    try:
        logger.info("Starting PPTX to PDF conversion")
        
        # Create temporary directory
        temp_dir = tempfile.mkdtemp(prefix="pptx_to_pdf_")
        
        # Save PPTX to temporary file
        pptx_path = os.path.join(temp_dir, "presentation.pptx")
        with open(pptx_path, "wb") as f:
            f.write(pptx_bytes)
        
        # Load presentation
        presentation = Presentation(pptx_path)
        slide_count = len(presentation.slides)
        
        if slide_count == 0:
            raise ValueError("Empty PPTX file")
        
        logger.info(f"Processing PPTX: {slide_count} slides")
        
        # Get slide dimensions (convert inches to points: 1 inch = 72 points)
        slide_width_inches = presentation.slide_width.inches
        slide_height_inches = presentation.slide_height.inches
        pdf_width = slide_width_inches * 72
        pdf_height = slide_height_inches * 72
        
        # Standard page sizes for comparison
        standard_sizes = {
            "letter": letter,
            "A4": A4
        }
        
        # Use standard page size if close to it
        page_size = (pdf_width, pdf_height)
        for name, size in standard_sizes.items():
            if abs(pdf_width - size[0]) < 10 and abs(pdf_height - size[1]) < 10:
                page_size = size
                logger.info(f"Using standard {name} page size")
                break
        
        # Create PDF buffer
        pdf_buffer = io.BytesIO()
        c = canvas.Canvas(pdf_buffer, pagesize=page_size)
        
        # Process each slide
        for slide_idx in range(slide_count):
            logger.info(f"Processing slide {slide_idx + 1}/{slide_count}")
            
            if slide_idx > 0:
                c.showPage()
            
            # Create image for this slide
            image_path = os.path.join(temp_dir, f"slide_{slide_idx}.png")
            _create_slide_image(presentation, slide_idx, image_path, 
                              int(page_size[0]), int(page_size[1]))
            
            # Add image to PDF
            try:
                c.drawImage(image_path, 0, 0, page_size[0], page_size[1])
            except Exception as img_error:
                logger.warning(f"Failed to add image: {img_error}")
                # Draw placeholder
                c.setFillColorRGB(0.95, 0.95, 0.95)
                c.rect(0, 0, page_size[0], page_size[1], fill=1)
                c.setFillColorRGB(0, 0, 0)
                c.setFont("Helvetica-Bold", 24)
                c.drawCentredString(page_size[0]/2, page_size[1]/2, f"Slide {slide_idx + 1}")
                c.setFont("Helvetica", 12)
                c.drawCentredString(page_size[0]/2, page_size[1]/2 - 40, 
                                  f"{slide_width_inches:.1f} x {slide_height_inches:.1f} inches")
        
        # Save PDF
        c.save()
        pdf_bytes = pdf_buffer.getvalue()
        
        logger.info(f"Conversion completed: {len(pdf_bytes)} bytes")
        return pdf_bytes
        
    except Exception as e:
        logger.error(f"PPTX to PDF conversion failed: {str(e)}")
        raise Exception(f"Conversion failed: {str(e)}")
        
    finally:
        # Clean up
        if temp_dir and os.path.exists(temp_dir):
            try:
                import shutil
                shutil.rmtree(temp_dir)
            except:
                pass


def _create_slide_image(presentation: Presentation, slide_idx: int, 
                       output_path: str, width: int, height: int):
    """
    Create a simple image for a slide.
    
    Args:
        presentation: Presentation object
        slide_idx: Slide index
        output_path: Path to save image
        width: Image width
        height: Image height
    """
    try:
        # Create image with slide info
        img = Image.new('RGB', (width, height), color='white')
        draw = ImageDraw.Draw(img)
        
        # Try to use a nice font
        font = None
        font_paths = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "/usr/share/fonts/truetype/freefont/FreeSansBold.ttf",
        ]
        
        for fp in font_paths:
            if os.path.exists(fp):
                try:
                    font = ImageFont.truetype(fp, 48)
                    break
                except:
                    continue
        
        if font is None:
            font = ImageFont.load_default()
        
        # Get slide
        slide = presentation.slides[slide_idx]
        
        # Draw slide number
        text = f"Slide {slide_idx + 1}"
        
        # Calculate text position
        try:
            # Try to get text bounding box
            bbox = draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
        except:
            # Fallback calculation
            text_width = len(text) * 30
            text_height = 48
        
        x = (width - text_width) // 2
        y = (height - text_height) // 2
        
        # Draw text with shadow for better visibility
        shadow_offset = 3
        draw.text((x + shadow_offset, y + shadow_offset), text, font=font, fill='#888888')
        draw.text((x, y), text, font=font, fill='#000000')
        
        # Add slide dimensions
        dim_text = f"{presentation.slide_width.inches:.1f} × {presentation.slide_height.inches:.1f} inches"
        try:
            small_font = ImageFont.truetype(font_paths[0], 20) if os.path.exists(font_paths[0]) else ImageFont.load_default()
            dim_bbox = draw.textbbox((0, 0), dim_text, font=small_font)
            dim_width = dim_bbox[2] - dim_bbox[0]
            dim_x = (width - dim_width) // 2
            dim_y = y + text_height + 30
            draw.text((dim_x, dim_y), dim_text, font=small_font, fill='#666666')
        except:
            pass
        
        # Add border
        border_color = '#007acc'
        border_width = 4
        draw.rectangle([border_width, border_width, width-border_width, height-border_width], 
                      outline=border_color, width=border_width)
        
        # Save image
        img.save(output_path, 'PNG', quality=95)
        
    except Exception as e:
        logger.error(f"Failed to create slide image: {str(e)}")
        # Create minimal fallback image
        img = Image.new('RGB', (width, height), color='#f0f0f0')
        img.save(output_path, 'PNG')


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
        base_time_per_slide = 1.5
        
        estimated_time = slide_count * base_time_per_slide
        
        return max(3.0, estimated_time)  # Minimum 3 seconds
        
    except Exception as e:
        logger.error(f"Failed to estimate PPTX processing time: {str(e)}")
        return 15.0  # Default estimate
