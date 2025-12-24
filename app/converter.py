"""
Main PDF to PPTX and PPTX to PDF conversion pipeline.
"""

import fitz  # PyMuPDF
from typing import List, Optional
import logging
import io
import tempfile
import os
import subprocess
import sys

from pptx import Presentation
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PIL import Image
import atexit

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
    Convert PPTX bytes to PDF bytes by converting each slide to an image.
    
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
        
        # Create temporary directory for all files
        temp_dir = tempfile.mkdtemp()
        temp_files_to_clean = []
        
        try:
            # Save PPTX to temporary file
            pptx_path = os.path.join(temp_dir, "input.pptx")
            with open(pptx_path, 'wb') as f:
                f.write(pptx_bytes)
            temp_files_to_clean.append(pptx_path)
            
            # Convert PPTX to images using different methods
            image_paths = _convert_pptx_to_images(pptx_path, temp_dir)
            
            if not image_paths:
                raise ValueError("Failed to convert any slides to images")
            
            # Get slide dimensions from the first image
            first_image = Image.open(image_paths[0])
            image_width, image_height = first_image.size
            
            # Convert image pixels to PDF points (72 DPI standard)
            # Assuming images are saved at 96 DPI, adjust if needed
            dpi_ratio = 72 / 96  # Convert from 96 DPI to 72 DPI (standard PDF)
            pdf_width = image_width * dpi_ratio
            pdf_height = image_height * dpi_ratio
            
            # Create PDF from images
            pdf_buffer = io.BytesIO()
            c = canvas.Canvas(pdf_buffer, pagesize=(pdf_width, pdf_height))
            
            for i, image_path in enumerate(image_paths):
                logger.info(f"Adding slide {i + 1}/{len(image_paths)} to PDF")
                
                # Add image to PDF page
                c.drawImage(image_path, 0, 0, pdf_width, pdf_height)
                
                # Add new page for next slide (except last one)
                if i < len(image_paths) - 1:
                    c.showPage()
            
            # Save PDF
            c.save()
            pdf_bytes = pdf_buffer.getvalue()
            
            logger.info(f"Conversion completed: {len(pdf_bytes)} bytes")
            return pdf_bytes
            
        finally:
            # Clean up all temporary files
            for file_path in temp_files_to_clean:
                try:
                    os.unlink(file_path)
                except:
                    pass
            
            # Clean up image files
            for image_path in image_paths if 'image_paths' in locals() else []:
                try:
                    os.unlink(image_path)
                except:
                    pass
            
            # Remove temp directory
            try:
                os.rmdir(temp_dir)
            except:
                pass
            
    except Exception as e:
        logger.error(f"PPTX to PDF conversion failed: {str(e)}")
        raise Exception(f"Conversion failed: {str(e)}")


def _convert_pptx_to_images(pptx_path: str, output_dir: str) -> List[str]:
    """
    Convert PPTX slides to images using available methods.
    
    Args:
        pptx_path: Path to PPTX file
        output_dir: Directory to save images
        
    Returns:
        List of paths to generated image files
    """
    image_paths = []
    
    # Method 1: Try using LibreOffice (recommended for production)
    if _try_libreoffice_conversion(pptx_path, output_dir):
        # Find generated PNG files
        for filename in sorted(os.listdir(output_dir)):
            if filename.endswith('.png') and filename.startswith('input_'):
                image_paths.append(os.path.join(output_dir, filename))
    
    # Method 2: If LibreOffice failed, use python-pptx with fallback
    if not image_paths:
        image_paths = _convert_with_python_pptx(pptx_path, output_dir)
    
    # Sort by slide number
    image_paths.sort()
    
    logger.info(f"Converted {len(image_paths)} slides to images")
    return image_paths


def _try_libreoffice_conversion(pptx_path: str, output_dir: str) -> bool:
    """
    Try to convert PPTX to images using LibreOffice.
    
    Args:
        pptx_path: Path to PPTX file
        output_dir: Directory to save images
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Check if LibreOffice is available
        result = subprocess.run(['which', 'libreoffice'], 
                              capture_output=True, text=True)
        if result.returncode != 0:
            logger.warning("LibreOffice not found in PATH")
            return False
        
        # Convert PPTX to PNG using LibreOffice
        cmd = [
            'libreoffice',
            '--headless',
            '--convert-to', 'png',
            '--outdir', output_dir,
            pptx_path
        ]
        
        logger.info(f"Running LibreOffice conversion: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        
        if result.returncode == 0:
            logger.info("LibreOffice conversion successful")
            return True
        else:
            logger.warning(f"LibreOffice conversion failed: {result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        logger.warning("LibreOffice conversion timed out")
        return False
    except Exception as e:
        logger.warning(f"LibreOffice conversion error: {str(e)}")
        return False


def _convert_with_python_pptx(pptx_path: str, output_dir: str) -> List[str]:
    """
    Convert PPTX to images using python-pptx with a workaround.
    
    Args:
        pptx_path: Path to PPTX file
        output_dir: Directory to save images
        
    Returns:
        List of paths to generated image files
    """
    try:
        from pptx import Presentation
        from pptx.shapes.picture import Picture
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        from io import BytesIO
        
        presentation = Presentation(pptx_path)
        image_paths = []
        
        for slide_idx, slide in enumerate(presentation.slides):
            # Create a simple image representation of the slide
            # This is a fallback method - in production, use LibreOffice
            
            # Get slide dimensions
            slide_width = presentation.slide_width.inches
            slide_height = presentation.slide_height.inches
            
            # Convert to pixels (96 DPI)
            width_px = int(slide_width * 96)
            height_px = int(slide_height * 96)
            
            # Create image with slide content
            image_path = os.path.join(output_dir, f"slide_{slide_idx + 1:03d}.png")
            
            # Try to extract and combine shapes
            _create_slide_image(slide, width_px, height_px, image_path)
            
            image_paths.append(image_path)
            
        return image_paths
        
    except Exception as e:
        logger.error(f"Python-pptx conversion failed: {str(e)}")
        return []


def _create_slide_image(slide, width_px: int, height_px: int, output_path: str):
    """
    Create an image representation of a slide.
    
    Args:
        slide: PPTX slide object
        width_px: Image width in pixels
        height_px: Image height in pixels
        output_path: Path to save the image
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
        import textwrap
        
        # Create a white background
        img = Image.new('RGB', (width_px, height_px), color='white')
        draw = ImageDraw.Draw(img)
        
        # Try to load a font
        try:
            font_large = ImageFont.truetype("arial.ttf", 24)
            font_medium = ImageFont.truetype("arial.ttf", 16)
            font_small = ImageFont.truetype("arial.ttf", 12)
        except:
            font_large = ImageFont.load_default()
            font_medium = ImageFont.load_default()
            font_small = ImageFont.load_default()
        
        # Draw slide title/header
        title = f"Slide {len(image_paths) + 1}"
        title_width = draw.textlength(title, font=font_large)
        draw.text(
            ((width_px - title_width) // 2, 50),
            title,
            fill='darkblue',
            font=font_large
        )
        
        # Extract and draw text from shapes
        y_offset = 100
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                text = shape.text.strip()
                
                # Wrap text to fit width
                max_width = width_px - 100
                wrapped_lines = textwrap.wrap(text, width=40)
                
                for line in wrapped_lines:
                    if y_offset < height_px - 50:
                        line_width = draw.textlength(line, font=font_small)
                        draw.text(
                            ((width_px - line_width) // 2, y_offset),
                            line,
                            fill='black',
                            font=font_small
                        )
                        y_offset += 20
        
        # Save the image
        img.save(output_path, 'PNG', quality=95)
        
    except Exception as e:
        logger.error(f"Failed to create slide image: {str(e)}")
        # Create a simple fallback image
        img = Image.new('RGB', (width_px, height_px), color='lightgray')
        draw = ImageDraw.Draw(img)
        
        try:
            font = ImageFont.truetype("arial.ttf", 20)
        except:
            font = ImageFont.load_default()
        
        text = f"Slide {len(image_paths) + 1}"
        text_width = draw.textlength(text, font=font)
        draw.text(
            ((width_px - text_width) // 2, height_px // 2 - 10),
            text,
            fill='darkred',
            font=font
        )
        
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
        base_time_per_slide = 2.0
        
        estimated_time = slide_count * base_time_per_slide
        
        return max(3.0, estimated_time)  # Minimum 3 seconds
        
    except Exception as e:
        logger.error(f"Failed to estimate PPTX processing time: {str(e)}")
        return 15.0  # Default estimate
