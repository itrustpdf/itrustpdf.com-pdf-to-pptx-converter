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
import shutil
from pathlib import Path

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
    Convert PPTX bytes to PDF bytes using LibreOffice for high-quality image conversion.
    
    Args:
        pptx_bytes: PPTX file content as bytes
        
    Returns:
        PDF file content as bytes
        
    Raises:
        ValueError: If PPTX processing fails
        Exception: If conversion fails
    """
    temp_dir = None
    temp_pptx_path = None
    
    try:
        logger.info("Starting PPTX to PDF conversion using LibreOffice")
        
        # Create temporary directory for all files
        temp_dir = tempfile.mkdtemp(prefix="pptx_to_pdf_")
        
        # Save PPTX to temporary file
        temp_pptx_path = os.path.join(temp_dir, "presentation.pptx")
        with open(temp_pptx_path, "wb") as f:
            f.write(pptx_bytes)
        
        # Load presentation to get slide count and dimensions
        presentation = Presentation(temp_pptx_path)
        slide_count = len(presentation.slides)
        
        if slide_count == 0:
            raise ValueError("Empty PPTX file")
        
        logger.info(f"Processing PPTX: {slide_count} slides")
        
        # Get slide dimensions
        slide_width = presentation.slide_width.inches
        slide_height = presentation.slide_height.inches
        pdf_width = slide_width * 72  # inches to points
        pdf_height = slide_height * 72
        
        # Method 1: Try using LibreOffice to export slides as images
        logger.info("Attempting LibreOffice conversion...")
        try:
            pdf_bytes = _convert_pptx_to_pdf_libreoffice(temp_pptx_path, temp_dir)
            if pdf_bytes:
                logger.info("LibreOffice conversion successful")
                return pdf_bytes
        except Exception as lo_error:
            logger.warning(f"LibreOffice conversion failed, falling back to basic method: {lo_error}")
        
        # Method 2: Fallback - convert slides to images using python-pptx and PIL
        logger.info("Using fallback conversion method")
        pdf_bytes = _convert_pptx_to_pdf_fallback(presentation, temp_dir, pdf_width, pdf_height)
        
        logger.info(f"Conversion completed: {len(pdf_bytes)} bytes")
        return pdf_bytes
        
    except Exception as e:
        logger.error(f"PPTX to PDF conversion failed: {str(e)}")
        raise Exception(f"Conversion failed: {str(e)}")
        
    finally:
        # Clean up temporary files
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                logger.warning(f"Failed to clean up temp directory: {e}")


def _convert_pptx_to_pdf_libreoffice(pptx_path: str, temp_dir: str) -> Optional[bytes]:
    """
    Convert PPTX to PDF using LibreOffice command line.
    
    Args:
        pptx_path: Path to PPTX file
        temp_dir: Temporary directory for output
        
    Returns:
        PDF bytes or None if conversion fails
    """
    try:
        output_path = os.path.join(temp_dir, "output.pdf")
        
        # Run LibreOffice in headless mode to convert PPTX to PDF
        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", temp_dir,
            pptx_path
        ]
        
        logger.info(f"Running LibreOffice command: {' '.join(cmd)}")
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60  # 60 second timeout
        )
        
        if result.returncode != 0:
            logger.error(f"LibreOffice conversion failed: {result.stderr}")
            return None
        
        # Check for output file
        expected_output = os.path.join(temp_dir, "presentation.pdf")
        if os.path.exists(expected_output):
            output_path = expected_output
        elif os.path.exists(output_path):
            pass  # Use the specified output path
        else:
            # Look for any PDF file in the temp directory
            pdf_files = list(Path(temp_dir).glob("*.pdf"))
            if pdf_files:
                output_path = str(pdf_files[0])
            else:
                logger.error("No PDF output found from LibreOffice")
                return None
        
        # Read the generated PDF
        with open(output_path, "rb") as f:
            pdf_bytes = f.read()
        
        if len(pdf_bytes) == 0:
            logger.error("Empty PDF generated by LibreOffice")
            return None
        
        return pdf_bytes
        
    except subprocess.TimeoutExpired:
        logger.error("LibreOffice conversion timed out")
        return None
    except Exception as e:
        logger.error(f"LibreOffice conversion error: {str(e)}")
        return None


def _convert_pptx_to_pdf_fallback(presentation: Presentation, temp_dir: str, 
                                  pdf_width: float, pdf_height: float) -> bytes:
    """
    Fallback method to convert PPTX to PDF by creating simple slide images.
    
    Args:
        presentation: Presentation object
        temp_dir: Temporary directory
        pdf_width: PDF page width in points
        pdf_height: PDF page height in points
        
    Returns:
        PDF bytes
    """
    slide_count = len(presentation.slides)
    image_paths = []
    
    try:
        # Create simple images for each slide
        for slide_idx in range(slide_count):
            logger.info(f"Creating image for slide {slide_idx + 1}/{slide_count}")
            
            # Create a simple placeholder image
            image_path = os.path.join(temp_dir, f"slide_{slide_idx}.png")
            image_paths.append(image_path)
            
            # Create image with slide information
            _create_slide_image(presentation, slide_idx, image_path, 
                               int(pdf_width), int(pdf_height))
        
        # Create PDF from images
        return _create_pdf_from_images(image_paths, pdf_width, pdf_height)
        
    finally:
        # Clean up image files
        for img_path in image_paths:
            try:
                if os.path.exists(img_path):
                    os.unlink(img_path)
            except:
                pass


def _create_slide_image(presentation: Presentation, slide_idx: int, 
                       output_path: str, width: int, height: int):
    """
    Create a simple image representing a slide.
    
    Args:
        presentation: Presentation object
        slide_idx: Slide index
        output_path: Path to save image
        width: Image width
        height: Image height
    """
    try:
        # Create a colored background based on slide number
        colors = ['#FFFFFF', '#F0F0F0', '#E8F4F8', '#F8F0E8']
        bg_color = colors[slide_idx % len(colors)]
        
        # Convert hex color to RGB
        bg_color = bg_color.lstrip('#')
        bg_rgb = tuple(int(bg_color[i:i+2], 16) for i in (0, 2, 4))
        
        # Create image
        img = Image.new('RGB', (width, height), color=bg_rgb)
        
        # Add slide information
        from PIL import ImageDraw, ImageFont
        
        draw = ImageDraw.Draw(img)
        
        # Try to use a font
        try:
            font_path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
            if os.path.exists(font_path):
                font = ImageFont.truetype(font_path, 48)
            else:
                font = ImageFont.load_default()
        except:
            font = ImageFont.load_default()
        
        # Draw slide number
        text = f"Slide {slide_idx + 1}"
        
        # Calculate text size
        try:
            text_bbox = draw.textbbox((0, 0), text, font=font)
            text_width = text_bbox[2] - text_bbox[0]
            text_height = text_bbox[3] - text_bbox[1]
        except:
            text_width = len(text) * 30
            text_height = 48
        
        # Center the text
        x = (width - text_width) // 2
        y = (height - text_height) // 2
        
        # Draw text with shadow
        shadow_color = (100, 100, 100)
        text_color = (0, 0, 0)
        
        draw.text((x + 2, y + 2), text, font=font, fill=shadow_color)
        draw.text((x, y), text, font=font, fill=text_color)
        
        # Add slide dimensions
        info_text = f"{presentation.slide_width.inches:.1f} x {presentation.slide_height.inches:.1f} inches"
        try:
            info_font = ImageFont.truetype(font_path, 24) if 'font_path' in locals() and os.path.exists(font_path) else ImageFont.load_default()
            info_bbox = draw.textbbox((0, 0), info_text, font=info_font)
            info_width = info_bbox[2] - info_bbox[0]
            info_x = (width - info_width) // 2
            info_y = y + text_height + 20
            draw.text((info_x, info_y), info_text, font=info_font, fill=(100, 100, 100))
        except:
            pass
        
        # Save image
        img.save(output_path, 'PNG', dpi=(300, 300))
        
    except Exception as e:
        logger.error(f"Failed to create slide image: {str(e)}")
        # Create a simple fallback image
        img = Image.new('RGB', (width, height), color='white')
        img.save(output_path, 'PNG')


def _create_pdf_from_images(image_paths: List[str], pdf_width: float, pdf_height: float) -> bytes:
    """
    Create a PDF from a list of image paths.
    
    Args:
        image_paths: List of image file paths
        pdf_width: PDF page width in points
        pdf_height: PDF page height in points
        
    Returns:
        PDF bytes
    """
    pdf_buffer = io.BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=(pdf_width, pdf_height))
    
    for i, image_path in enumerate(image_paths):
        if i > 0:
            c.showPage()
        
        if os.path.exists(image_path):
            try:
                # Add image to PDF page
                c.drawImage(image_path, 0, 0, pdf_width, pdf_height)
            except Exception as e:
                logger.error(f"Failed to add image {image_path} to PDF: {str(e)}")
                # Draw placeholder
                c.setFillColorRGB(0.9, 0.9, 0.9)
                c.rect(0, 0, pdf_width, pdf_height, fill=1)
                c.setFillColorRGB(0, 0, 0)
                c.setFont("Helvetica", 24)
                c.drawCentredString(pdf_width/2, pdf_height/2, f"Slide {i+1}")
    
    c.save()
    return pdf_buffer.getvalue()


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
        
        # Base time per slide (seconds) - LibreOffice is fast
        base_time_per_slide = 2.0
        
        estimated_time = slide_count * base_time_per_slide
        
        return max(3.0, estimated_time)  # Minimum 3 seconds
        
    except Exception as e:
        logger.error(f"Failed to estimate PPTX processing time: {str(e)}")
        return 15.0  # Default estimate
