"""
Text extraction and normalization module for PDF processing.
"""

import fitz  # PyMuPDF
import re
from typing import List, Tuple
import logging

from .models import TextBlock, MINIMUM_TEXT_THRESHOLD
from .utils import normalize_coordinates

logger = logging.getLogger(__name__)


def extract_text_blocks_pymupdf(page: fitz.Page) -> List[TextBlock]:
    """
    Extract text blocks from a PDF page using PyMuPDF.
    
    Args:
        page: PyMuPDF page object
        
    Returns:
        List of text blocks with coordinates and content
    """
    try:
        # Get text blocks from the page
        text_dict = page.get_text("dict")
        blocks = text_dict["blocks"]
        text_blocks = []
        
        for block in blocks:
            # Skip image blocks
            if "image" in block:
                continue
            
            # Process text blocks
            if "lines" in block:
                block_text_lines = []
                block_bbox = block["bbox"]
                
                for line in block["lines"]:
                    line_text = ""
                    for span in line["spans"]:
                        line_text += span["text"]
                    
                    line_text = line_text.strip()
                    if line_text:
                        block_text_lines.append(line_text)
                
                # Combine lines into block text
                if block_text_lines:
                    block_text = "\n".join(block_text_lines)
                    x0, y0, x1, y1 = normalize_coordinates(
                        block_bbox[0], block_bbox[1], block_bbox[2], block_bbox[3]
                    )
                    text_blocks.append((x0, y0, x1, y1, block_text))
        
        logger.info(f"Extracted {len(text_blocks)} text blocks using PyMuPDF")
        return text_blocks
        
    except Exception as e:
        logger.error(f"PyMuPDF text extraction failed: {str(e)}")
        return []


def has_sufficient_text(text_blocks: List[TextBlock]) -> bool:
    """
    Check if the extracted text blocks contain sufficient text.
    
    Args:
        text_blocks: List of text blocks
        
    Returns:
        True if text is sufficient, False if OCR fallback is needed
    """
    total_chars = sum(len(block[4].strip()) for block in text_blocks)
    return total_chars >= MINIMUM_TEXT_THRESHOLD


def normalize_and_group_text_blocks(text_blocks: List[TextBlock], 
                                   dehyphenate: bool = True) -> List[TextBlock]:
    """
    Normalize text content and group text blocks into natural content flow.
    
    Args:
        text_blocks: List of text blocks to process
        dehyphenate: Whether to remove end-of-line hyphenation
        
    Returns:
        List of normalized and consolidated text blocks for natural presentation
    """
    if not text_blocks:
        return []
    
    # Filter out empty blocks and normalize text
    normalized_blocks = []
    for x0, y0, x1, y1, text in text_blocks:
        # Clean and normalize text
        cleaned_text = _normalize_text(text, dehyphenate)
        if cleaned_text.strip():
            normalized_blocks.append((x0, y0, x1, y1, cleaned_text))
    
    # Sort by reading order first
    sorted_blocks = _sort_by_reading_order(normalized_blocks)
    
    # Group text into larger, more natural content blocks
    content_blocks = _group_into_content_blocks(sorted_blocks)
    
    logger.info(f"Normalized {len(text_blocks)} blocks to {len(content_blocks)} content blocks")
    return content_blocks


def _group_into_content_blocks(text_blocks: List[TextBlock]) -> List[TextBlock]:
    """
    Group text blocks into larger content blocks for natural presentation flow.
    
    Args:
        text_blocks: Sorted text blocks
        
    Returns:
        List of consolidated content blocks
    """
    if not text_blocks:
        return []
    
    content_blocks = []
    current_content = []
    current_y_region = None
    region_tolerance = 50.0  # Tolerance for grouping text in same region
    
    for x0, y0, x1, y1, text in text_blocks:
        text = text.strip()
        if not text:
            continue
            
        # Check if this text belongs to current content region
        if current_y_region is None:
            # Start new region
            current_y_region = (y0, y1)
            current_content = [text]
        else:
            # Check if this text is in similar vertical region
            region_top, region_bottom = current_y_region
            if abs(y0 - region_top) <= region_tolerance or abs(y1 - region_bottom) <= region_tolerance:
                # Add to current content block
                current_content.append(text)
                # Expand region
                current_y_region = (min(region_top, y0), max(region_bottom, y1))
            else:
                # Save current content block and start new one
                if current_content:
                    combined_text = _combine_content_text(current_content)
                    if combined_text.strip():
                        # Create a content block that spans the slide width for natural flow
                        content_blocks.append((50, region_top, 700, region_bottom, combined_text))
                
                # Start new region
                current_y_region = (y0, y1)
                current_content = [text]
    
    # Don't forget the last content block
    if current_content and current_y_region:
        combined_text = _combine_content_text(current_content)
        if combined_text.strip():
            region_top, region_bottom = current_y_region
            content_blocks.append((50, region_top, 700, region_bottom, combined_text))
    
    return content_blocks


def _combine_content_text(text_pieces: List[str]) -> str:
    """
    Combine text pieces into natural flowing content.
    
    Args:
        text_pieces: List of text strings to combine
        
    Returns:
        Combined text with natural flow
    """
    if not text_pieces:
        return ""
    
    combined = []
    for piece in text_pieces:
        piece = piece.strip()
        if not piece:
            continue
            
        # Add appropriate spacing between pieces
        if combined:
            # Check if we need paragraph break or just space
            last_piece = combined[-1]
            if (piece[0].isupper() and last_piece.endswith('.')) or \
               len(piece) > 50 or \
               piece.startswith('•') or piece.startswith('-'):
                combined.append('\n\n' + piece)  # Paragraph break
            else:
                combined.append(' ' + piece)  # Just space
        else:
            combined.append(piece)
    
    return ''.join(combined)


def _normalize_text(text: str, dehyphenate: bool = True) -> str:
    """
    Normalize text by cleaning whitespace and optionally removing hyphenation.
    
    Args:
        text: Input text to normalize
        dehyphenate: Whether to remove end-of-line hyphenation
        
    Returns:
        Normalized text
    """
    if not text:
        return ""
    
    # Remove excessive whitespace while preserving paragraph breaks
    text = re.sub(r'[ \t]+', ' ', text)  # Multiple spaces/tabs to single space
    text = re.sub(r'\n[ \t]+', '\n', text)  # Remove leading whitespace after newlines
    text = re.sub(r'[ \t]+\n', '\n', text)  # Remove trailing whitespace before newlines
    text = re.sub(r'\n{3,}', '\n\n', text)  # Multiple newlines to double newline
    
    # Remove hyphenation if requested
    if dehyphenate:
        text = _dehyphenate_text(text)
    
    return text.strip()


def _dehyphenate_text(text: str) -> str:
    """
    Remove end-of-line hyphenation from text.
    
    Args:
        text: Input text with potential hyphenation
        
    Returns:
        Text with hyphenation removed
    """
    # Pattern for hyphenated words across line breaks
    # Matches: word- followed by newline and then word continuation
    hyphen_pattern = r'([a-zA-Z])-\s*\n\s*([a-zA-Z])'
    
    # Replace hyphenated line breaks with joined words
    dehyphenated = re.sub(hyphen_pattern, r'\1\2', text)
    
    return dehyphenated


def _sort_by_reading_order(text_blocks: List[TextBlock]) -> List[TextBlock]:
    """
    Sort text blocks by reading order (top-to-bottom, left-to-right).
    
    Args:
        text_blocks: List of text blocks to sort
        
    Returns:
        Sorted text blocks
    """
    if not text_blocks:
        return []
    
    # Sort by y-coordinate (top to bottom), then by x-coordinate (left to right)
    # Use tolerance for y-coordinate to handle slight variations in line height
    def sort_key(block: TextBlock) -> Tuple[float, float]:
        x0, y0, x1, y1, _ = block
        # Use top edge for primary sort, left edge for secondary sort
        return (y0, x0)
    
    return sorted(text_blocks, key=sort_key)


def merge_overlapping_blocks(text_blocks: List[TextBlock], 
                           overlap_threshold: float = 0.5) -> List[TextBlock]:
    """
    Merge text blocks that significantly overlap.
    
    Args:
        text_blocks: List of text blocks
        overlap_threshold: Minimum overlap ratio to trigger merge
        
    Returns:
        List of text blocks with overlapping blocks merged
    """
    if len(text_blocks) < 2:
        return text_blocks
    
    merged_blocks = []
    used_indices = set()
    
    for i, block1 in enumerate(text_blocks):
        if i in used_indices:
            continue
        
        current_block = block1
        merged_with = [i]
        
        for j, block2 in enumerate(text_blocks[i+1:], i+1):
            if j in used_indices:
                continue
            
            if _blocks_overlap(current_block, block2, overlap_threshold):
                # Merge blocks
                current_block = _merge_two_blocks(current_block, block2)
                merged_with.append(j)
        
        merged_blocks.append(current_block)
        used_indices.update(merged_with)
    
    return merged_blocks


def _blocks_overlap(block1: TextBlock, block2: TextBlock, threshold: float) -> bool:
    """
    Check if two text blocks overlap significantly.
    
    Args:
        block1, block2: Text blocks to compare
        threshold: Minimum overlap ratio
        
    Returns:
        True if blocks overlap above threshold
    """
    x0_1, y0_1, x1_1, y1_1, _ = block1
    x0_2, y0_2, x1_2, y1_2, _ = block2
    
    # Calculate intersection
    x_overlap = max(0, min(x1_1, x1_2) - max(x0_1, x0_2))
    y_overlap = max(0, min(y1_1, y1_2) - max(y0_1, y0_2))
    
    if x_overlap == 0 or y_overlap == 0:
        return False
    
    overlap_area = x_overlap * y_overlap
    
    # Calculate areas
    area1 = (x1_1 - x0_1) * (y1_1 - y0_1)
    area2 = (x1_2 - x0_2) * (y1_2 - y0_2)
    
    min_area = min(area1, area2)
    if min_area == 0:
        return False
    
    overlap_ratio = overlap_area / min_area
    return overlap_ratio >= threshold


def _merge_two_blocks(block1: TextBlock, block2: TextBlock) -> TextBlock:
    """
    Merge two overlapping text blocks.
    
    Args:
        block1, block2: Text blocks to merge
        
    Returns:
        Merged text block
    """
    x0_1, y0_1, x1_1, y1_1, text1 = block1
    x0_2, y0_2, x1_2, y1_2, text2 = block2
    
    # Calculate combined bounding box
    new_x0 = min(x0_1, x0_2)
    new_y0 = min(y0_1, y0_2)
    new_x1 = max(x1_1, x1_2)
    new_y1 = max(y1_1, y1_2)
    
    # Combine text (preserve order by y-coordinate)
    if y0_1 <= y0_2:
        combined_text = f"{text1}\n{text2}"
    else:
        combined_text = f"{text2}\n{text1}"
    
    return (new_x0, new_y0, new_x1, new_y1, combined_text)
