"""
Layout and positioning engine for converting PDF coordinates to PPTX coordinates.
"""

from typing import List, Tuple
import logging

from .models import TextBlock, SlideConfig, pdf_points_to_emu, emu_to_pdf_points
from .utils import apply_margin, scale_coordinates

logger = logging.getLogger(__name__)


def transform_blocks_to_pptx(text_blocks: List[TextBlock], 
                           pdf_width: float, pdf_height: float,
                           slide_config: SlideConfig) -> List[Tuple[int, int, int, int, str]]:
    """
    Transform PDF text blocks to PPTX coordinates with proper scaling.
    
    Args:
        text_blocks: List of text blocks in PDF coordinates
        pdf_width: PDF page width in points
        pdf_height: PDF page height in points
        slide_config: Slide configuration with dimensions
        
    Returns:
        List of text blocks in PPTX EMU coordinates
    """
    if not text_blocks:
        return []
    
    transformed_blocks = []
    
    # Calculate scaling factors
    scale_x = slide_config.width_pts / pdf_width
    scale_y = slide_config.height_pts / pdf_height
    
    # Calculate margin amounts
    margin_x = slide_config.width_pts * slide_config.margin_factor
    margin_y = slide_config.height_pts * slide_config.margin_factor
    
    logger.info(f"Transforming {len(text_blocks)} blocks with scale ({scale_x:.3f}, {scale_y:.3f})")
    
    for x0, y0, x1, y1, text in text_blocks:
        # Scale coordinates
        scaled_x0, scaled_y0 = scale_coordinates(x0, y0, scale_x, scale_y)
        scaled_x1, scaled_y1 = scale_coordinates(x1, y1, scale_x, scale_y)
        
        # Apply margins and bounds checking
        final_x0, final_y0, final_x1, final_y1 = apply_margin(
            scaled_x0, scaled_y0, scaled_x1, scaled_y1,
            margin_x, margin_y,
            slide_config.width_pts, slide_config.height_pts
        )
        
        # Convert to EMU units
        emu_x0 = pdf_points_to_emu(final_x0)
        emu_y0 = pdf_points_to_emu(final_y0)
        emu_x1 = pdf_points_to_emu(final_x1)
        emu_y1 = pdf_points_to_emu(final_y1)
        
        transformed_blocks.append((emu_x0, emu_y0, emu_x1, emu_y1, text))
    
    return transformed_blocks


def calculate_font_size(text_width_emu: int, text_height_emu: int, 
                       text_content: str) -> int:
    """
    Calculate appropriate font size based on text box dimensions and content.
    
    Args:
        text_width_emu: Text box width in EMU
        text_height_emu: Text box height in EMU
        text_content: Text content to fit
        
    Returns:
        Font size in half-points (PPTX format)
    """
    # Convert EMU to approximate points for calculation
    width_pts = emu_to_pdf_points(text_width_emu)
    height_pts = emu_to_pdf_points(text_height_emu)
    
    # Count lines and estimate character density
    lines = text_content.split('\n')
    actual_lines = len(lines)
    max_line_length = max(len(line) for line in lines) if lines else 1
    
    # Use more generous font size calculations for better readability
    # Target: readable text that fits well in the available space
    
    # Base font size on text box area - larger areas get larger fonts
    area_pts = width_pts * height_pts
    
    if area_pts > 20000:  # Large text boxes
        base_font_size = 16
    elif area_pts > 10000:  # Medium text boxes  
        base_font_size = 14
    elif area_pts > 5000:   # Small text boxes
        base_font_size = 12
    else:                   # Very small text boxes
        base_font_size = 10
    
    # Adjust based on text density
    if actual_lines > 10:
        base_font_size = max(10, base_font_size - 2)
    elif actual_lines > 5:
        base_font_size = max(11, base_font_size - 1)
    
    if max_line_length > 80:
        base_font_size = max(10, base_font_size - 2)
    elif max_line_length > 50:
        base_font_size = max(11, base_font_size - 1)
    
    # Ensure minimum readability
    font_size = max(10, min(20, base_font_size))  # 10pt to 20pt range for better readability
    
    # Convert to half-points
    font_size_half_points = font_size * 2
    
    return font_size_half_points


def optimize_text_layout(text_blocks: List[Tuple[int, int, int, int, str]]) -> List[Tuple[int, int, int, int, str]]:
    """
    Optimize text block layout by removing overlaps and adjusting positions.
    
    Args:
        text_blocks: List of text blocks in EMU coordinates
        
    Returns:
        Optimized text blocks
    """
    if len(text_blocks) < 2:
        return text_blocks
    
    # Sort blocks by y-coordinate for processing
    sorted_blocks = sorted(text_blocks, key=lambda b: (b[1], b[0]))  # y, then x
    
    optimized_blocks = []
    
    for i, current_block in enumerate(sorted_blocks):
        x0, y0, x1, y1, text = current_block
        
        # Check for overlaps with previous blocks
        adjusted_block = current_block
        
        for prev_block in optimized_blocks:
            px0, py0, px1, py1, _ = prev_block
            
            # Check for overlap
            if _blocks_overlap_emu(adjusted_block, prev_block):
                # Adjust position to avoid overlap
                adjusted_block = _resolve_overlap_emu(adjusted_block, prev_block)
        
        optimized_blocks.append(adjusted_block)
    
    return optimized_blocks


def _blocks_overlap_emu(block1: Tuple[int, int, int, int, str], 
                       block2: Tuple[int, int, int, int, str]) -> bool:
    """
    Check if two text blocks in EMU coordinates overlap.
    
    Args:
        block1, block2: Text blocks in EMU coordinates
        
    Returns:
        True if blocks overlap
    """
    x0_1, y0_1, x1_1, y1_1, _ = block1
    x0_2, y0_2, x1_2, y1_2, _ = block2
    
    # Check for no overlap conditions
    if x1_1 <= x0_2 or x1_2 <= x0_1:  # No horizontal overlap
        return False
    if y1_1 <= y0_2 or y1_2 <= y0_1:  # No vertical overlap
        return False
    
    return True


def _resolve_overlap_emu(current_block: Tuple[int, int, int, int, str],
                        existing_block: Tuple[int, int, int, int, str]) -> Tuple[int, int, int, int, str]:
    """
    Resolve overlap by adjusting the current block position.
    
    Args:
        current_block: Block to adjust
        existing_block: Fixed block that's already positioned
        
    Returns:
        Adjusted current block
    """
    x0, y0, x1, y1, text = current_block
    px0, py0, px1, py1, _ = existing_block
    
    width = x1 - x0
    height = y1 - y0
    
    # Try to move block to avoid overlap
    # First try moving down
    new_y0 = py1 + 18288  # Small gap (0.02 inches in EMU)
    new_y1 = new_y0 + height
    
    # If that doesn't work, try moving right
    if new_y1 > 6858000:  # Approximate slide height limit
        new_x0 = px1 + 18288
        new_x1 = new_x0 + width
        new_y0 = y0  # Keep original y position
        new_y1 = y1
        
        # If still doesn't fit, keep original position (better than invisible)
        if new_x1 > 12192000:  # Approximate slide width limit
            return current_block
        
        return (new_x0, new_y0, new_x1, new_y1, text)
    
    return (x0, new_y0, x1, new_y1, text)


def ensure_minimum_dimensions(text_blocks: List[Tuple[int, int, int, int, str]]) -> List[Tuple[int, int, int, int, str]]:
    """
    Ensure all text blocks meet minimum dimension requirements.
    
    Args:
        text_blocks: List of text blocks in EMU coordinates
        
    Returns:
        Text blocks with enforced minimum dimensions
    """
    min_width_emu = 45720   # 0.05 inches
    min_height_emu = 18288  # 0.02 inches
    
    adjusted_blocks = []
    
    for x0, y0, x1, y1, text in text_blocks:
        width = x1 - x0
        height = y1 - y0
        
        # Adjust width if too small
        if width < min_width_emu:
            x1 = x0 + min_width_emu
        
        # Adjust height if too small
        if height < min_height_emu:
            y1 = y0 + min_height_emu
        
        adjusted_blocks.append((x0, y0, x1, y1, text))
    
    return adjusted_blocks
