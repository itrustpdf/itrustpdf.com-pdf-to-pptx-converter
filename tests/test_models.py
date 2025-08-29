"""
Unit tests for the models module.
"""

import pytest
from app.models import (
    TextBlock, PageDimensions, SlideConfig,
    validate_text_block, calculate_text_area,
    pdf_points_to_emu, emu_to_pdf_points,
    MINIMUM_TEXT_THRESHOLD, DEFAULT_OCR_DPI
)


class TestTextBlockValidation:
    """Test text block validation functions."""
    
    def test_valid_text_block(self):
        """Test validation of a valid text block."""
        block = (10.0, 20.0, 100.0, 50.0, "Sample text")
        assert validate_text_block(block) is True
    
    def test_invalid_coordinates(self):
        """Test validation with invalid coordinates."""
        # x1 <= x0
        block = (100.0, 20.0, 10.0, 50.0, "Sample text")
        assert validate_text_block(block) is False
        
        # y1 <= y0
        block = (10.0, 50.0, 100.0, 20.0, "Sample text")
        assert validate_text_block(block) is False
    
    def test_empty_text(self):
        """Test validation with empty text."""
        block = (10.0, 20.0, 100.0, 50.0, "")
        assert validate_text_block(block) is False
        
        block = (10.0, 20.0, 100.0, 50.0, "   ")
        assert validate_text_block(block) is False
    
    def test_wrong_tuple_length(self):
        """Test validation with wrong tuple length."""
        block = (10.0, 20.0, 100.0, "Sample text")  # Missing y1
        assert validate_text_block(block) is False


class TestTextBlockArea:
    """Test text block area calculation."""
    
    def test_calculate_area(self):
        """Test area calculation."""
        block = (0.0, 0.0, 10.0, 5.0, "text")
        area = calculate_text_area(block)
        assert area == 50.0
    
    def test_zero_area(self):
        """Test area calculation with zero dimensions."""
        block = (0.0, 0.0, 0.0, 5.0, "text")
        area = calculate_text_area(block)
        assert area == 0.0


class TestCoordinateConversion:
    """Test coordinate conversion functions."""
    
    def test_pdf_points_to_emu(self):
        """Test PDF points to EMU conversion."""
        # 72 points = 1 inch = 914400 EMU
        result = pdf_points_to_emu(72.0)
        assert result == 914400
    
    def test_emu_to_pdf_points(self):
        """Test EMU to PDF points conversion."""
        # 914400 EMU = 1 inch = 72 points
        result = emu_to_pdf_points(914400)
        assert result == 72.0
    
    def test_round_trip_conversion(self):
        """Test round-trip conversion accuracy."""
        original_points = 100.0
        emu = pdf_points_to_emu(original_points)
        converted_back = emu_to_pdf_points(emu)
        assert abs(converted_back - original_points) < 0.01


class TestSlideConfig:
    """Test SlideConfig class."""
    
    def test_slide_config_creation(self):
        """Test basic SlideConfig creation."""
        config = SlideConfig(1000000, 750000)
        assert config.width_emu == 1000000
        assert config.height_emu == 750000
        assert config.margin_factor == 0.02  # Default
    
    def test_slide_config_with_custom_margin(self):
        """Test SlideConfig with custom margin."""
        config = SlideConfig(1000000, 750000, 0.05)
        assert config.margin_factor == 0.05
    
    def test_width_pts_property(self):
        """Test width_pts property conversion."""
        config = SlideConfig(914400, 750000)  # 1 inch width
        assert abs(config.width_pts - 72.0) < 0.01
    
    def test_height_pts_property(self):
        """Test height_pts property conversion."""
        config = SlideConfig(1000000, 914400)  # 1 inch height
        assert abs(config.height_pts - 72.0) < 0.01


class TestConstants:
    """Test module constants."""
    
    def test_constants_exist(self):
        """Test that required constants are defined."""
        assert MINIMUM_TEXT_THRESHOLD == 20
        assert DEFAULT_OCR_DPI == 300
    
    def test_constant_types(self):
        """Test that constants have correct types."""
        assert isinstance(MINIMUM_TEXT_THRESHOLD, int)
        assert isinstance(DEFAULT_OCR_DPI, int)
