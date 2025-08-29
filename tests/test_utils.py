"""
Unit tests for the utils module.
"""

from app.utils import (
    get_pdf_dimensions, pixels_to_pdf_points, normalize_coordinates,
    calculate_aspect_ratio, scale_coordinates, apply_margin
)


class TestCoordinateUtilities:
    """Test coordinate utility functions."""
    
    def test_pixels_to_pdf_points(self):
        """Test pixel to PDF points conversion."""
        # At 72 DPI, 1 pixel = 1 point
        x_pts, y_pts = pixels_to_pdf_points(72, 36, 72, 100, 200)
        assert x_pts == 72.0
        assert y_pts == 36.0
        
        # At 300 DPI
        x_pts, y_pts = pixels_to_pdf_points(300, 150, 300, 100, 200)
        assert x_pts == 72.0  # 300/300 * 72
        assert y_pts == 36.0  # 150/300 * 72
    
    def test_normalize_coordinates(self):
        """Test coordinate normalization."""
        # Already normalized
        result = normalize_coordinates(10, 20, 100, 80)
        assert result == (10, 20, 100, 80)
        
        # Reversed x coordinates
        result = normalize_coordinates(100, 20, 10, 80)
        assert result == (10, 20, 100, 80)
        
        # Reversed y coordinates
        result = normalize_coordinates(10, 80, 100, 20)
        assert result == (10, 20, 100, 80)
        
        # Both reversed
        result = normalize_coordinates(100, 80, 10, 20)
        assert result == (10, 20, 100, 80)
    
    def test_calculate_aspect_ratio(self):
        """Test aspect ratio calculation."""
        # Normal case
        ratio = calculate_aspect_ratio(16, 9)
        assert abs(ratio - 16/9) < 0.001
        
        # Square
        ratio = calculate_aspect_ratio(100, 100)
        assert ratio == 1.0
        
        # Zero height (edge case)
        ratio = calculate_aspect_ratio(100, 0)
        assert ratio == 1.0
    
    def test_scale_coordinates(self):
        """Test coordinate scaling."""
        x, y = scale_coordinates(10, 20, 2.0, 1.5)
        assert x == 20.0
        assert y == 30.0
    
    def test_apply_margin(self):
        """Test margin application."""
        # Normal case
        result = apply_margin(10, 20, 90, 70, 5, 5, 100, 100)
        x0, y0, x1, y1 = result
        assert x0 == 15  # 10 + 5
        assert y0 == 25  # 20 + 5
        assert x1 == 85  # 90 - 5
        assert y1 == 65  # 70 - 5
        
        # Margin causes negative dimensions - should enforce minimum
        result = apply_margin(40, 45, 50, 55, 10, 10, 100, 100)
        x0, y0, x1, y1 = result
        assert x1 > x0  # Should maintain minimum width
        assert y1 > y0  # Should maintain minimum height


class TestAspectRatioCalculation:
    """Test aspect ratio calculation in various scenarios."""
    
    def test_standard_ratios(self):
        """Test standard aspect ratios."""
        # 16:9 widescreen
        ratio = calculate_aspect_ratio(1920, 1080)
        assert abs(ratio - 16/9) < 0.01
        
        # 4:3 traditional
        ratio = calculate_aspect_ratio(1024, 768)
        assert abs(ratio - 4/3) < 0.01
    
    def test_edge_cases(self):
        """Test edge cases for aspect ratio."""
        # Very wide
        ratio = calculate_aspect_ratio(1000, 1)
        assert ratio == 1000.0
        
        # Very tall
        ratio = calculate_aspect_ratio(1, 1000)
        assert ratio == 0.001
