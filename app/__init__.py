"""
PDF to PPTX Converter Package

A containerized web service that converts PDF documents to text-only PowerPoint presentations.
"""

__version__ = "1.0.0"
__author__ = "PDF to PPTX Converter"
__description__ = "Convert PDF documents to text-only PPTX presentations"

from .converter import pdf_to_pptx, validate_pdf, get_pdf_info
from .models import TextBlock, PageDimensions, SlideConfig

__all__ = [
    'pdf_to_pptx',
    'validate_pdf', 
    'get_pdf_info',
    'TextBlock',
    'PageDimensions', 
    'SlideConfig'
]
