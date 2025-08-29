"""
Integration tests for the FastAPI web service.
"""

from fastapi.testclient import TestClient
import io
from app.main import app

client = TestClient(app)


class TestAPIEndpoints:
    """Test FastAPI endpoints."""
    
    def test_root_endpoint(self):
        """Test the root endpoint returns usage instructions."""
        response = client.get("/")
        assert response.status_code == 200
        assert "PDF to PPTX Converter" in response.text
        assert "text/html" in response.headers["content-type"]
    
    def test_health_endpoint(self):
        """Test the health check endpoint."""
        response = client.get("/health")
        assert response.status_code == 200
        
        data = response.json()
        assert "status" in data
        assert "service" in data
        assert "dependencies" in data
        assert data["service"] == "PDF to PPTX Converter"
    
    def test_convert_without_file(self):
        """Test convert endpoint without file."""
        response = client.post("/convert")
        assert response.status_code == 422  # Validation error
    
    def test_convert_with_non_pdf(self):
        """Test convert endpoint with non-PDF file."""
        # Create a fake text file
        fake_file = io.BytesIO(b"This is not a PDF")
        
        response = client.post(
            "/convert",
            files={"file": ("test.txt", fake_file, "text/plain")}
        )
        assert response.status_code == 400
        assert "Please upload a PDF file" in response.json()["detail"]
    
    def test_info_without_file(self):
        """Test info endpoint without file."""
        response = client.post("/info")
        assert response.status_code == 422  # Validation error
    
    def test_info_with_non_pdf(self):
        """Test info endpoint with non-PDF file."""
        fake_file = io.BytesIO(b"This is not a PDF")
        
        response = client.post(
            "/info",
            files={"file": ("test.txt", fake_file, "text/plain")}
        )
        assert response.status_code == 400
        assert "Please upload a PDF file" in response.json()["detail"]
    
    def test_404_endpoint(self):
        """Test non-existent endpoint."""
        response = client.get("/nonexistent")
        assert response.status_code == 404
        assert "404 - Page Not Found" in response.text


class TestAPIErrorHandling:
    """Test API error handling."""
    
    def test_empty_file_upload(self):
        """Test handling of empty file upload."""
        empty_file = io.BytesIO(b"")
        
        response = client.post(
            "/convert",
            files={"file": ("empty.pdf", empty_file, "application/pdf")}
        )
        assert response.status_code == 400
        assert "Empty file uploaded" in response.json()["detail"]
    
    def test_malformed_pdf(self):
        """Test handling of malformed PDF."""
        malformed_pdf = io.BytesIO(b"Not a real PDF content")
        
        response = client.post(
            "/convert",
            files={"file": ("malformed.pdf", malformed_pdf, "application/pdf")}
        )
        assert response.status_code == 400
        assert "Invalid PDF file" in response.json()["detail"]
