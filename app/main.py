"""
FastAPI web service for PDF to PPTX conversion.
"""

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import io
import logging
from typing import Optional

from .converter import pdf_to_pptx, validate_pdf, get_pdf_info, estimate_processing_time
from .ocr import test_tesseract_installation, get_tesseract_version

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Create FastAPI app
app = FastAPI(
    title="PDF to PPTX Converter",
    description="A web service that converts PDF documents to text-only PowerPoint presentations",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
async def startup_event():
    """Initialize the service and check dependencies."""
    logger.info("Starting PDF to PPTX Converter service")
    
    # Test Tesseract installation
    if test_tesseract_installation():
        version = get_tesseract_version()
        logger.info(f"Tesseract OCR is available: {version}")
    else:
        logger.warning("Tesseract OCR is not available - OCR functionality will be limited")


@app.get("/", response_class=HTMLResponse)
async def root():
    """Return service usage instructions."""
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>PDF to PPTX Converter</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; }
            h1 { color: #333; }
            .upload-form { background: #f9f9f9; padding: 20px; margin: 20px 0; border-radius: 8px; border: 2px solid #007acc; }
            .form-group { margin: 15px 0; }
            label { display: block; margin-bottom: 5px; font-weight: bold; color: #555; }
            input[type="file"] { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
            input[type="checkbox"] { margin-right: 8px; }
            .btn { background: #007acc; color: white; padding: 12px 24px; border: none; border-radius: 4px; cursor: pointer; font-size: 16px; }
            .btn:hover { background: #005a99; }
            .endpoint { background: #f5f5f5; padding: 15px; margin: 10px 0; border-radius: 5px; }
            .method { color: #fff; background: #007acc; padding: 5px 10px; border-radius: 3px; font-weight: bold; }
            pre { background: #f0f0f0; padding: 10px; border-radius: 3px; overflow-x: auto; }
            .info { background: #e7f3ff; padding: 15px; border-radius: 5px; border-left: 4px solid #007acc; }
            .progress { display: none; margin: 15px 0; color: #007acc; }
        </style>
    </head>
    <body>
        <h1>PDF to PPTX Converter</h1>
        
        <div class="info">
            <p><strong>Welcome to the PDF to PPTX Converter!</strong></p>
            <p>This service converts PDF documents to text-only PowerPoint presentations with 1:1 page-to-slide mapping.</p>
        </div>
        
        <div class="upload-form">
            <h2>📄 Upload PDF File</h2>
            <form id="uploadForm" enctype="multipart/form-data">
                <div class="form-group">
                    <label for="pdfFile">Select PDF File:</label>
                    <input type="file" id="pdfFile" name="file" accept=".pdf" required>
                </div>
                
                <div class="form-group">
                    <label for="ocrLanguages">OCR Languages (for scanned PDFs):</label>
                    <input type="text" id="ocrLanguages" name="ocr_languages" value="eng" placeholder="eng, fra, deu, etc.">
                </div>
                
                <div class="form-group">
                    <label>
                        <input type="checkbox" id="dehyphenate" name="dehyphenate" checked>
                        Remove hyphenation from text
                    </label>
                </div>
                
                <button type="submit" class="btn">🔄 Convert to PPTX</button>
                <div class="progress" id="progress">Converting... Please wait.</div>
            </form>
        </div>
        
        <script>
        document.getElementById('uploadForm').onsubmit = async function(e) {
            e.preventDefault();
            
            const fileInput = document.getElementById('pdfFile');
            const progressDiv = document.getElementById('progress');
            const submitBtn = e.target.querySelector('button[type="submit"]');
            
            if (!fileInput.files[0]) {
                alert('Please select a PDF file');
                return;
            }
            
            const formData = new FormData();
            formData.append('file', fileInput.files[0]);
            formData.append('ocr_languages', document.getElementById('ocrLanguages').value);
            formData.append('dehyphenate', document.getElementById('dehyphenate').checked);
            
            submitBtn.disabled = true;
            submitBtn.textContent = 'Converting...';
            progressDiv.style.display = 'block';
            
            try {
                const response = await fetch('/convert', {
                    method: 'POST',
                    body: formData
                });
                
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.style.display = 'none';
                    a.href = url;
                    a.download = fileInput.files[0].name.replace('.pdf', '_converted.pptx');
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    progressDiv.textContent = 'Conversion completed! Download started.';
                } else {
                    const error = await response.text();
                    throw new Error(`Conversion failed: ${error}`);
                }
            } catch (error) {
                progressDiv.textContent = `Error: ${error.message}`;
                progressDiv.style.color = 'red';
            } finally {
                submitBtn.disabled = false;
                submitBtn.textContent = '🔄 Convert to PPTX';
                setTimeout(() => {
                    progressDiv.style.display = 'none';
                    progressDiv.style.color = '#007acc';
                }, 3000);
            }
        };
        </script>
        
        <h2>Available Endpoints</h2>
        
        <div class="endpoint">
            <p><span class="method">POST</span> <strong>/convert</strong></p>
            <p>Upload a PDF file and receive a PPTX file in response.</p>
            <p><strong>Parameters:</strong></p>
            <ul>
                <li><code>file</code> (required): PDF file to convert</li>
                <li><code>ocr_languages</code> (optional): OCR language codes (default: 'eng')</li>
                <li><code>dehyphenate</code> (optional): Remove hyphenation (default: true)</li>
            </ul>
        </div>
        
        <div class="endpoint">
            <p><span class="method">GET</span> <strong>/health</strong></p>
            <p>Check service health and dependencies status.</p>
        </div>
        
        <div class="endpoint">
            <p><span class="method">POST</span> <strong>/info</strong></p>
            <p>Get information about a PDF file without converting it.</p>
        </div>
        
        <h2>Example Usage</h2>
        <pre>
# Convert PDF to PPTX
curl -X POST "http://localhost:8000/convert" \\
     -F "file=@document.pdf" \\
     --output "presentation.pptx"

# Get PDF information
curl -X POST "http://localhost:8000/info" \\
     -F "file=@document.pdf"
        </pre>
        
        <h2>Features</h2>
        <ul>
            <li>Extracts text from native PDF content</li>
            <li>Falls back to OCR for scanned documents</li>
            <li>Preserves text positioning and layout</li>
            <li>No images included in output</li>
            <li>Optimized slide dimensions</li>
            <li>Automatic font sizing</li>
        </ul>
        
        <p><a href="/docs">View API Documentation</a> | <a href="/health">Check Health</a></p>
    </body>
    </html>
    """
    return html_content


@app.post("/convert")
async def convert_pdf_to_pptx(
    file: UploadFile = File(...),
    ocr_languages: str = "eng",
    dehyphenate: bool = True
):
    """
    Convert a PDF file to PPTX format.
    
    Args:
        file: PDF file to convert
        ocr_languages: Tesseract language codes for OCR (default: 'eng')
        dehyphenate: Whether to remove end-of-line hyphenation (default: True)
        
    Returns:
        StreamingResponse with PPTX file
        
    Raises:
        HTTPException: If file validation or conversion fails
    """
    # Validate file type
    if not file.filename or not file.filename.lower().endswith('.pdf'):
        raise HTTPException(
            status_code=400,
            detail="Please upload a PDF file"
        )
    
    try:
        # Read file content
        logger.info(f"Processing upload: {file.filename}")
        pdf_content = await file.read()
        
        if len(pdf_content) == 0:
            raise HTTPException(
                status_code=400,
                detail="Empty file uploaded"
            )
        
        # Validate PDF
        if not validate_pdf(pdf_content):
            raise HTTPException(
                status_code=400,
                detail="Invalid PDF file"
            )
        
        # Log processing info
        pdf_info = get_pdf_info(pdf_content)
        estimated_time = estimate_processing_time(pdf_content)
        logger.info(f"Converting {pdf_info.get('page_count', 'unknown')} pages, "
                   f"estimated time: {estimated_time:.1f}s")
        
        # Convert PDF to PPTX
        pptx_content = pdf_to_pptx(
            pdf_content, 
            ocr_langs=ocr_languages, 
            dehyphenate=dehyphenate
        )
        
        # Generate response filename
        base_filename = file.filename.rsplit('.', 1)[0]
        output_filename = f"{base_filename}.pptx"
        
        # Create streaming response
        pptx_stream = io.BytesIO(pptx_content)
        
        logger.info(f"Conversion completed: {output_filename} ({len(pptx_content)} bytes)")
        
        return StreamingResponse(
            io.BytesIO(pptx_content),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f"attachment; filename={output_filename}",
                "Content-Length": str(len(pptx_content))
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Conversion failed: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Conversion failed: {str(e)}"
        )


@app.post("/info")
async def get_pdf_information(file: UploadFile = File(...)):
    """
    Get information about a PDF file without converting it.
    
    Args:
        file: PDF file to analyze
        
    Returns:
        Dictionary with PDF information
        
    Raises:
        HTTPException: If file validation fails
    """
    # Validate file type
    if not file.filename or not file.filename.lower().endswith('.pdf'):
        raise HTTPException(
            status_code=400,
            detail="Please upload a PDF file"
        )
    
    try:
        # Read file content
        pdf_content = await file.read()
        
        if len(pdf_content) == 0:
            raise HTTPException(
                status_code=400,
                detail="Empty file uploaded"
            )
        
        # Validate PDF
        if not validate_pdf(pdf_content):
            raise HTTPException(
                status_code=400,
                detail="Invalid PDF file"
            )
        
        # Get PDF information
        pdf_info = get_pdf_info(pdf_content)
        
        # Add processing estimates
        pdf_info['estimated_processing_time_seconds'] = estimate_processing_time(pdf_content)
        pdf_info['file_size_bytes'] = len(pdf_content)
        pdf_info['filename'] = file.filename
        
        return pdf_info
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"PDF info extraction failed: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Failed to analyze PDF: {str(e)}"
        )


@app.get("/health")
async def health_check():
    """
    Check service health and dependencies.
    
    Returns:
        Dictionary with health status
    """
    health_status = {
        "status": "healthy",
        "service": "PDF to PPTX Converter",
        "version": "1.0.0",
        "dependencies": {
            "tesseract_ocr": test_tesseract_installation(),
            "tesseract_version": get_tesseract_version(),
        }
    }
    
    # Check if critical dependencies are available
    if not health_status["dependencies"]["tesseract_ocr"]:
        health_status["status"] = "degraded"
        health_status["warnings"] = ["Tesseract OCR not available - OCR functionality limited"]
    
    return health_status


@app.exception_handler(404)
async def not_found_handler(request, exc):
    """Handle 404 errors with custom message."""
    return HTMLResponse(
        content="""
        <html>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
            <h1>404 - Page Not Found</h1>
            <p>The requested endpoint does not exist.</p>
            <p><a href="/">Return to Home</a> | <a href="/docs">View API Documentation</a></p>
        </body>
        </html>
        """,
        status_code=404
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
