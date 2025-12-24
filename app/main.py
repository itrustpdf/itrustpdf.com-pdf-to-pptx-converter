"""
FastAPI web service for PDF to PPTX and PPTX to PDF conversion.
"""

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
import io
import logging
from typing import Optional

from .converter import (
    pdf_to_pptx, 
    pptx_to_pdf,
    validate_pdf,
    validate_pptx,
    get_pdf_info,
    get_pptx_info,
    estimate_processing_time,
    estimate_pptx_processing_time
)
from .ocr import test_tesseract_installation, get_tesseract_version

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Create FastAPI app
app = FastAPI(
    title="PDF to PPTX and PPTX to PDF Converter",
    description="A web service that converts PDF documents to PowerPoint presentations and vice versa",
    version="2.0.0",
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
    logger.info("Starting PDF/PPTX Converter service")
    
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
        <title>PDF ⇄ PPTX Converter</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f5f5; }
            h1 { color: #333; text-align: center; }
            .container { max-width: 1200px; margin: 0 auto; }
            .upload-form { background: white; padding: 30px; margin: 20px 0; border-radius: 12px; border: 2px solid #007acc; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
            .upload-form h2 { color: #007acc; margin-top: 0; display: flex; align-items: center; gap: 10px; }
            .form-group { margin: 20px 0; }
            label { display: block; margin-bottom: 8px; font-weight: bold; color: #555; }
            input[type="file"] { width: 100%; padding: 12px; border: 2px dashed #007acc; border-radius: 8px; background: #f9f9f9; cursor: pointer; }
            input[type="checkbox"] { margin-right: 10px; transform: scale(1.2); }
            .btn { background: linear-gradient(135deg, #007acc, #005a99); color: white; padding: 15px 30px; border: none; border-radius: 8px; cursor: pointer; font-size: 16px; font-weight: bold; width: 100%; transition: all 0.3s; display: flex; align-items: center; justify-content: center; gap: 10px; }
            .btn:hover { background: linear-gradient(135deg, #005a99, #003d66); transform: translateY(-2px); box-shadow: 0 6px 12px rgba(0,0,0,0.15); }
            .btn:disabled { background: #ccc; cursor: not-allowed; transform: none; box-shadow: none; }
            .endpoint { background: white; padding: 20px; margin: 15px 0; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.05); border-left: 4px solid #007acc; }
            .method { color: #fff; background: #007acc; padding: 5px 12px; border-radius: 4px; font-weight: bold; font-size: 14px; }
            pre { background: #f8f8f8; padding: 15px; border-radius: 6px; overflow-x: auto; font-size: 14px; border: 1px solid #eee; }
            .info { background: #e7f3ff; padding: 20px; border-radius: 8px; border-left: 5px solid #007acc; margin: 20px 0; }
            .progress { display: none; margin: 20px 0; padding: 15px; background: #f0f8ff; border-radius: 8px; border: 1px solid #007acc; color: #007acc; font-weight: bold; text-align: center; }
            .converter-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 30px; margin: 30px 0; }
            .converter-arrow { display: flex; align-items: center; justify-content: center; font-size: 40px; color: #007acc; }
            @media (max-width: 768px) {
                .converter-grid { grid-template-columns: 1fr; }
                .converter-arrow { display: none; }
            }
            .tab-container { background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
            .tabs { display: flex; background: #f0f0f0; }
            .tab { flex: 1; padding: 15px; text-align: center; cursor: pointer; font-weight: bold; color: #666; transition: all 0.3s; }
            .tab.active { background: white; color: #007acc; border-bottom: 3px solid #007acc; }
            .tab-content { display: none; padding: 30px; }
            .tab-content.active { display: block; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📄 PDF ⇄ PPTX Converter</h1>
            
            <div class="info">
                <p><strong>Welcome to the PDF ⇄ PPTX Converter!</strong></p>
                <p>Convert PDF documents to PowerPoint presentations and PowerPoint presentations back to PDF with 1:1 page/slide mapping.</p>
            </div>
            
            <div class="tab-container">
                <div class="tabs">
                    <div class="tab active" onclick="switchTab('pdf-to-pptx')">📄 PDF to PPTX</div>
                    <div class="tab" onclick="switchTab('pptx-to-pdf')">🔄 PPTX to PDF</div>
                </div>
                
                <div id="pdf-to-pptx" class="tab-content active">
                    <div class="upload-form">
                        <h2>📄 → 📊 Convert PDF to PPTX</h2>
                        <form id="pdfUploadForm" enctype="multipart/form-data">
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
                            
                            <button type="submit" class="btn">
                                <span>🔄 Convert PDF to PPTX</span>
                            </button>
                            <div class="progress" id="pdfProgress">Converting... Please wait.</div>
                        </form>
                    </div>
                </div>
                
                <div id="pptx-to-pdf" class="tab-content">
                    <div class="upload-form">
                        <h2>📊 → 📄 Convert PPTX to PDF</h2>
                        <form id="pptxUploadForm" enctype="multipart/form-data">
                            <div class="form-group">
                                <label for="pptxFile">Select PPTX File:</label>
                                <input type="file" id="pptxFile" name="file" accept=".pptx,.pptm,.ppt" required>
                            </div>
                            
                            <button type="submit" class="btn">
                                <span>🔄 Convert PPTX to PDF</span>
                            </button>
                            <div class="progress" id="pptxProgress">Converting... Please wait.</div>
                        </form>
                    </div>
                </div>
            </div>
            
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
                <p><span class="method">POST</span> <strong>/convert-pptx</strong></p>
                <p>Upload a PPTX file and receive a PDF file in response.</p>
                <p><strong>Parameters:</strong></p>
                <ul>
                    <li><code>file</code> (required): PPTX file to convert</li>
                </ul>
            </div>
            
            <div class="endpoint">
                <p><span class="method">POST</span> <strong>/info</strong></p>
                <p>Get information about a PDF file without converting it.</p>
            </div>
            
            <div class="endpoint">
                <p><span class="method">POST</span> <strong>/info-pptx</strong></p>
                <p>Get information about a PPTX file without converting it.</p>
            </div>
            
            <div class="endpoint">
                <p><span class="method">GET</span> <strong>/health</strong></p>
                <p>Check service health and dependencies status.</p>
            </div>
            
            <h2>Example Usage</h2>
            <pre>
# Convert PDF to PPTX
curl -X POST "http://localhost:8000/convert" \\
     -F "file=@document.pdf" \\
     --output "presentation.pptx"

# Convert PPTX to PDF
curl -X POST "http://localhost:8000/convert-pptx" \\
     -F "file=@presentation.pptx" \\
     --output "document.pdf"

# Get PDF information
curl -X POST "http://localhost:8000/info" \\
     -F "file=@document.pdf"

# Get PPTX information
curl -X POST "http://localhost:8000/info-pptx" \\
     -F "file=@presentation.pptx"
            </pre>
            
            <h2>Features</h2>
            <ul>
                <li>📄 → 📊 PDF to PPTX: Extracts text from native PDF content with OCR fallback</li>
                <li>📊 → 📄 PPTX to PDF: Preserves text layout, formatting, and positioning 1:1</li>
                <li>Preserves text positioning and layout in both directions</li>
                <li>Automatic font sizing and slide/page dimension optimization</li>
                <li>Support for multiple OCR languages for scanned PDFs</li>
            </ul>
            
            <p><a href="/docs">View API Documentation</a> | <a href="/health">Check Health</a></p>
        </div>
        
        <script>
        function switchTab(tabId) {
            // Hide all tab contents
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });
            
            // Remove active class from all tabs
            document.querySelectorAll('.tab').forEach(tab => {
                tab.classList.remove('active');
            });
            
            // Show selected tab content
            document.getElementById(tabId).classList.add('active');
            
            // Activate clicked tab
            event.target.classList.add('active');
        }
        
        // PDF to PPTX form handler
        document.getElementById('pdfUploadForm').onsubmit = async function(e) {
            e.preventDefault();
            
            const fileInput = document.getElementById('pdfFile');
            const progressDiv = document.getElementById('pdfProgress');
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
            submitBtn.innerHTML = '<span>Converting... Please wait</span>';
            progressDiv.style.display = 'block';
            progressDiv.textContent = 'Converting PDF to PPTX... This may take a moment.';
            progressDiv.style.color = '#007acc';
            
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
                    a.download = fileInput.files[0].name.replace(/\.pdf$/i, '_converted.pptx');
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    progressDiv.textContent = '✓ Conversion completed! Download started.';
                    progressDiv.style.color = '#28a745';
                } else {
                    const error = await response.text();
                    throw new Error(`Conversion failed: ${error}`);
                }
            } catch (error) {
                progressDiv.textContent = `✗ Error: ${error.message}`;
                progressDiv.style.color = '#dc3545';
            } finally {
                submitBtn.disabled = false;
                submitBtn.innerHTML = '<span>🔄 Convert PDF to PPTX</span>';
                setTimeout(() => {
                    progressDiv.style.display = 'none';
                }, 5000);
            }
        };
        
        // PPTX to PDF form handler
        document.getElementById('pptxUploadForm').onsubmit = async function(e) {
            e.preventDefault();
            
            const fileInput = document.getElementById('pptxFile');
            const progressDiv = document.getElementById('pptxProgress');
            const submitBtn = e.target.querySelector('button[type="submit"]');
            
            if (!fileInput.files[0]) {
                alert('Please select a PPTX file');
                return;
            }
            
            const formData = new FormData();
            formData.append('file', fileInput.files[0]);
            
            submitBtn.disabled = true;
            submitBtn.innerHTML = '<span>Converting... Please wait</span>';
            progressDiv.style.display = 'block';
            progressDiv.textContent = 'Converting PPTX to PDF... This may take a moment.';
            progressDiv.style.color = '#007acc';
            
            try {
                const response = await fetch('/convert-pptx', {
                    method: 'POST',
                    body: formData
                });
                
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.style.display = 'none';
                    a.href = url;
                    a.download = fileInput.files[0].name.replace(/\.pptx$/i, '_converted.pdf');
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    progressDiv.textContent = '✓ Conversion completed! Download started.';
                    progressDiv.style.color = '#28a745';
                } else {
                    const error = await response.text();
                    throw new Error(`Conversion failed: ${error}`);
                }
            } catch (error) {
                progressDiv.textContent = `✗ Error: ${error.message}`;
                progressDiv.style.color = '#dc3545';
            } finally {
                submitBtn.disabled = false;
                submitBtn.innerHTML = '<span>🔄 Convert PPTX to PDF</span>';
                setTimeout(() => {
                    progressDiv.style.display = 'none';
                }, 5000);
            }
        };
        </script>
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
        logger.info(f"Processing PDF upload: {file.filename}")
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
        logger.error(f"PDF to PPTX conversion failed: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Conversion failed: {str(e)}"
        )


@app.post("/convert-pptx")
async def convert_pptx_to_pdf(
    file: UploadFile = File(...)
):
    """
    Convert a PPTX file to PDF format.
    
    Args:
        file: PPTX file to convert
        
    Returns:
        StreamingResponse with PDF file
        
    Raises:
        HTTPException: If file validation or conversion fails
    """
    # Validate file type
    valid_extensions = ['.pptx', '.pptm', '.ppt']
    if not file.filename or not any(file.filename.lower().endswith(ext) for ext in valid_extensions):
        raise HTTPException(
            status_code=400,
            detail="Please upload a PPTX, PPTM, or PPT file"
        )
    
    try:
        # Read file content
        logger.info(f"Processing PPTX upload: {file.filename}")
        pptx_content = await file.read()
        
        if len(pptx_content) == 0:
            raise HTTPException(
                status_code=400,
                detail="Empty file uploaded"
            )
        
        # Validate PPTX
        if not validate_pptx(pptx_content):
            raise HTTPException(
                status_code=400,
                detail="Invalid PPTX file"
            )
        
        # Log processing info
        pptx_info = get_pptx_info(pptx_content)
        estimated_time = estimate_pptx_processing_time(pptx_content)
        logger.info(f"Converting {pptx_info.get('slide_count', 'unknown')} slides, "
                   f"estimated time: {estimated_time:.1f}s")
        
        # Convert PPTX to PDF
        pdf_content = pptx_to_pdf(pptx_content)
        
        # Generate response filename
        base_filename = file.filename.rsplit('.', 1)[0]
        output_filename = f"{base_filename}.pdf"
        
        logger.info(f"Conversion completed: {output_filename} ({len(pdf_content)} bytes)")
        
        return StreamingResponse(
            io.BytesIO(pdf_content),
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename={output_filename}",
                "Content-Length": str(len(pdf_content))
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"PPTX to PDF conversion failed: {str(e)}")
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


@app.post("/info-pptx")
async def get_pptx_information(file: UploadFile = File(...)):
    """
    Get information about a PPTX file without converting it.
    
    Args:
        file: PPTX file to analyze
        
    Returns:
        Dictionary with PPTX information
        
    Raises:
        HTTPException: If file validation fails
    """
    # Validate file type
    valid_extensions = ['.pptx', '.pptm', '.ppt']
    if not file.filename or not any(file.filename.lower().endswith(ext) for ext in valid_extensions):
        raise HTTPException(
            status_code=400,
            detail="Please upload a PPTX, PPTM, or PPT file"
        )
    
    try:
        # Read file content
        pptx_content = await file.read()
        
        if len(pptx_content) == 0:
            raise HTTPException(
                status_code=400,
                detail="Empty file uploaded"
            )
        
        # Validate PPTX
        if not validate_pptx(pptx_content):
            raise HTTPException(
                status_code=400,
                detail="Invalid PPTX file"
            )
        
        # Get PPTX information
        pptx_info = get_pptx_info(pptx_content)
        
        # Add processing estimates
        pptx_info['estimated_processing_time_seconds'] = estimate_pptx_processing_time(pptx_content)
        pptx_info['file_size_bytes'] = len(pptx_content)
        pptx_info['filename'] = file.filename
        
        return pptx_info
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"PPTX info extraction failed: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"Failed to analyze PPTX: {str(e)}"
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
        "service": "PDF ⇄ PPTX Converter",
        "version": "2.0.0",
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
