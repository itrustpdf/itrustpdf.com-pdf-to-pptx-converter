# PDF to PowerPoint Converter 📊

A powerful web service that converts PDF documents to PowerPoint presentations with natural formatting, built with FastAPI and Docker.

## ✨ Features

- 🔄 **Natural PowerPoint Format**: Creates proper PowerPoint slides instead of preserving PDF layout
- 🌐 **Web Interface**: Easy-to-use drag-and-drop upload interface
- 🐳 **Docker Ready**: Fully containerized with Docker Compose
- 📝 **OCR Support**: Handles scanned PDFs with Tesseract OCR
- ⚡ **Fast Processing**: Efficient text extraction with PyMuPDF
- 🎨 **Smart Formatting**: Automatic title detection and content structuring

## 🚀 Quick Start

### Prerequisites
- Docker and Docker Compose
- Git

### Installation

1. **Clone the repository:**
```bash
git clone https://github.com/illfindyouagain/natural-pdf-pptx-converter.git
cd natural-pdf-pptx-converter
```

2. **Start the service:**
```bash
docker-compose up -d
```

3. **Open your browser:**
   Navigate to `http://localhost:8080`

4. **Convert PDFs:**
   - Drag and drop a PDF file
   - Download the generated PowerPoint presentation

## 🛠️ Technology Stack

- **Backend**: FastAPI (Python 3.12)
- **PDF Processing**: PyMuPDF (fitz)
- **OCR**: Tesseract 5.5.0
- **PowerPoint Generation**: python-pptx
- **Containerization**: Docker & Docker Compose
- **Web Interface**: HTML5 with JavaScript

## 📋 API Endpoints

- `GET /` - Web interface
- `POST /convert/` - Convert PDF to PPTX
- `GET /health` - Service health check

## 🎯 Key Improvements

This converter focuses on creating **natural PowerPoint content** rather than preserving exact PDF layout:

- ✅ Readable font sizes (14-24pt)
- ✅ Standard PowerPoint layouts
- ✅ Intelligent text grouping
- ✅ Proper paragraph spacing
- ✅ Title and content detection

## 🔧 Development

### Local Development
```bash
# Install dependencies
pip install -r requirements.txt

# Run locally
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

### Testing
```bash
# Run tests
python -m pytest tests/

# Test the API
curl -X POST -F "file=@test.pdf" http://localhost:8080/convert/ -o output.pptx
```

## 📁 Project Structure

```
natural-pdf-pptx-converter/
├── app/
│   ├── main.py              # FastAPI application
│   ├── converter.py         # Main conversion logic
│   ├── text_extraction.py   # PDF text extraction
│   ├── pptx_generator.py    # PowerPoint generation
│   ├── models.py           # Data models
│   └── utils.py            # Utility functions
├── tests/                   # Unit tests
├── docker-compose.yml       # Docker configuration
├── Dockerfile              # Container definition
└── requirements.txt        # Python dependencies
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

## 🙏 Acknowledgments

- PyMuPDF for excellent PDF processing
- python-pptx for PowerPoint generation
- Tesseract OCR for text recognition
- FastAPI for the web framework
