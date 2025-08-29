# Requirements Document

## Introduction

This feature implements a web service that converts PDF documents to text-only PowerPoint presentations (PPTX). The service extracts text from PDFs using native text extraction and OCR fallback for scanned documents, then creates slides with 1:1 page-to-slide mapping while preserving text positioning and readability. The service is containerized using Docker and provides a REST API endpoint for file upload and conversion.

## Requirements

### Requirement 1

**User Story:** As a user, I want to upload a PDF file through a web API and receive a PPTX file, so that I can convert document content to presentation format.

#### Acceptance Criteria

1. WHEN a user sends a POST request to /convert with a PDF file THEN the system SHALL return a PPTX file as a downloadable response
2. WHEN a user uploads a non-PDF file THEN the system SHALL return a 400 error with appropriate message
3. WHEN the conversion process fails THEN the system SHALL return a 500 error with error details
4. WHEN a user accesses the root endpoint THEN the system SHALL return usage instructions

### Requirement 2

**User Story:** As a user, I want each PDF page to become exactly one PowerPoint slide, so that the document structure is preserved in the presentation.

#### Acceptance Criteria

1. WHEN a PDF has N pages THEN the resulting PPTX SHALL have exactly N slides
2. WHEN converting a PDF THEN each slide SHALL correspond to the same-numbered PDF page
3. WHEN a PDF page is empty THEN the corresponding slide SHALL be created but remain blank

### Requirement 3

**User Story:** As a user, I want text from PDF pages to be extracted and positioned appropriately on slides, so that the content remains readable and maintains its layout.

#### Acceptance Criteria

1. WHEN a PDF page contains native text THEN the system SHALL extract text using PyMuPDF
2. WHEN text blocks are extracted THEN they SHALL be positioned on slides at scaled coordinates matching their PDF positions
3. WHEN text is placed on slides THEN font size SHALL be automatically determined based on text box dimensions
4. WHEN text contains line breaks THEN paragraph structure SHALL be preserved

### Requirement 4

**User Story:** As a user, I want scanned or image-heavy PDF pages to be processed with OCR, so that text within images is extracted and included in the presentation.

#### Acceptance Criteria

1. WHEN a PDF page has insufficient native text (less than 20 characters) THEN the system SHALL use Tesseract OCR with English language
2. WHEN performing OCR THEN the system SHALL render the page at 300 DPI for optimal text recognition
3. WHEN OCR extracts text THEN line-level grouping SHALL be used to maintain readability
4. WHEN OCR processing fails THEN the system SHALL create a blank slide for that page

### Requirement 5

**User Story:** As a user, I want images to be excluded from the final presentation, so that I receive a clean text-only PPTX file.

#### Acceptance Criteria

1. WHEN converting any PDF THEN the system SHALL NOT include any images in the resulting PPTX
2. WHEN a PDF page contains only images with no text THEN OCR SHALL be applied to extract any text from those images
3. WHEN the final PPTX is generated THEN it SHALL contain only text content and no visual elements

### Requirement 6

**User Story:** As a user, I want the service to handle text normalization and hyphenation, so that the extracted text is clean and readable.

#### Acceptance Criteria

1. WHEN text contains end-of-line hyphenation THEN the system SHALL merge hyphenated words across line breaks
2. WHEN text blocks are processed THEN excessive whitespace SHALL be normalized
3. WHEN text is grouped THEN reading order SHALL be preserved (top-to-bottom, left-to-right)
4. WHEN text blocks are empty or whitespace-only THEN they SHALL be excluded from the output

### Requirement 7

**User Story:** As a user, I want the service to be containerized and easily deployable, so that I can run it in any Docker environment.

#### Acceptance Criteria

1. WHEN the service is built THEN it SHALL create a single Docker container with all dependencies
2. WHEN the container starts THEN it SHALL expose the web service on port 8000
3. WHEN the container is deployed THEN it SHALL include Tesseract OCR with English language support
4. WHEN using docker-compose THEN the service SHALL be accessible at localhost:8000

### Requirement 8

**User Story:** As a user, I want slide dimensions to be optimized for the PDF content, so that text fits naturally and remains legible.

#### Acceptance Criteria

1. WHEN creating slides THEN the system SHALL use widescreen format as the base
2. WHEN a PDF has a specific aspect ratio THEN slide height SHALL be adjusted to better match the PDF ratio
3. WHEN positioning text boxes THEN a small margin SHALL be applied to prevent clipping
4. WHEN text boxes are created THEN minimum dimensions SHALL be enforced to ensure visibility