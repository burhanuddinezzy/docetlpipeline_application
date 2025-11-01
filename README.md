# PDF Template & Data Extraction System

An advanced Python application for creating templates from PDF documents and extracting structured data using those templates. The system provides an interactive GUI for template creation and a robust extraction pipeline for processing multiple documents.

## Key Features

### Template Creator
- Interactive GUI for creating PDF extraction templates
- Multiple box types for different content:
  - General boxes for basic text extraction
  - Table boxes with automatic cell detection
  - Paragraph boxes with text flow preservation
- Multi-page template support
- Visual table grid detection with adjustable sensitivity
- Box editing and deletion capabilities
- Save/load template configurations
- Support for unboxed content extraction

### Template Extractor
- Automatic template matching for documents
- Fuzzy matching with configurable confidence threshold
- Support for different extraction strategies:
  - Boxed content extraction
  - Table structure preservation
  - Paragraph flow detection
  - Reading order preservation
- Output in Markdown format
- OCR error correction
- Text normalization

## Installation

1. Clone the repository:
```sh
git clone [repository-url]

2. Create a virtual environment (recommended):
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

3. Install dependencies:
pip install -r requirements.txt

Usage
Template Creation
1. Run the template creator:
python template_creator.py

2. Use the GUI to:
Load PDF documents
Navigate between pages
Draw extraction boxes
Configure table detection
Save templates

Data Extraction
1. Run the extraction pipeline:
python main.py

The system will:

Process PDFs from new_docs folder
Apply matching templates
Generate Markdown output in processed_docs
Create CSV summaries in output

Project Structure
├── config.json               # Configuration file
├── main.py                  # Main extraction pipeline
├── template_creator.py      # Template creation GUI
├── template_extractor.py    # Template extraction engine
├── requirements.txt         # Python dependencies
├── new_docs/               # Input PDFs
├── processed_docs/         # Extracted markdown files
├── templates/              # Saved templates
└── output/                # CSV output files

Configuration
Edit config.json to configure:
{
  "new_docs_folder": "new_docs",
  "processed_docs_folder": "processed_docs",
  "output_file": "output/results.csv"
}
Requirements
Python 3.7+
PyMuPDF
OpenCV
NumPy
PIL
pdfplumber
tkinter
