# Scan2Excel AI

Scan2Excel AI is a desktop application that processes scanned A4 production forms and automatically extracts structured table data into Excel.

The application uses computer vision and OCR to detect tables, recognize text and handwritten numbers, and export the result into Excel.

The tool runs fully **offline** and can be built into a **Windows executable (.exe)**.

---

# Main Features

### Document Processing

- Detect A4 document from scanned image
- Perspective correction
- Image preprocessing
- Noise reduction
- Thresholding

### Table Recognition

- Detect table grid lines
- Segment table cells
- Extract rows of structured data

### OCR Recognition

Using PaddleOCR:

- Printed text recognition
- Handwritten number recognition
- Text region detection

### Data Extraction

Extract fields such as:

- Step Number (STT)
- Operation Name
- Standard Time
- Worker
- Quantity

### Data Cleaning

Parse expressions like:

51992 = 3140

Extract the final number:

3140

### Excel Export

Export structured data using:

- pandas
- openpyxl

### Desktop GUI

Built with PySide6.

Features:

- Upload image
- Drag and drop
- Preview image
- Start processing
- Progress bar
- Log panel
- Export Excel

---

# Technology Stack

Language

Python 3.10+

Libraries

OpenCV  
PaddleOCR  
Pandas  
OpenPyXL  
PySide6  
PyInstaller

---

# Project Structure
Scan2Excel-AI/
│
├── main.py
├── gui.py
│
├── vision/
│ ├── preprocess.py
│ ├── document_detect.py
│ ├── table_detection.py
│ └── cell_extraction.py
│
├── ocr/
│ ├── ocr_engine.py
│ └── text_parser.py
│
├── data/
│ ├── data_model.py
│ └── data_cleaner.py
│
├── export/
│ └── excel_export.py
│
├── utils/
│ ├── logger.py
│ └── config.py
│
├── assets/
├── models/
└── output/


---

# Module Overview

### `main.py`

Application entry point.

Responsibilities:

- Initialize application
- Load configuration
- Initialize logging
- Launch the GUI

---

### `gui.py`

Desktop interface implemented with **PySide6**.

Main features:

- Upload image
- Drag & drop image
- Preview scanned form
- Start processing pipeline
- Display progress bar
- Display logs
- Export results to Excel

---

### `vision/`

Computer vision pipeline responsible for image processing and table detection.

**Files**

`preprocess.py`

- image loading
- grayscale conversion
- thresholding
- noise reduction

`document_detect.py`

- detect A4 document
- perspective correction
- warp document image

`table_detection.py`

- detect horizontal lines
- detect vertical lines
- detect table structure

`cell_extraction.py`

- extract table cells
- sort cells into rows

---

### `ocr/`

OCR recognition pipeline.

**Files**

`ocr_engine.py`

- initialize PaddleOCR
- detect text regions
- recognize text

`text_parser.py`

- parse OCR output
- detect numbers
- extract quantities
- interpret expressions

Example:
51992 = 3140

Result: 3140


---

### `data/`

Data structure and data cleaning logic.

`data_model.py`

Defines structured JSON format:

{
"product": "",
"total_quantity": "",
"operations":[
{
"step":1,
"name":"",
"time":"",
"worker":"",
"quantity":""
}
]
}


`data_cleaner.py`

- clean OCR errors
- normalize numbers
- validate extracted rows

---