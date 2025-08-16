# 📄 PDF Comment Extractor (CRS Automation Tool)

A desktop application that automates the creation of **Client Response Sheets (CRS)** from annotated PDF files.  
Instead of manually searching through 400+ page PDFs and copying comments into Excel, this app extracts annotations in one click and fills them into a structured Excel template.

---

## ✨ Features
- Extracts **FreeText** and **Sticky Note** annotations from PDFs  
- Exports data into a **predefined CRS Excel template** (`CRS.xlsx`)  
- Automatically fills **client name, project description, project number, and PO reference** when available  
- Adds page number and neatly formatted comments (line breaks per sentence)  
- **Excel output is styled** with borders, wrapped text, and column widths  
- Simple **Tkinter GUI** for ease of use (no command line required)  
- Buildable as a **standalone Windows EXE** with PyInstaller  

---

## 🚀 Quick Start

### Option 1 — Run from Source
1. Install Python 3.10 or newer.  
2. Clone this repository:
   ```bash
   git clone https://github.com/Bni-bourk/pdf-comment-extractor.git
   cd pdf-comment-extractor
