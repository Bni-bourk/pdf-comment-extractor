# 📄 PDF Comment Extractor (CRS Automation Tool)

A desktop application that automates the creation of **Client Response Sheets (CRS)** from annotated PDF files.  
Instead of manually searching through 400+ page PDFs and copying comments into Excel, this app extracts annotations in one click and fills them into a structured Excel template.

---

## ✨ Features
- Extracts **FreeText, Sticky Notes, and Highlights** from PDF annotations  
- Exports data into a **predefined CRS Excel template** (`CRS.xlsx`)  
- Auto-fills client and project information when available in the PDF  
- Includes **page number, type, subject, color, and comment text**  
- User-friendly **Tkinter GUI** (no command line needed)  
- Packaged with **PyInstaller** for easy distribution (Windows `.exe`)  

---

## 🖼️ Screenshots
*(Add screenshots in a folder called `assets/` and update links below)*

### GUI
![GUI Example](assets/gui.png)

### Excel Output
![Excel Example](assets/excel.png)

---

## 🚀 Quick Start

### Run from Source
1. Clone this repository:
   ```bash
   git clone https://github.com/<your-username>/pdf-comment-extractor.git
   cd pdf-comment-extractor
