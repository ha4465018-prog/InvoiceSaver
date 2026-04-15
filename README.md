# ◈ STYLUXE INVOICE PRO

**Styluxe Invoice Pro** is a comprehensive, local desktop application designed to streamline the processing, organization, and storage of invoices. Built with Python and Tkinter, it leverages OpenCV and Tesseract OCR to automatically extract dates and total amounts from uploaded or webcam-captured invoice images. 

## ✨ Key Features

* **📷 Dual Input Modes:** Upload existing images (PNG, JPG, TIFF, WEBP) or capture physical invoices directly using your webcam.
* **🛠️ Image Enhancement Studio:** Built-in tools to adjust contrast, brightness, rotation, and apply B/W thresholds to improve OCR accuracy on low-quality scans.
* **🧠 Smart Date & Total Extraction:** Uses advanced RegEx and OCR-noise reduction to accurately find dates and grand totals from raw text.
* **🗄️ Local Database:** Securely stores all extracted data (Date, Total, Raw Text, and Image paths) in a local SQLite database (`styluxe_invoices.db`).
* **📊 Analytics Dashboard:** Real-time stats showing total invoices processed, sum of total amounts, and last scan time.
* **📥 CSV Export & Management:** A dedicated database viewer to search records, view raw OCR text, open original images, delete entries, and export data to CSV.

## ⚙️ Prerequisites & Installation

To run this application, you need Python installed on your computer, along with the Tesseract OCR engine.

### Step 1: Install Tesseract OCR Engine (Required)
The application relies on Tesseract to read text from images. You **must** install this on your operating system:
* **Windows:** Download the installer from [UB-Mannheim Tesseract Wiki](https://github.com/UB-Mannheim/tesseract/wiki).
  *(Install it in the default directory: `C:\Program Files\Tesseract-OCR\tesseract.exe`)*
* **Mac (Homebrew):** `brew install tesseract`
* **Linux (Ubuntu/Debian):** `sudo apt-get install tesseract-ocr`

### Step 2: Install Python Dependencies
Open your terminal or command prompt, navigate to the project folder, and run:
```bash
pip install -r requirements.txt# InvoiceSaver
Is help to save invoice and total expanse 
