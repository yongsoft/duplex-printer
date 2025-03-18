# Duplex Printing Wizard

A desktop application that helps you achieve duplex (double-sided) printing on printers without built-in duplex support.

## Overview

Duplex Printing Wizard is a Python-based desktop application that simplifies the process of printing double-sided documents on printers that don't have native duplex printing capabilities. The application works by splitting your document into odd and even pages, printing them in the correct order, and guiding you through the paper reinsertion process.

## Features

- **Multi-format Support**: Print double-sided documents from various file formats:
  - PDF files (*.pdf)
  - Word documents (*.doc, *.docx)
  - PowerPoint presentations (*.ppt, *.pptx)
  - Text files (*.txt)
  - Image files (*.jpg, *.jpeg, *.png)

- **Automatic Format Conversion**: Automatically converts non-PDF files to PDF format for printing

- **Custom Page Range**: Select specific pages or page ranges to print

- **Printer Detection**: Automatically detects available printers on your system

- **User-friendly Interface**: Simple, intuitive interface guides you through the printing process

## System Requirements

- Python 3.x
- macOS (currently optimized for macOS, may work on other platforms with modifications)
- LibreOffice (optional, for enhanced document conversion support)

## Installation

1. Clone this repository or download the source code

2. Create a virtual environment (recommended):
   ```bash
   python3 -m venv menv
   source menv/bin/activate  # On macOS/Linux
   ```

3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Additional dependencies (not included in requirements.txt):
   ```bash
   pip install pillow docx2pdf pymupdf
   ```

## Usage

1. Run the application:
   ```bash
   python main.py
   ```

2. The application will open with a simple interface:
   - Click "Browse..." to select the file you want to print
   - Select your printer from the dropdown list (click "Refresh Printer List" if needed)
   - Choose to print all pages or specify a custom page range
   - Click "Start Printing" to begin the process

3. The application will first print the odd-numbered pages

4. When prompted, reinsert the printed pages back into your printer according to your printer's paper feed orientation

5. Click "OK" when ready to print the even-numbered pages on the reverse side

## How It Works

The application:
1. Converts your document to PDF format if necessary
2. Analyzes the document to determine the total number of pages
3. Splits the document into odd and even pages
4. Prints the odd-numbered pages first
5. Prompts you to reinsert the paper
6. Prints the even-numbered pages on the reverse side

## Project Structure

- `main.py` - Entry point for the application
- `duplex_printer.py` - Main application code with the DuplexPrinterApp class
- `requirements.txt` - List of Python dependencies

## Dependencies

- PyPDF2 - For PDF manipulation
- Pillow (PIL) - For image processing
- docx2pdf - For Word document conversion
- PyMuPDF (fitz) - For text file conversion and PDF handling
- tkinter - For the graphical user interface

## License

This project is open source and available for personal and commercial use.

## Contributing

Contributions are welcome! Feel free to submit issues or pull requests to improve the application.