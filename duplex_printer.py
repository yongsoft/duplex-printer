import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
import docx2pdf
import fitz  # PyMuPDF
import io

class DuplexPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Duplex Printing Wizard")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        self.root.resizable(True, True)
        
        self.file_path = None
        self.total_pages = 0
        self.temp_pdf_path = None
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Duplex Printing Wizard", font=("Arial", 18, "bold"))
        title_label.pack(pady=10)
        
        # Description
        desc_label = ttk.Label(main_frame, text="This wizard helps you achieve duplex printing on printers without built-in duplex support.\nSupports PDF, Word, PowerPoint, text files and image formats.")
        desc_label.pack(pady=5)
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="Select File", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        
        file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=file_path_var, width=50)
        file_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        browse_button = ttk.Button(file_frame, text="Browse...", command=lambda: self.browse_file(file_path_var))
        browse_button.pack(side=tk.RIGHT, padx=5)
        
        # Printer selection frame
        printer_frame = ttk.LabelFrame(main_frame, text="Select Printer", padding="10")
        printer_frame.pack(fill=tk.X, pady=10)
        
        self.printer_var = tk.StringVar()
        self.printer_combo = ttk.Combobox(printer_frame, textvariable=self.printer_var, width=50)
        self.printer_combo.pack(fill=tk.X, padx=5)
        
        refresh_button = ttk.Button(printer_frame, text="Refresh Printer List", command=self.refresh_printers)
        refresh_button.pack(pady=5)
        
        # Print options frame
        options_frame = ttk.LabelFrame(main_frame, text="Print Options", padding="10")
        options_frame.pack(fill=tk.X, pady=10)
        
        # Page range
        range_frame = ttk.Frame(options_frame)
        range_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(range_frame, text="Page Range:").pack(side=tk.LEFT, padx=5)
        
        self.range_var = tk.StringVar(value="all")
        ttk.Radiobutton(range_frame, text="All Pages", variable=self.range_var, value="all").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(range_frame, text="Custom Range", variable=self.range_var, value="range").pack(side=tk.LEFT, padx=5)
        
        self.range_entry_var = tk.StringVar()
        self.range_entry = ttk.Entry(range_frame, textvariable=self.range_entry_var, width=15)
        self.range_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(range_frame, text="(e.g.: 1-5,8,11-13)").pack(side=tk.LEFT, padx=5)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=20)
        
        start_button = ttk.Button(buttons_frame, text="Start Printing", command=self.start_printing)
        start_button.pack(side=tk.RIGHT, padx=5)
        
        # Status frame
        self.status_frame = ttk.LabelFrame(main_frame, text="Status", padding="10")
        self.status_frame.pack(fill=tk.X, pady=10)
        
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(self.status_frame, textvariable=self.status_var)
        status_label.pack(fill=tk.X)
        
        # Initialize
        self.refresh_printers()
    
    def browse_file(self, file_path_var):
        file_path = filedialog.askopenfilename(filetypes=[
            ("All Supported Files", ("*.pdf", "*.doc", "*.docx", "*.ppt", "*.pptx", "*.txt", "*.jpg", "*.jpeg", "*.png")),
            ("PDF Files", "*.pdf"),
            ("Word Documents", ("*.doc", "*.docx")),
            ("PowerPoint Presentations", ("*.ppt", "*.pptx")),
            ("Text Files", "*.txt"),
            ("Image Files", ("*.jpg", "*.jpeg", "*.png")),
            ("All Files", "*.*")
        ], title="Select File to Print")
        
        if file_path:
            file_path_var.set(file_path)
            self.file_path = file_path
            self.convert_to_pdf()
    
    def convert_to_pdf(self):
        if not self.file_path:
            return
        
        file_ext = os.path.splitext(self.file_path)[1].lower()
        temp_dir = os.path.join(os.path.expanduser("~"), ".duplexprinter")
        os.makedirs(temp_dir, exist_ok=True)
        
        try:
            if file_ext == '.pdf':
                self.temp_pdf_path = self.file_path
            else:
                self.temp_pdf_path = os.path.join(temp_dir, "converted.pdf")
                
                if file_ext in ['.doc', '.docx']:
                    self.status_var.set("Converting Word document...")
                    self.root.update()
                    try:
                        docx2pdf.convert(self.file_path, self.temp_pdf_path)
                    except Exception as docx_error:
                        self.status_var.set("Using alternative method to convert Word document...")
                        self.root.update()
                        try:
                            # Fallback to LibreOffice conversion
                            cmd = ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', temp_dir, self.file_path]
                            subprocess.run(cmd, check=True)
                            os.rename(os.path.join(temp_dir, os.path.splitext(os.path.basename(self.file_path))[0] + '.pdf'),
                                     self.temp_pdf_path)
                        except Exception as lo_error:
                            raise Exception(f"Word conversion failed: {str(docx_error)}\nBackup conversion also failed: {str(lo_error)}")
                
                elif file_ext in ['.ppt', '.pptx']:
                    self.status_var.set("Converting PowerPoint document...")
                    self.root.update()
                    # Using system command for PowerPoint conversion
                    cmd = ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', temp_dir, self.file_path]
                    subprocess.run(cmd, check=True)
                    os.rename(os.path.join(temp_dir, os.path.splitext(os.path.basename(self.file_path))[0] + '.pdf'),
                             self.temp_pdf_path)
                
                elif file_ext in ['.txt']:
                    self.status_var.set("Converting text file...")
                    self.root.update()
                    doc = fitz.open()
                    page = doc.new_page()
                    with open(self.file_path, 'r', encoding='utf-8') as f:
                        text = f.read()
                    page.insert_text((50, 50), text)
                    doc.save(self.temp_pdf_path)
                    doc.close()
                
                elif file_ext in ['.jpg', '.jpeg', '.png']:
                    self.status_var.set("Converting image...")
                    self.root.update()
                    image = Image.open(self.file_path)
                    if image.mode == 'RGBA':
                        image = image.convert('RGB')
                    image.save(self.temp_pdf_path, 'PDF')
            
            self.analyze_document()
            
        except Exception as e:
            messagebox.showerror("Error", f"File conversion failed: {str(e)}")
            self.status_var.set("File conversion failed")
            self.temp_pdf_path = None
    
    def analyze_document(self):
        if not self.temp_pdf_path or not os.path.exists(self.temp_pdf_path):
            messagebox.showerror("Error", "Please select a valid file first")
            return
        
        try:
            with open(self.temp_pdf_path, 'rb') as file:
                pdf = PdfReader(file)
                self.total_pages = len(pdf.pages)
                self.status_var.set(f"Document loaded, {self.total_pages} pages total")
        except Exception as e:
            messagebox.showerror("Error", f"Cannot analyze document: {str(e)}")
    
    def refresh_printers(self):
        try:
            # Get list of available printers on macOS
            result = subprocess.run(['lpstat', '-a'], capture_output=True, text=True)
            printers = []
            
            if result.returncode == 0:
                for line in result.stdout.splitlines():
                    if line.strip():
                        printer_name = line.split()[0].strip()
                        printers.append(printer_name)
            
            self.printer_combo['values'] = printers
            
            if printers:
                self.printer_combo.current(0)
                self.status_var.set(f"Found {len(printers)} printers")
            else:
                self.status_var.set("No printers found")
                
        except Exception as e:
            self.status_var.set(f"Failed to get printer list: {str(e)}")
    
    def parse_page_range(self):
        if self.range_var.get() == "all":
            return list(range(1, self.total_pages + 1))
        
        page_range = self.range_entry_var.get().strip()
        if not page_range:
            return list(range(1, self.total_pages + 1))
        
        pages = []
        parts = page_range.split(',')
        
        for part in parts:
            if '-' in part:
                start, end = map(int, part.split('-'))
                pages.extend(range(start, end + 1))
            else:
                pages.append(int(part))
        
        return pages
    
    def create_odd_even_pdfs(self, pages):
        if not self.file_path:
            return None, None
        
        try:
            with open(self.file_path, 'rb') as file:
                pdf = PdfReader(file)
                
                odd_writer = PdfWriter()
                even_writer = PdfWriter()
                
                # Split pages into odd and even lists
                odd_pages = []
                even_pages = []
                for page_num in pages:
                    if page_num > len(pdf.pages):
                        continue
                    
                    # PDF pages are 0-indexed
                    pdf_page = pdf.pages[page_num - 1]
                    
                    if page_num % 2 == 1:  # Odd page
                        odd_pages.append(pdf_page)
                    else:  # Even page
                        even_pages.append(pdf_page)
                
                # Reverse the order of pages
                odd_pages.reverse()
                even_pages.reverse()
                
                # Add pages in reversed order
                for page in odd_pages:
                    odd_writer.add_page(page)
                for page in even_pages:
                    even_writer.add_page(page)
                
                # Add a blank page if there are even pages and total page count is odd
                if even_writer.pages and len(pages) % 2 == 1:
                    blank_writer = PdfWriter()
                    blank_writer.add_blank_page(width=pdf.pages[0].mediabox.width, height=pdf.pages[0].mediabox.height)
                    even_writer.add_page(blank_writer.pages[0])
                
                # Save temporary files
                temp_dir = os.path.join(os.path.expanduser("~"), ".duplexprinter")
                os.makedirs(temp_dir, exist_ok=True)
                
                odd_path = os.path.join(temp_dir, "odd_pages.pdf")
                even_path = os.path.join(temp_dir, "even_pages.pdf")
                
                with open(odd_path, 'wb') as odd_file:
                    odd_writer.write(odd_file)
                
                with open(even_path, 'wb') as even_file:
                    even_writer.write(even_file)
                
                return odd_path, even_path
                
        except Exception as e:
            messagebox.showerror("Failed to create temporary files: {str(e)}", parent=self.root)
            return None, None
    
    def print_pdf(self, pdf_path, printer_name):
        try:
            cmd = ['lp', '-d', printer_name, pdf_path]
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode != 0:
                raise Exception(result.stderr)
            
            return True
        except Exception as e:
            # Center the error dialog relative to the main window
            self.root.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() - 300) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - 100) // 2
            messagebox.showerror("Print Error", f"Printing failed: {str(e)}", parent=self.root)
            return False
    
    def start_printing(self):
        if not self.file_path:
            # Center the error dialog relative to the main window
            self.root.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() - 300) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - 100) // 2
            messagebox.showerror("Error", "Please select a file to print first", parent=self.root)
            return
        
        printer_name = self.printer_var.get()
        if not printer_name:
            # Center the error dialog relative to the main window
            self.root.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() - 300) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - 100) // 2
            messagebox.showerror("Error", "Please select a printer first", parent=self.root)
            return
        
        pages = self.parse_page_range()
        odd_path, even_path = self.create_odd_even_pdfs(pages)
        
        if not odd_path or not even_path:
            return
        
        self.status_var.set("Starting to print odd pages...")
        self.root.update()
        if not self.print_pdf(odd_path, printer_name):
            return
        
        # Center the notice dialog relative to the main window
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 300) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 100) // 2
        messagebox.showinfo("Notice", "Even pages printing completed. Please reinsert the paper in the correct orientation.\nClick OK when ready to print odd pages.", parent=self.root)
        
        # Then print odd pages
        self.status_var.set("Printing odd pages...")
        self.root.update()
        if not self.print_pdf(odd_path, printer_name):
            return
        
        self.status_var.set("Printing completed")
        # Center the completion dialog relative to the main window
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 300) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 100) // 2
        messagebox.showinfo("Complete", "All pages have been sent to the printer!", parent=self.root)