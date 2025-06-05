import os
import sys
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER
from PyPDF2 import PdfReader, PdfWriter
import tkinter as tk
from tkinter import filedialog

print(f"Automatic Exhibit Cover Sheets v0.1.  Anson Fung")


def create_cover_sheet(filename):
    """
    Creates a cover sheet PDF in memory with the given filename (without extension) centered on the page using Times New Roman size 20.
    The text is wrapped if it exceeds the specified width.
    
    :param filename: The filename to display on the cover sheet.
    :return: A BytesIO buffer containing the cover sheet PDF.
    """
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    name_without_ext = os.path.splitext(filename)[0]
    
    # Set up the style for the paragraph
    styles = getSampleStyleSheet()
    style = styles['Normal']
    style.fontName = 'Times-Bold'
    style.fontSize = 20
    style.leading = 28
    style.alignment = TA_CENTER  # Center alignment
    
    # Create the paragraph with the filename
    p = Paragraph(name_without_ext, style)
    
    # Set the available width for word wrapping (e.g., 468 points for 1-inch margins)
    avail_width = 468
    p.wrap(avail_width, 1000)  # Wrap the paragraph to the available width
    
    # Calculate the position to center the paragraph horizontally and vertically
    x = (width - avail_width) / 2
    y = (height / 2) + (p.height / 2)
    
    # Draw the paragraph on the canvas
    p.drawOn(c, x, y)
    
    c.save()
    buffer.seek(0)
    return buffer

def add_cover_to_pdf(cover_buffer, original_path, output_path):
    """
    Adds the cover sheet to the original PDF and saves the result to output_path.
    
    :param cover_buffer: BytesIO buffer containing the cover sheet PDF.
    :param original_path: Path to the original PDF file.
    :param output_path: Path to save the new PDF with the cover sheet.
    """
    cover_reader = PdfReader(cover_buffer)
    original_reader = PdfReader(original_path)
    writer = PdfWriter()
    
    # Add the cover page
    writer.add_page(cover_reader.pages[0])
    
    # Add the original pages
    for page in original_reader.pages:
        page.compress_content_streams()  # Lossless compression.  This is CPU intensive!
        writer.add_page(page)
    
    # Write to the output file
    with open(output_path, 'wb') as f:
        writer.write(f)

def main(folder_path):
    """
    Processes all PDF files in the given folder by adding a cover sheet with the filename (without extension).
    
    :param folder_path: Path to the folder containing PDF files.
    """
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pdf'):
            original_path = os.path.join(folder_path, filename)
            print(f"Processing {filename}...")
            cover_buffer = create_cover_sheet(filename)
            output_path = os.path.join(folder_path, f"cover_{filename}")
            add_cover_to_pdf(cover_buffer, original_path, output_path)
            print(f"Saved {output_path}")

if __name__ == "__main__":
    # Check if a folder path is provided via command line
    if len(sys.argv) == 2:
        folder_path = sys.argv[1]
    else:
        # Create a hidden Tkinter root window
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        # Open a dialog box for folder selection
        folder_path = filedialog.askdirectory(title="Select a folder containing PDFs")
        if not folder_path:
            print("No folder selected. Exiting.")
            sys.exit(1)
    
    # Validate that the path is a valid directory
    if not os.path.isdir(folder_path):
        print("The specified path is not a valid directory. Exiting.")
        sys.exit(1)
    
    # Process the PDFs in the folder
    main(folder_path)