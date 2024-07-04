import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from docx2pdf import convert as docx_to_pdf_convert
import fitz  # PyMuPDF
import pdfkit

def pdf_to_word(pdf_file, docx_file):
    cv = Converter(pdf_file)
    cv.convert(docx_file, start=0, end=None)
    cv.close()

def word_to_pdf(docx_file, pdf_file):
    docx_to_pdf_convert(docx_file, pdf_file)

def pdf_to_text(pdf_file, text_file):
    pdf_document = fitz.open(pdf_file)
    with open(text_file, 'w') as f:
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text = page.get_text()
            f.write(text)

def browse_file(file_type):
    file_path = filedialog.askopenfilename()
    if file_path:
        save_as(file_type, file_path)

def save_as(file_type, input_file):
    file_path = filedialog.asksaveasfilename(defaultextension=file_type)
    if file_path:
        if file_type == '.docx':
            pdf_to_word(input_file, file_path)
        elif file_type == '.pdf':
            word_to_pdf(input_file, file_path)
        elif file_type == '.txt':
            pdf_to_text(input_file, file_path)
        messagebox.showinfo("Success", f"File has been saved as {file_path}")

def create_gui():
    root = tk.Tk()
    root.title("Document Converter")

    label = tk.Label(root, text="Select a conversion type:")
    label.pack(pady=10)

    pdf_to_word_btn = tk.Button(root, text="PDF to Word", command=lambda: browse_file('.docx'))
    pdf_to_word_btn.pack(pady=5)

    word_to_pdf_btn = tk.Button(root, text="Word to PDF", command=lambda: browse_file('.pdf'))
    word_to_pdf_btn.pack(pady=5)

    pdf_to_text_btn = tk.Button(root, text="PDF to Text", command=lambda: browse_file('.txt'))
    pdf_to_text_btn.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
