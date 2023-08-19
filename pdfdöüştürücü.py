import fitz  # PyMuPDF
import docx
import os
import tkinter as tk
from tkinter import filedialog


def pdf_to_docx(pdf_path, docx_path):
    pdf_document = fitz.open(pdf_path)
    doc = docx.Document()

    for page in pdf_document:
        text = page.get_text("text")
        doc.add_paragraph(text)

    doc.save(docx_path)


def convert_button_clicked():
    pdf_files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if pdf_files:
        for pdf_file in pdf_files:
            docx_file = os.path.splitext(pdf_file)[0] + ".docx"
            pdf_to_docx(pdf_file, docx_file)
        result_label.config(text="Seçilen PDF dosyaalrı için dönüşüm tamamlandı.")


# Arayüz
root = tk.Tk()
root.title("PDF - DOCX DÖNÜŞTÜRÜCÜ")    #başlık
root.config (bg="white")    #arkaplan rengi
root.geometry("300x300")    #pencere boyutu
convert_button = tk.Button(root, text="Dönüştürülecek PDF dosyalarını seçin", command=convert_button_clicked) #convert_button #buton
convert_button.pack(pady=110) #y ekseni padding

result_label = tk.Label(root, text="", font=("Helvetica", 12)) #font
result_label.pack()

root.mainloop()
