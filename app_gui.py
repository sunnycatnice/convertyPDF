import os
import subprocess
import platform
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from workers.converter import print_file

def remove_table_borders(doc):
    for table in doc.tables:
        tbl = table._element
        tblBorders = tbl.xpath('.//w:tblBorders')
        for tblBorder in tblBorders:
            for border in tblBorder:
                border.attrib.clear()
                border.set(qn('w:val'), 'nil')

def add_cell_padding(doc, padding_value=100):
    for table in doc.tables:
        for cell in table._cells:
            tc = cell._element
            tcPr = tc.get_or_add_tcPr()
            tcMar = OxmlElement('w:tcMar')
            for margin in ['top', 'left', 'bottom', 'right']:
                node = OxmlElement(f'w:{margin}')
                node.set(qn('w:w'), str(padding_value))
                node.set(qn('w:type'), 'dxa')
                tcMar.append(node)
            tcPr.append(tcMar)

def add_paragraph_spacing(doc, spacing_value=3):
    for para in doc.paragraphs:
        if 'Heading' in para.style.name:
            para.paragraph_format.space_before = Pt(spacing_value / 2)
            para.paragraph_format.space_after = Pt(spacing_value / 2)
        else:
            para.paragraph_format.space_before = Pt(spacing_value)
            para.paragraph_format.space_after = Pt(spacing_value)

def convert_word_to_pdf(word_files, print_after_conversion=False, formatted=False, log_widget=None):
    pdf_files = []
    for word_file in word_files:
        try:
            if formatted:
                doc = Document(word_file)
                remove_table_borders(doc)
                add_cell_padding(doc)
                add_paragraph_spacing(doc)
                temp_word_file = word_file.replace('.docx', '_formatted.docx')
                doc.save(temp_word_file)
                if log_widget:
                    log_widget.insert(tk.END, f"Formatted file saved as: {temp_word_file}\n")
                word_file_to_convert = temp_word_file
            else:
                word_file_to_convert = word_file

            output_dir = os.path.dirname(word_file)
            pdf_file = os.path.join(output_dir, os.path.splitext(os.path.basename(word_file_to_convert))[0] + '.pdf')
            
            if platform.system() == "Windows":
                soffice_path = "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
                command = [soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, word_file_to_convert]
            else:
                command = ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, word_file_to_convert]
                
            if log_widget:
                log_widget.insert(tk.END, f"Running command: {' '.join(command)}\n")

            result = subprocess.run(command, capture_output=True)

            if log_widget:
                log_widget.insert(tk.END, f"Command output: {result.stdout.decode()}\n")
                log_widget.insert(tk.END, f"Command error: {result.stderr.decode()}\n")
                
            if result.returncode != 0:
                raise RuntimeError(f"Error converting {word_file} to PDF: {result.stderr.decode()}")

            if os.path.exists(pdf_file):
                pdf_files.append(pdf_file)
                if log_widget:
                    log_widget.insert(tk.END, f"Converted {word_file} to {pdf_file}\n")
            else:
                raise RuntimeError(f"Converted file {pdf_file} not found")

        except Exception as e:
            if log_widget:
                log_widget.insert(tk.END, f"✘ Conversion failed: {word_file}\nError: {str(e)}\n")
            else:
                print(f"✘ Conversion failed: {word_file}\nError: {str(e)}\n")

    if print_after_conversion:
        for pdf_file in pdf_files:
            print_file(pdf_file)

def select_files():
    files = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx")])
    if files:
        file_list.extend(files)
        update_file_list()

def update_file_list():
    listbox.delete(0, tk.END)
    for f in file_list:
        listbox.insert(tk.END, f)

def clear_file_list():
    file_list.clear()
    update_file_list()

def convert_files(print_after=False, formatted=False):
    if not file_list:
        messagebox.showwarning("No files selected", "Please select one or more Word files to convert.")
        return

    log_widget.delete(1.0, tk.END)  # Clear previous log
    convert_word_to_pdf(file_list, print_after, formatted, log_widget)

app = tk.Tk()
app.title("ConvertifyPDF")
app.geometry("600x400")

file_list = []

frame = tk.Frame(app)
frame.pack(pady=10)

listbox = tk.Listbox(frame, width=80, height=10)
listbox.pack(side=tk.LEFT, padx=(0, 10))

scrollbar = tk.Scrollbar(frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

listbox.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=listbox.yview)

button_frame = tk.Frame(app)
button_frame.pack(pady=10)

select_button = tk.Button(button_frame, text="Select Word Files to Convert to PDF", command=select_files)
select_button.grid(row=0, column=0, padx=5)

clear_button = tk.Button(button_frame, text="Clear List", command=clear_file_list)
clear_button.grid(row=0, column=1, padx=5)

convert_button = tk.Button(app, text="Convert", command=lambda: convert_files(print_after=False, formatted=False))
convert_button.pack(pady=5)

convert_print_button = tk.Button(app, text="Convert and Print", command=lambda: convert_files(print_after=True, formatted=False))
convert_print_button.pack(pady=5)

convert_format_button = tk.Button(app, text="Convert with Formatting", command=lambda: convert_files(print_after=False, formatted=True))
convert_format_button.pack(pady=5)

convert_format_print_button = tk.Button(app, text="Convert with Formatting and Print", command=lambda: convert_files(print_after=True, formatted=True))
convert_format_print_button.pack(pady=5)

log_widget = ScrolledText(app, width=80, height=10, wrap=tk.WORD)
log_widget.pack(pady=10)

app.mainloop()
