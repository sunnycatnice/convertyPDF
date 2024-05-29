import subprocess
import os
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from .printer import print_file

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
            # Less space before headings to avoid splitting images
            para.paragraph_format.space_before = Pt(spacing_value / 2)
            para.paragraph_format.space_after = Pt(spacing_value / 2)
        else:
            para.paragraph_format.space_before = Pt(spacing_value)
            para.paragraph_format.space_after = Pt(spacing_value)

def convert_word_to_pdf(word_files, output_dir, print_after_conversion=False, formatted=False):
    pdf_files = []
    for word_file in word_files:
        if formatted:
            doc = Document(word_file)
            remove_table_borders(doc)
            add_cell_padding(doc)
            add_paragraph_spacing(doc)
            temp_word_file = os.path.join(output_dir, os.path.basename(word_file).replace('.docx', '_formatted.docx'))
            doc.save(temp_word_file)
            word_file_to_convert = temp_word_file
        else:
            word_file_to_convert = word_file

        pdf_file = os.path.join(output_dir, os.path.basename(word_file_to_convert).replace('.docx', '.pdf'))
        
        # Convert using LibreOffice
        result = subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', word_file_to_convert, '--outdir', output_dir], capture_output=True)
        if result.returncode != 0:
            raise RuntimeError(f"Error converting {word_file} to PDF: {result.stderr.decode()}")
        
        # Check if the converted file exists
        if os.path.exists(pdf_file):
            pdf_files.append(pdf_file)
            print(f"Converted {word_file} to {pdf_file}")
        else:
            raise RuntimeError(f"Converted file {pdf_file} not found")
    
    if print_after_conversion:
        for pdf_file in pdf_files:
            print_file(pdf_file)
