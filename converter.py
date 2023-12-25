from pdf2docx import Converter
from docx import Document
from docx2pdf import convert
import time


def convert_pdf_to_docx(pdf_file, docx_file):
    cv = Converter(pdf_file)
    cv.convert(docx_file, start=0, end=None, )
    cv.close()

def convert_docx_to_pdf(docx_file, pdf_file=None):
    # Convert the docx file to pdf
    # If pdf_file is None, it will save the PDF in the same directory as the DOCX file
    convert(docx_file, pdf_file)


def replace_text_in_docx(file_path, old_text, new_text):
    # Open the DOCX file
    doc = Document(file_path)

    counter = 0
    # Iterate through each paragraph in the document
    for p in doc.paragraphs:
        if old_text in p.text:
            inline = p.runs
            # Loop through each run in the paragraph
            for i in range(len(inline)):
                if old_text in inline[i].text:
                    text = inline[i].text.replace(old_text, new_text)
                    inline[i].text = text
                    counter += 1

    print(f'found {counter} occurences')
    # Save the modified document
    doc.save(file_path)


start = time.time()
print("hello")
convert_pdf_to_docx("testTemplate.pdf", "example.docx")

replace_text_in_docx("example.docx", "Test", "Test3")

convert_docx_to_pdf("example.docx", "output.pdf")
end = time.time()
print(f'{end - start} seconds elapsed')