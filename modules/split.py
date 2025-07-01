from PyPDF2 import PdfReader, PdfWriter
def split_pdf(filepath, start, end, output_path):
    reader = PdfReader(filepath)
    writer = PdfWriter()
    for page in range(start - 1, end):
        writer.add_page(reader.pages[page])
    with open(output_path, 'wb') as f:
        writer.write(f)
