from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import tempfile
def create_watermark(watermark_text):
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    c = canvas.Canvas(temp.name, pagesize=letter)
    c.setFont("Helvetica", 40)
    c.setFillGray(0.5, 0.5)
    c.drawString(100, 500, watermark_text)
    c.save()
    return temp.name
def add_watermark(pdf_path, watermark_text, output_path):
    watermark = PdfReader(create_watermark(watermark_text))
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    watermark_page = watermark.pages[0]
    for page in reader.pages:
        page.merge_page(watermark_page)
        writer.add_page(page)
    with open(output_path, 'wb') as f:
        writer.write(f)
