from PyPDF2 import PdfMerger
def merge_pdfs(filepaths, output_path):
    merger = PdfMerger()
    for path in filepaths:
        merger.append(path)
    merger.write(output_path)
    merger.close()
