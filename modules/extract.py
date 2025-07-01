import fitz
import os
def extract_text_images(filepath, image_dir):
    doc = fitz.open(filepath)
    text = ""
    image_paths = []
    for i, page in enumerate(doc):
        text += page.get_text()
        for img_index, img in enumerate(page.get_images()):
            xref = img[0]
            pix = fitz.Pixmap(doc, xref)
            img_path = os.path.join(image_dir, f"page{i+1}_img{img_index+1}.png")
            pix.save(img_path)
            image_paths.append(img_path)
    return text, image_paths
