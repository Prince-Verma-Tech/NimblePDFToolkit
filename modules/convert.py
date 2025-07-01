from pdf2image import convert_from_path
import img2pdf
import os
def pdf_to_images(pdf_path, image_dir):
    images = convert_from_path(pdf_path)
    image_paths = []
    for i, img in enumerate(images):
        path = os.path.join(image_dir, f"page_{i+1}.png")
        img.save(path, 'PNG')
        image_paths.append(path)
    return image_paths
def images_to_pdf(image_paths, output_path):
    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(image_paths))
