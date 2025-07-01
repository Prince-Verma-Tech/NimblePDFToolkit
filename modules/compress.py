from pdf2image import convert_from_path
from PIL import Image
def compress_pdf(filepath, output_path):
    images = convert_from_path(filepath, 200)
    compressed_images = [img.convert('RGB') for img in images]
    compressed_images[0].save(output_path, save_all=True, append_images=compressed_images[1:], quality=60)
