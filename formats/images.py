import os
from svglib.svglib import svg2rlg
from PIL import Image as PILImage
from reportlab.graphics import renderPDF
from reportlab.pdfgen import canvas

def convert_svg_to_pdf(input_path: str, output_path: str) -> None:
    drawing = svg2rlg(input_path)
    renderPDF.drawToFile(drawing, output_path)

def convert_image_to_pdf(input_path: str, output_path: str) -> None:
    image = PILImage.open(input_path)
    image = image.convert('RGB')
    image.save(output_path, 'PDF', resolution=100.0)

def convert_image_to_pdf(input_path: str, output_path: str) -> None:
    file_extension = os.path.splitext(input_path)[1].lower()
    
    if file_extension == '.svg':
        convert_svg_to_pdf(input_path, output_path)
    elif file_extension in ['.png', '.jpg', '.jpeg']:
        image = PILImage.open(input_path)
        image = image.convert('RGB')
        image.save(output_path, 'PDF', resolution=100.0)
    else:
        raise ValueError(f"INVALID_FORMAT: {file_extension}")

    return output_path