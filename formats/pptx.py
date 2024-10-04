from pptx import Presentation
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.lib import colors
from io import BytesIO
from PIL import Image as PILImage
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
from reportlab.platypus import PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

#conversor PPTX
# Extraer y ajustar las propiedades del texto
def extract_text_properties(paragraph, pdf, left, top, width, height):
    for run in paragraph.runs:
        text = run.text
        
        # Aplicar negrita, cursiva y subrayado
        font_name = "Helvetica"
        if run.font.bold:
            font_name += "-Bold"
        if run.font.italic:
            font_name = "Helvetica-Oblique"
        
        # Tamaño de fuente
        font_size = 12  # Tamaño por defecto
        if run.font.size:
            font_size = run.font.size.pt
        
        pdf.setFont(font_name, font_size)

        # Color del texto
        if run.font.color and hasattr(run.font.color, 'rgb') and run.font.color.rgb:
            color = f"#{run.font.color.rgb:06x}"
            pdf.setFillColor(HexColor(color))
        else:
            pdf.setFillColor(HexColor("#000000"))  # Color por defecto (negro)

        # Alineación del párrafo
        alignment = paragraph.alignment  # Extraer alineación
        if alignment == 1:  # Centrado
            text_width = pdf.stringWidth(text, font_name, font_size)
            left = (width - text_width) / 2
        elif alignment == 2:  # Derecha
            text_width = pdf.stringWidth(text, font_name, font_size)
            left = width - text_width

        # Dibujar el texto con espaciado vertical
        pdf.drawString(left, top, text)
        top -= font_size + 5  # Ajustar el espaciado entre líneas

        
#Extrae el texto de un shape respetando negritas, listas y colores o eso intento xd
def extract_text_from_shape(shape):
    text = ""
    if hasattr(shape, "text_frame"):
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.font.bold:
                    text += f"<b>{run.text}</b>"
                else:
                    text += run.text
            text += "\n"
    return text


def extract_color_from_shape(shape):
    if hasattr(shape, "text_frame") and shape.text_frame.paragraphs:
        run = shape.text_frame.paragraphs[0].runs[0]
        if run.font.color and hasattr(run.font.color, 'rgb') and run.font.color.rgb:
            return f"#{run.font.color.rgb:06x}"
    return None  # Si no hay color definido, retornar None

# Extraer y ajustar imágenes con justificación
def extract_image_from_shape(shape, pdf, width, height):
    if shape.shape_type == 13:  # Es una imagen
        img_stream = BytesIO(shape.image.blob)
        img = PILImage.open(img_stream)
        
        img_width = shape.width / 914400 * inch
        img_height = shape.height / 914400 * inch
        left = shape.left / 914400 * inch
        top = height - (shape.top / 914400 * inch) - img_height

        # Escalar la imagen si es más grande que la página
        if img_width > width:
            scale_factor = width / img_width
            img_width *= scale_factor
            img_height *= scale_factor
        
        if img_height > height:
            scale_factor = height / img_height
            img_width *= scale_factor
            img_height *= scale_factor

        # Justificación de la imagen (centrada)
        if left + img_width > width:  # Si la imagen se sale de los márgenes, centrarla
            left = (width - img_width) / 2

        pdf.drawInlineImage(img, left, top, img_width, img_height)



def extract_background_image(slide):
    background = slide.background
    if background.fill.type == 6:  # Background con imagen
        image = background.fill.picture.image
        image_stream = BytesIO(image.blob)
        return PILImage.open(image_stream)
    return None

def extract_table_from_shape(shape):
    if shape.has_table:
        table_data = []
        for row in shape.table.rows:
            row_data = [cell.text for cell in row.cells]
            table_data.append(row_data)
        return table_data
    return None

def convert_pptx_to_pdf(input_path: str, output_path: str):
    prs = Presentation(input_path)
    pdf = canvas.Canvas(output_path, pagesize=landscape(letter))
    width, height = landscape(letter)

    for slide in prs.slides:
        # Fondo
        fill = slide.background.fill
        if fill.type == 1:  # Color sólido
            bg_color = f"#{fill.fore_color.rgb:06x}"
            pdf.setFillColor(HexColor(bg_color))
            pdf.rect(0, 0, width, height, fill=1)

        # Procesar cada forma
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    left = shape.left / 914400 * inch
                    top = height - (shape.top / 914400 * inch) - (shape.height / 914400 * inch)
                    
                    # Ajustar texto con justificación y propiedades
                    extract_text_properties(paragraph, pdf, left, top, width, height)
            
            elif shape.shape_type == 13:  # Imagen
                extract_image_from_shape(shape, pdf, width, height)

        pdf.showPage()

    pdf.save()