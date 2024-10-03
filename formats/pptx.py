from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from pptx import Presentation
from io import BytesIO
from PIL import Image as PILImage
from reportlab.lib.colors import HexColor


#conversor PPTX
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

#Extrae la imagen del shape y la devuelve como un flujo en memoria.
def extract_image_from_shape(shape):
    if shape.shape_type == 13:  # Es una imagen
        image_stream = BytesIO(shape.image.blob)
        return PILImage.open(image_stream)
    return None

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
    pdf = SimpleDocTemplate(output_path, pagesize=landscape(letter))

    elements = []
    styles = getSampleStyleSheet()
    style_normal = styles['Normal']
    style_heading = styles['Heading1']

    for slide_idx, slide in enumerate(prs.slides):
        elements.append(Paragraph(f"Slide {slide_idx + 1}", style_heading))

        # Añadir imagen de fondo si existe
        background_image = extract_background_image(slide)
        if background_image:
            img_buffer = BytesIO()
            background_image.save(img_buffer, format='PNG')
            img_buffer.seek(0)
            img = Image(img_buffer, width=8*inch, height=6*inch)
            elements.append(img)

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                # Extraer el texto y color
                text = extract_text_from_shape(shape)
                text_color = extract_color_from_shape(shape)
                if text:
                    style = ParagraphStyle(
                        'custom_style',
                        parent=style_normal,
                        textColor=HexColor(text_color) if text_color else colors.black,
                        alignment=TA_LEFT
                    )
                    elements.append(Paragraph(text, style))

            # Extraer e insertar las imágenes
            image = extract_image_from_shape(shape)
            if image:
                img_buffer = BytesIO()
                image.save(img_buffer, format='PNG')
                img_buffer.seek(0)
                img = Image(img_buffer, width=4*inch, height=3*inch)
                elements.append(img)

            # Extraer tablas
            table_data = extract_table_from_shape(shape)
            if table_data:
                table = Table(table_data)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                elements.append(table)

        elements.append(Spacer(1, 12))  # Añadir espacio entre slides
        elements.append(PageBreak())  # Nueva página para cada diapositiva

    pdf.build(elements)
    #print(f"PDF creado exitosamente en {output_path}")
