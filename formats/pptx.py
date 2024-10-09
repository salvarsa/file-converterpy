from pptx import Presentation
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from io import BytesIO
from PIL import Image as PILImage

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
    if shape.shape_type == 13:  # Verifica si es una imagen
        img_stream = BytesIO(shape.image.blob)  # Extrae los datos binarios de la imagen
        img = PILImage.open(img_stream)  # Abre la imagen usando Pillow
        
        # Dimensiones y posición de la imagen en el PowerPoint, convertidas a puntos para el PDF
        img_width = shape.width / 914400 * inch
        img_height = shape.height / 914400 * inch
        left = shape.left / 914400 * inch
        top = height - (shape.top / 914400 * inch) - img_height
        
        # Calcular el factor de escala para mantener la proporción y ajustar al tamaño de la página
        scale_factor = min(width / img_width, height / img_height, 1)
        scaled_width = img_width * scale_factor
        scaled_height = img_height * scale_factor
        
        # Calcular la posición horizontal respetando la alineación original
        slide_center = width / 2
        image_center = left + (img_width / 2)
        
        if image_center < slide_center:
            # La imagen está originalmente en la mitad izquierda
            new_left = left
        elif image_center > slide_center:
            # La imagen está originalmente en la mitad derecha
            new_left = min(left, width - scaled_width)
        else:
            # La imagen está centrada originalmente
            new_left = (width - scaled_width) / 2
        
        # Asegurarse de que la imagen no se salga de los márgenes
        new_left = max(0, min(new_left, width - scaled_width))
        
        # Ajustar la posición vertical si es necesario
        new_top = max(0, min(top, height - scaled_height))
        
        # Dibuja la imagen en el PDF con el tamaño y posición calculados
        pdf.drawInlineImage(img, new_left, new_top, scaled_width, scaled_height)

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
    
    # Obtener el tamaño de la primera diapositiva
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    
    # Convertir a puntos (1 punto = 1/72 pulgadas)
    width = slide_width / 914400 * 72
    height = slide_height / 914400 * 72
    
    pdf = canvas.Canvas(output_path, pagesize=(width, height))
    
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