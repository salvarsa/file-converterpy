import xml.etree.ElementTree as ET
from reportlab.lib.colors import HexColor
from reportlab.lib.units import inch

def extract_text_properties(shape,xml,nsmap):
    if not hasattr(shape, 'text_frame'):
        return None
    properties = []
    
    # Parse XML para obtener detalles de la fuente
    root = ET.fromstring(shape.element.xml)
    #print(f'\n====size===={shape.element.xml}\n')
    nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
             'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}

    for paragraph in shape.text_frame.paragraphs:
        p_xml = root.find('.//a:p', nsmap)
        if p_xml is None:
            continue

        align = p_xml.find('.//a:pPr', nsmap)
        alignment = 0 
        if align is not None:
            algn_value = align.get('algn', '0')
            if algn_value == 'ctr':
                alignment = 1  # Center
            elif algn_value == 'r':
                alignment = 2  # Right
            else:
                try:
                    alignment = int(algn_value)
                except ValueError:
                    alignment = 0

        for run in paragraph.runs:
            r_xml = p_xml.find('.//a:r', nsmap)
            if r_xml is None:
                continue

            # Get font properties
            font_props = r_xml.find('.//a:rPr', nsmap)
            if font_props is None:
                continue

            # Font name
            typeface = font_props.find('.//a:rPr', nsmap)
            font_name = typeface.get('typeface', 'Helvetica') if typeface is not None else 'Helvetica'

            # Font size
            size = font_props.get('sz')
            
            font_size = int(size) / 100 if size is not None else 12

            # Font style
            bold = font_props.get('b') == '1'
            italic = font_props.get('i') == '1'

            # Font color
            color = font_props.find('.//a:solidFill/a:srgbClr', nsmap)
            text_color = f"#{color.get('val', '000000')}" if color is not None else "#000000"

            properties.append({
                'text': run.text,
                'font_name': font_name,
                'font_size': font_size,
                'bold': bold,
                'italic': italic,
                'color': text_color,
                'alignment': alignment
            })

    return properties

def get_shape_position(shape, slide_height):
    left = shape.left / 914400 * inch
    top = slide_height - (shape.top / 914400 * inch)
    width = shape.width / 914400 * inch
    height = shape.height / 914400 * inch
    return left, top, width, height