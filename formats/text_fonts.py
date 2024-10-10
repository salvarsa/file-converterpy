from pptx import Presentation
import xml.etree.ElementTree as ET
from reportlab.lib.colors import HexColor

def text_fonts(presentation, pdf):
    for slide_number, slide in enumerate(presentation.slides, start=1):
  
        # Suponiendo que tienes el XML en una cadena
        root = ET.fromstring(slide.background.element.xml)
        nsmap = slide.part.slide.element.nsmap

        # Extraer texto
        text_elements = root.findall('.//a:t', nsmap)
        text_content = '<br>'.join([elem.text for elem in text_elements if elem.text])
    
        #Extraer tamaño
        size = 12
        text_size = root.find('.//a:r/a:rPr', nsmap)
        if text_size is not None:
            sz = text_size.get('sz')
            if sz is not None:
                size = f'{(int(sz) / 100):.0f}'

        #Extraer cordenadas del texto
        text_position = root.find('.//p:sp[1]/p:spPr/a:xfrm/a:off', nsmap)
        if text_position is not None:
            text_x = text_position.get('x')
            text_y = text_position.get('y')

        # Buscar la etiqueta <a:buFont> typo de fuente del texto
        text_font_type = root.find('.//a:buFont', nsmap)
        if text_font_type is not None:
            font_type = text_font_type.get('typeface')

        text_line_spacing = root.find('.//a:spcPct', nsmap)
        if text_line_spacing is not None:
            line_spacing_percentage = text_line_spacing.get('val')

        # Extraer márgenes de <a:bodyPr>
        body_pr = root.find('.//a:bodyPr', nsmap)
        if body_pr is not None:
            lIns = int(body_pr.get('lIns', 0))
            rIns = int(body_pr.get('rIns', 0))
            tIns = int(body_pr.get('tIns', 0))
            bIns = int(body_pr.get('bIns', 0))

        p_pr = root.find('.//a:pPr', nsmap)
        # if p_pr is not None:
        #     spcBef = int(p_pr.find('.//a:spcBef', nsmap).get('val', 0))
        #     spcAft = int(p_pr.find('.//a:spcAft', nsmap).get('val', 0))    

        lIns_points = lIns / 1000
        rIns_points = rIns / 1000
        tIns_points = tIns / 1000
        bIns_points = bIns / 1000
        # spcBef_points = spcBef / 1000
        # spcAft_points = spcAft / 1000
        
        pdf.setFont('Helvetica', size)
        pdf.setFillColor(HexColor("#000000"))
        
        pdf.drawString(lIns_points, tIns_points, text_content)
   
    
        # elements = {
        # 'text': text_content,
        # 'size': size,
        # 'text_x': text_x,  # Ajusta la posición en la página
        # 'text_y': text_x,
        # 'text_font_type': font_type,  # Asegúrate de que la fuente esté disponible
        # 'line_percentage': line_spacing_percentage,
        # 'lIns': lIns_points,
        # 'rIns': rIns_points,
        # 'tIns': tIns_points,
        # 'bIns': bIns_points,
        # # 'spcBef': spcBef_points,
        # # 'spcAft': spcAft_points,
        # }
        
