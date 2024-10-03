import os
import sys
import pypandoc
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
from formats.docx import convert_docx_to_pdf
from formats.xlxs import convert_xlsx_to_pdf
from formats.pptx import convert_pptx_to_pdf


def download_pdf():
    return os.path.join(os.path.expanduser('~'), 'Downloads')

def generate_unique_filename(base_path, base_name, extension):
    counter = 1
    new_path = os.path.join(base_path, f"{base_name}{extension}")
    while os.path.exists(new_path):
        new_path = os.path.join(base_path, f"{base_name}_{counter}{extension}")
        counter += 1
    return new_path

#conversor TXT   
def convert_txt_to_pdf(input_path: str, output_path: str) -> None:
    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()
    pypandoc.convert_text(content, 'pdf', format='markdown', outputfile=output_path)

def convert_svg_to_pdf(input_path: str, output_path: str) -> None:
    drawing = svg2rlg(input_path)
    renderPDF.drawToFile(drawing, output_path)

def convert_file_to_pdf(input_path: str, output_path: str) -> None:
    try:
        ext = os.path.splitext(input_path)[1].lower()
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        
        downloads = download_pdf()
        output_path = generate_unique_filename(downloads, base_name, '.pdf')
        
        if ext == '.docx':
            convert_docx_to_pdf(input_path, output_path)
        elif ext == '.xlsx':
            convert_xlsx_to_pdf(input_path, output_path)
        elif ext == '.pptx':
            convert_pptx_to_pdf(input_path, output_path)
        elif ext == '.txt':
            convert_txt_to_pdf(input_path, output_path)
        elif ext == '.svg':
            convert_svg_to_pdf(input_path, output_path)
        else:
            raise ValueError(f"Unsupported file format: {ext}")
        
        print(f"{output_path}")
        return output_path
    except Exception as e:
        error_message = f"Error during conversion: {str(e)}"
        print(error_message)
        raise ValueError(error_message)

# if __name__ == '__main__':
#     input_file = '/Users/dev13/desktop/Buscador de Dioses.docx'  
#     output_file = 'sisas.pdf' 
#     convert_file_to_pdf(input_file, output_file)
    
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python convert.py <input_file>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    convert_file_to_pdf(input_file, "")
    