import os
import sys
from formats.docx import convert_docx_to_pdf
from formats.xlxs import convert_xlsx_to_pdf
from formats.pptx import convert_pptx_to_pdf
from formats.images import convert_image_to_pdf
from formats.txt import convert_txt_to_pdf


def download_pdf():
    return os.path.join(os.path.expanduser('~'), 'Downloads')

def generate_unique_filename(base_path, base_name, extension):
    counter = 1
    new_path = os.path.join(base_path, f"{base_name}{extension}")
    while os.path.exists(new_path):
        new_path = os.path.join(base_path, f"{base_name}_{counter}{extension}")
        counter += 1
    return new_path

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
        elif ext in ['.svg', '.png', '.jpg', '.jpeg']:
            convert_image_to_pdf(input_path, output_path)
        else:
            raise ValueError(f"INVALID_FILE_FORMAT: {ext}")
        
        print(f"{output_path}")
        return output_path
    except Exception as e:
        error_message = f"CONVERTION_ERROR: {str(e)}"
        print(error_message)
        raise ValueError(error_message)
    
if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit(1)
    
    input_file = sys.argv[1]
    convert_file_to_pdf(input_file, "")