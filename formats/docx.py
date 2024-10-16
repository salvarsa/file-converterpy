import pypandoc
import pdfkit

def convert_docx_to_pdf(input_path: str, output_path: str) -> None:
    output_path = output_path if output_path.endswith('.pdf') else output_path + '.pdf'
    try:
        pypandoc.convert_file(input_path, 'pdf', outputfile=output_path, 
                              extra_args=['--pdf-engine=/Library/TeX/texbin/pdflatex'])
    except Exception as e:
        print(f"CONVERTION_ERROR: {e}")
        try:
            html_content = pypandoc.convert_file(input_path, 'html')
            pdfkit.from_string(html_content, output_path)
        except Exception as e:
            print(f"HTML_ERROR: {e}")
            raise