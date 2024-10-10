import pypandoc

#conversor TXT   
def convert_txt_to_pdf(input_path: str, output_path: str) -> None:
    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()
    pypandoc.convert_text(content, 'pdf', format='markdown', outputfile=output_path)