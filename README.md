md

# File Converter

Este proyecto permite convertir diferentes formatos de archivo a PDF, incluyendo `.docx`, `.xlsx`, `.pptx`, `.txt` y algunos formatos de imagen.

## Requisitos del sistema

- **Python 3.10+**
- **pip** y **venv**

### Instalación de pre-requisitos

```bash
sudo apt update
sudo apt install python3-pip python3-venv

Instrucciones para la instalación y ejecución
1. Clonar el repositorio

bash

git clone https://github.com/salvarsa/file-converterpy.git
cd file-converterpy

2. Crear y activar el entorno virtual

bash

python3 -m venv venv
source venv/bin/activate

3. Instalar dependencias

Instala las bibliotecas necesarias:

bash

pip install -r requirements.txt

Si no tienes el archivo requirements.txt, las dependencias son:

bash

pip install pypandoc pdfkit pillow reportlab openpyxl python-pptx svglib

Además, asegúrate de tener instalados los motores de conversión adicionales:

bash

sudo apt install pandoc texlive texlive-latex-extra texlive-xetex wkhtmltopdf

4. Ejecutar el proyecto

bash

python3 convert.py <input_file>

Esto convertirá el archivo de entrada en un PDF y lo guardará en la carpeta de descargas.
Estructura del proyecto

├── convert.py
├── formats
│   ├── docx.py
│   ├── images.py
│   ├── pptx.py
│   └── xlxs.py
├── requirements.txt
└── venv

Sugerencias

    Crea un archivo requirements.txt con las dependencias.
    Añade scripts de manejo de errores para dependencias faltantes.