import os, fnmatch
from zipfile import ZipFile
from xml.etree import ElementTree
from datetime import datetime
from tkinter import filedialog

def Document(inputfile: str) -> str:
    if not os.path.exists(inputfile):
        raise Exception('File does not exist!')

    os.makedirs(out_dir)
    os.makedirs(img_dir)

    paragraphs = []
    with ZipFile(inputfile, 'r') as zf:
        for file in zf.infolist():
             if file.filename.startswith('word/media/'):
                file_name = os.path.basename(file.filename)  # Получаем имя файла без пути
                output_path = os.path.join(img_dir, file_name)
                with zf.open(file) as file_:
                    with open(output_path, 'wb') as output_file:
                        output_file.write(file_.read())

        print("images extract success")
        with zf.open('word/document.xml') as doc:
            tree = ElementTree.parse(doc)
            root = tree.getroot()
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
            number_image = 0
            for paragraph in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
                text_elements = paragraph.findall('.//w:t', namespaces=ns)
                paragraph_text = ''.join([text.text for text in text_elements if text.text is not None])

                images = paragraph.findall('.//wp:docPr', namespaces=ns)
                if images:
                    number_image += 1
                    paragraphs.append(f'<img class=listing" alt="картинка" src=image{number_image}.png>')
                    continue

                if paragraph_text:
                    paragraphs.append('<p>'+ paragraph_text + '</p>')

        print("text extract success")

    return '\n'.join(paragraphs) if paragraphs else 'No paragraphs found'

inputfile = filedialog.Open().show()
#timestamp + filename
out_dir = f'{datetime.now().strftime("%Y-%m-%d_%H-%M-%S")}_{inputfile.split("/")[-1]}'
img_dir = f'{out_dir}/img'
html_output_file = 'index.html'
output = Document(inputfile)
with open(os.path.join(out_dir, html_output_file), 'w', encoding='utf-8') as f:
    f.write(output)
