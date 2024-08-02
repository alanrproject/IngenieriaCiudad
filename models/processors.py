import pandas as pd
from wordreader import WordReader
from memoryreader import MemoryReader
from docx.shared import Pt
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx import Document

class DocumentProcessor:
    def __init__(self, word_path, sheetnames, image_dict):
        self.word_reader = WordReader(word_path)
        self.memory_reader = MemoryReader(sheetnames)
        self.sheetnames = sheetnames
        self.image_dict = image_dict

    def process_document(self):
        # Leer documento word base
        doc = self.word_reader.read()  # Utiliza un método para leer el documento como objeto Document de python-docx        
        # Reemplazar valores en los párrafos de texto
        self.replace_text(doc)
        # Reemplazar valores en las tablas
        self.replace_in_tables(doc)    
        # Reemplazar imágenes en el documento
        self.replace_images(doc)



        # Guardar el documento actualizado
        output_docx = 'processed/MemoriadeCalculo_V0.docx'
        doc.save(output_docx)
        print(f"Archivo '{output_docx}' generado con éxito.")

    def replace_text(self, doc):
        paragraphs = doc.paragraphs
        df = pd.DataFrame(self.memory_reader.read())

        # Reemplazar valores en los párrafos de texto
        for _, row in df.iterrows():
            variable_name = row['Nombre de la variable']
            value = row['Valor']
            for paragraph in paragraphs:
                if variable_name in paragraph.text:
                    for run in paragraph.runs:
                        run.font.size = Pt(12)  # Set the font size to 12 point
                        run.text = run.text.replace(variable_name, str(value))

    def replace_in_tables(self, doc):
        for sheetname, placeholder in self.sheetnames.items():
            # Obtener la tabla como un objeto Table de python-docx
            table = self.memory_reader.get_tables(sheetname)
            
            if table is not None:
                # Buscar y reemplazar el marcador de posición con la tabla
                for paragraph in doc.paragraphs:
                    if placeholder in paragraph.text:
                        # Limpiar el contenido existente en las ejecuciones del párrafo
                        for run in paragraph.runs:
                            run.text = ''
                        
                        # Insertar la tabla después del párrafo
                        parent_element = paragraph._element.getparent()
                        insert_index = parent_element.index(paragraph._element) + 1
                        
                        # Convertir la tabla a XML y agregarla al documento
                        tbl_xml = table._tbl  # Obtener el elemento XML de la tabla
                        parent_element.insert(insert_index, tbl_xml)
                        
                        # Opcionalmente eliminar el marcador de posición del párrafo
                        paragraph.clear()
                        
                        # Salir del bucle después de reemplazar el primer marcador de posición
                        break

    def replace_images(self, doc):
        for placeholder, image_path in self.image_dict.items():
            for paragraph in doc.paragraphs:
                if placeholder in paragraph.text:
                    # Create a new paragraph for the image
                    new_paragraph = paragraph.insert_paragraph_before()
                    new_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Center align the paragraph
                    run = new_paragraph.add_run()
                    run.add_picture(image_path, width=Inches(4))  # Adjust the width as needed


