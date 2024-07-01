import pandas as pd
from wordreader import WordReader
from memoryreader import MemoryReader
from wordmemory import WordMemory
from pngmemory import PngMemory
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class DocumentProcessor:
    def __init__(self, word_path, memory_sheets, output_file):
        self.word_reader = WordReader(word_path)
        self.memory_reader = MemoryReader(memory_sheets, output_file)
        self.word_memory = WordMemory()
        self.png_memory = PngMemory()

    def process_document(self):
        # Leer documento word base
        doc = self.word_reader.read()  # Utiliza un método para leer el documento como objeto Document de python-docx        
        # Reemplazar valores en los párrafos de texto
        self.replace_text(doc)
        
        sheetnames = {'DimensionSistema':'TablaDimensionSistema',
                      'DistCadenas':'TablaDistCadenas'}
        # Reemplazar valores en las tablas
        self.replace_in_tables(doc, sheetnames)    

        # Guardar el documento actualizado
        output_docx = 'PE1126_updated.docx'
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

    def replace_in_tables(self, doc, sheetnames):
        for sheetname, placeholder in sheetnames.items():
            table = self.memory_reader.get_tables(sheetname)
            
            if table is not None:
                # Search and replace placeholder with table
                for paragraph in doc.paragraphs:
                    if placeholder in paragraph.text:
                        # Clear existing content in the paragraph runs
                        for run in paragraph.runs:
                            run.text = ''
                        
                        # Insert table after the paragraph
                        paragraph._element.getparent().insert(paragraph._element.getparent().index(paragraph._element) + 1, table)