import pandas as pd
from wordreader import WordReader
from memoryreader import MemoryReader
from wordmemory import WordMemory
from pngmemory import PngMemory
from docx import Document

class DocumentProcessor:
    def __init__(self, word_path, memory_sheets, output_file):
        self.word_reader = WordReader(word_path)
        self.memory_reader = MemoryReader(memory_sheets, output_file)
        self.word_memory = WordMemory()
        self.png_memory = PngMemory()

    def process_document(self):
        # Leer documento word base
        text, headings, images = self.word_reader.read()
        print("Text:", text)
        print("Headings:", headings)
        print("Images:", images)
        
        df = pd.DataFrame(self.memory_reader.read())
        print(df)
        
        # Reemplazar valores en el texto
        for _, row in df.iterrows():
            variable_name = row['Nombre de la variable']
            value = row['Valor']
            text = text.replace(variable_name, str(value))
        
        print("Updated Text:", text)

        # Guardar el texto actualizado en un nuevo archivo DOCX
        output_docx = 'PE1126_updated.docx'
        doc = Document()
        doc.add_paragraph(text)  # Agrega el texto actualizado como un párrafo en el documento
        doc.save(output_docx)
        print(f"Archivo '{output_docx}' generado con éxito.")



        ### Aquí se debe continuar la incrustación de las variables en df
        ### en el objeto text
        
        # self.memory_reader.read()
        # self.word_memory.process(text, headings)
        # self.png_memory.process(images)