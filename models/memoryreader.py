import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table
from docx.oxml import OxmlElement
from docx import Document
from docx.shared import Pt
from lxml import etree
from docx.oxml.ns import qn  # Asegúrate de importar qn para los nombres cualificados

class MemoryReader:
    def __init__(self,sheetnames):
        self.memory_sheets = sheetnames

    def read(self):
            memory_df = pd.read_excel('raw/MemoriadeCalculo.xlsx', sheet_name='Vars')    
            return memory_df
            
    def get_tables(self, sheetname):
        # Cargar el libro de trabajo y la hoja
        wb = load_workbook('raw/MemoriadeCalculo.xlsx', data_only=True)
        ws = wb[sheetname]
        
        # Encontrar la primera tabla con formato en la hoja
        table = next((tbl for tbl in ws.tables.values() if isinstance(tbl, Table)), None)
        
        if table:
            # Convertir el rango de la tabla a DataFrame
            data_range = ws[table.ref]
            data = []
            
            # Extraer los valores de las celdas
            for row in data_range:
                row_values = [cell.value for cell in row]
                data.append(row_values)
            
            # Convertir la lista de listas a DataFrame
            df = pd.DataFrame(data[1:], columns=data[0])  # Usar la primera fila como encabezado
            return self.df_to_table(df)  # Convertir el DataFrame a XML para Word
        else:
            raise ValueError("No se encontró una tabla con formato en la hoja.")

    def df_to_table(self, df):
        # Crear un nuevo documento y tabla
        doc = Document()
        tbl = doc.add_table(rows=1, cols=len(df.columns))
        
        # Aplicar el estilo 'Grid Table 1 Light' al objeto Table de python-docx
        tbl.style = 'Table Grid'
        
        # Agregar la fila de encabezado de la tabla
        hdr_cells = tbl.rows[0].cells
        for i, header_text in enumerate(df.columns):
            hdr_cells[i].text = str(header_text)
            # Formato del encabezado
            hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(10)
            hdr_cells[i].paragraphs[0].style = doc.styles['Normal']
        
        # Agregar las filas y celdas de datos
        for _, row_data in df.iterrows():
            row_cells = tbl.add_row().cells
            for i, value in enumerate(row_data):
                row_cells[i].text = str(value)
                # Formato de las celdas de datos
                row_cells[i].paragraphs[0].runs[0].font.size = Pt(8)
        
        return tbl