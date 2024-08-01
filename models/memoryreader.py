import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table
from docx.oxml import OxmlElement
from docx import Document
from docx.shared import Pt

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
        # Crear un nuevo elemento de tabla
        tbl = OxmlElement('w:tbl')

        # Agregar la cuadrícula de la tabla
        tbl_grid = OxmlElement('w:tblGrid')
        for _ in range(len(df.columns)):
            grid_col = OxmlElement('w:gridCol')
            tbl_grid.append(grid_col)
        tbl.append(tbl_grid)

        # Agregar la fila de encabezado de la tabla
        tr_header = OxmlElement('w:tr')
        for header_text in df.columns:
            tc_header = OxmlElement('w:tc')
            p_header = OxmlElement('w:p')
            run_header = OxmlElement('w:r')
            run_header_t = OxmlElement('w:t')
            run_header_t.text = str(header_text)
            run_header.append(run_header_t)
            p_header.append(run_header)
            tc_header.append(p_header)
            tr_header.append(tc_header)
        tbl.append(tr_header)

        # Agregar las filas y celdas de datos
        for _, row_data in df.iterrows():
            tr = OxmlElement('w:tr')
            for value in row_data:
                tc = OxmlElement('w:tc')
                p = OxmlElement('w:p')
                run = OxmlElement('w:r')
                run_t = OxmlElement('w:t')
                run_t.text = str(value)
                run.append(run_t)
                p.append(run)
                tc.append(p)
                tr.append(tc)
            tbl.append(tr)

        return tbl
            
            
            