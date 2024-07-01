import pandas as pd
from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Pt

class MemoryReader:
    def __init__(self,memory_sheets, output_file):
        self.memory_sheets = memory_sheets
        self.output_file = output_file

    def read(self):
            memory_df = pd.read_excel('Variables para automatización memorias de calculo.xlsx', sheet_name='Vars')    
            return memory_df
            
    def get_tables(self, sheetname):
        df = pd.read_excel('PE1126_Cálculos Electricos_CR.xlsx', sheet_name=sheetname)
        table = self.df_to_table(df)
        return table
    
    def df_to_table(self, df):
        # Create a new table element
        tbl = OxmlElement('w:tbl')
        
        # Add table grid (necessary for proper table structure)
        tbl_grid = OxmlElement('w:tblGrid')
        for _ in range(len(df.columns)):
            grid_col = OxmlElement('w:gridCol')
            tbl_grid.append(grid_col)
        tbl.append(tbl_grid)
        
        # Add table header row
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
        
        # Add table rows and cells for data
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
            
            
            