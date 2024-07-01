import pandas as pd

class MemoryReader:
    def __init__(self,memory_sheets, output_file):
        self.memory_sheets = memory_sheets
        self.output_file = output_file

    def read(self):
            memory_df = pd.read_excel('Variables para automatización memorias de calculo.xlsx', sheet_name='Vars')    
            
            
            
            ### Buscar en la memoria las variables
            for item in self.memory_sheets:
                df = pd.read_excel('PE1126_Cálculos Electricos_CR.xlsx', sheet_name=item)
                ### De aquí en adelante se deben buscar
                ### las variables en la memoria y llevarlas
                ### a la tabla resumen de variables, el return de esta función
                ### debe ser el dataframe con las variables
            
            return memory_df
            