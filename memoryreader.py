import pandas as pd

def memoryreader(memorysheets):
    for item in memorysheets:
        df = pd.read_excel('PE1126_Cálculos Electricos_CR.xlsx',sheet_name=item)
        
        return df.to_excel()