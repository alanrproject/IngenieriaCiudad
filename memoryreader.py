import pandas as pd

def MemoryReader(memorysheets, output_file):
    with pd.ExcelWriter(output_file) as writer:  # Create Excel writer object
        for item in memorysheets:
            df = pd.read_excel('PE1126_Cálculos Electricos_CR.xlsx', sheet_name=item)
            df.to_excel(writer, sheet_name=item)  # Write DataFrame to Excel writer object