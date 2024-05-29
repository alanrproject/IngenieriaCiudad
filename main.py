from processors import DocumentProcessor

def main():
    word_path = 'PE1126_Memorias de cálculo.docx'
    memory_sheets = ['PV_INV_STRING', 'PV_INV_PARALELO', 'INV1_TABLERO FV', 'INV2_TABLERO FV', 'TABLERO FV_INTX']
    output_file = 'output.xlsx'  # Specify the output Excel file path
    
    processor = DocumentProcessor(word_path, memory_sheets, output_file)
    processor.process_document()

if __name__ == "__main__":
    main()
