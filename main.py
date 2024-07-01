from processors import DocumentProcessor

def main():
    word_path = 'PE1126_Memorias de cálculo.docx'
    sheetnames = {'DimensionSistema':'TablaDimensionSistema',
                      'DistCadenas':'TablaDistCadenas'}
    output_file = 'output.xlsx'  # Specify the output Excel file path
    
    processor = DocumentProcessor(word_path, sheetnames)
    processor.process_document(sheetnames)

if __name__ == "__main__":
    main()
