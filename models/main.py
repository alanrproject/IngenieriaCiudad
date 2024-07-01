from processors import DocumentProcessor

def main():
    word_path = 'raw/PE1126_Memorias de cálculo.docx'
    sheetnames = {'DimensionSistema':'TablaDimensionSistema',
                      'DistCadenas':'TablaDistCadenas'}
    image_dict = {'Imagen 1: Vista aérea':'raw/vistaaerea.png'}  # Specify the output Excel file path
    
    processor = DocumentProcessor(word_path, sheetnames, image_dict)
    processor.process_document()

if __name__ == "__main__":
    main()
