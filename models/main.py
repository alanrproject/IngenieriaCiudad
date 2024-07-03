from processors import DocumentProcessor

def main():
    word_path = 'raw/PE1126_Memorias de cálculo.docx'
    # Diccionario de las tablas que se van a incrustar en la memoria
    sheetnames = {'DimensionSistema':'TablaDimensionSistema',
                      'DistCadenas':'TablaDistCadenas'}
    # Diccionario de las imágenes que se van a incrustar en la memoria
    image_dict = {'Imagen 1: Vista aérea':'raw/vistaaerea.png',
                  'Imagen 2: Temperatura Mensual':'raw/Temperaturamensual.png'}
    
    processor = DocumentProcessor(word_path, sheetnames, image_dict)
    processor.process_document()

if __name__ == "__main__":
    main()
