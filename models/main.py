from processors import DocumentProcessor

def main():
    word_path = 'raw/PE1126_Memorias de cálculo.docx'
    sheetnames = {'DimensionSistema':'TablaDimensionSistema',
                      'DistCadenas':'TablaDistCadenas'}
    image_dict = {'Imagen 1: Vista aérea':'raw/vistaaerea.png', #Specify the output Excel file path # Specify the output Excel file path
                  'Imagen 2: Temperatura promedio mensual':'raw/temperaturapromedio.png',
                  'Imagen 3: Caracteristícas eléctricas del módulo':'raw/caracteristicasdelmodulo.png', 
                  'Imagen 4: Tabla 310-16 de la NTC 2050':'raw/tabla31016.png',
                  'Imagen 5: Tabla 250-95 de la NTC 2050':'raw/tabla25095.png',
                  'Imagen 6: Características técnicas del inversor':'raw/caracteristicasdelinversor.png',
                  'Imagen 7: Tabla 310-16 de la NTC 2050':'raw/tabla31016.png',
                  'Imagen 8: Tabla 250-95 de la NTC 2050':'raw/tabla25095.png',
                  'Imagen 9: Tabla 1 del capítulo 9 de la NTC 2050':'raw/porcentajellenadodetuberias.png'}
                 
    
    processor = DocumentProcessor(word_path, sheetnames, image_dict)
    processor.process_document()

if __name__ == "__main__":
    main()



