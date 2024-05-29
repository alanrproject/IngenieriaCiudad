from wordreader import wordreader
from memoryreader import memoryreader
from wordmemory import wordmemory
from pngmemory import pngmemory

wordpath = 'PE1126_Memorias de cálculo.docx'
memorysheets = ['PV_INV_STRING','PV_INV_PARALELO','INV1_TABLERO FV','INV2_TABLERO FV','TABLERO FV_INTX']

def main():
    print(wordreader(wordpath))
    memoryreader(memorysheets)
    # wordmemory()
    # pngmemory()

if __name__ == "__main__":
     main()