from wordreader import WordReader
from memoryreader import MemoryReader
from wordmemory import WordMemory
from pngmemory import PngMemory

class DocumentProcessor:
    def __init__(self, word_path, memory_sheets, output_file):
        self.word_reader = WordReader(word_path)
        self.memory_reader = MemoryReader(memory_sheets, output_file)
        self.word_memory = WordMemory()
        self.png_memory = PngMemory()

    def process_document(self):
        text, headings, images = self.word_reader.read()
        print("Text:", text)
        print("Headings:", headings)
        print("Images:", images)
        
        self.memory_reader.read()
        self.word_memory.process(text, headings)
        self.png_memory.process(images)