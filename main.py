# main.py
from tkinter import Tk
from gui import PDFAnalyzerGUI
from tesseract_config import TesseractConfig


def main():
    print("Iniciando aplicação...")
    root = Tk()
    root.withdraw()

    # Configure Tesseract OCR
    tessdata_prefix = r'C:/Program Files/Tesseract-OCR/tessdata/'
    tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
    tesseract_config = TesseractConfig(tessdata_prefix, tesseract_cmd)
    tessdata_prefix_env = tessdata_prefix  # Ensure this path exists
    tesseract_config.test_setup()

    root.destroy()

    # Launch GUI
    app = PDFAnalyzerGUI()


if __name__ == "__main__":
    main()
