# utils.py
import os
import sys
import pytesseract
from tkinter import messagebox

class TesseractConfig:
    def __init__(self, tessdata_path, tesseract_cmd):
        self.tessdata_prefix = tessdata_path
        self.tesseract_cmd = tesseract_cmd
        self.configure_tesseract()

    def configure_tesseract(self):
        os.environ['TESSDATA_PREFIX'] = self.tessdata_prefix
        pytesseract.pytesseract.tesseract_cmd = self.tesseract_cmd

    def test_setup(self):
        try:
            print("Verificando a configuração do Tesseract OCR...")
            tessdata_prefix_env = os.environ.get('TESSDATA_PREFIX')
            if not tessdata_prefix_env or not os.path.isdir(tessdata_prefix_env):
                raise EnvironmentError("TESSDATA_PREFIX não está configurado corretamente.")
            tesseract_version = pytesseract.get_tesseract_version()
            print(f"Tesseract OCR instalado corretamente. Versão: {tesseract_version}")
            messagebox.showinfo("Tesseract OCR", f"Tesseract OCR instalado corretamente.\nVersão: {tesseract_version}")
        except Exception as e:
            print(f"Erro ao inicializar o Tesseract OCR: {e}")
            messagebox.showerror("Erro Tesseract OCR", f"Erro ao inicializar o Tesseract OCR.\n{str(e)}")
            sys.exit(1)
