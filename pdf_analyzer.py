# analyzer.py
import fitz  # PyMuPDF
import cv2
import numpy as np
import re
from PIL import Image, ImageEnhance, ImageFilter
import io
import pytesseract

class PDFAnalyzer:
    def __init__(self, min_text_length=10, pixel_threshold=0.98):
        self.min_text_length = min_text_length
        self.pixel_threshold = pixel_threshold
        self.pages_blank_count = 0
        self.pages_blank_after_ocr_count = 0
        self.pages_ocr_analyzed_count = 0

    def is_blank_or_noisy(self, image):
        """
        Determines if the image is blank or noisy.
        Returns:
            is_blank (bool), white_pixel_percentage (float), cropped_image (PIL.Image)
        """
        print("Verificando se a imagem é em branco ou ruidosa...")

        # Crop 5% from each side
        width, height = image.size
        crop_percent = 0.05
        left = int(width * crop_percent)
        right = int(width * (1 - crop_percent))
        cropped_image = image.crop((left, 0, right, height))
        print(f"Imagem cortada para remover bordas: {left}px à {right}px")

        # Convert to grayscale
        gray_image = cv2.cvtColor(np.array(cropped_image), cv2.COLOR_RGB2GRAY)
        print("Imagem convertida para escala de cinza.")

        # Histogram equalization
        equalized_image = cv2.equalizeHist(gray_image)
        print("Equalização de histograma aplicada para aumentar o contraste.")

        # Morphological operations to remove noise
        kernel = np.ones((3, 3), np.uint8)
        morph_image = cv2.morphologyEx(equalized_image, cv2.MORPH_CLOSE, kernel, iterations=2)
        morph_image = cv2.morphologyEx(morph_image, cv2.MORPH_OPEN, kernel, iterations=1)
        print("Operações morfológicas aplicadas para remover ruídos.")

        # Adaptive thresholding
        binary_image = cv2.adaptiveThreshold(
            morph_image, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            blockSize=15,
            C=10
        )
        print("Binarização adaptativa aplicada.")

        # Calculate white pixel percentage
        white_pixel_percentage = np.mean(binary_image == 255)
        print(f"Proporção de pixels brancos: {white_pixel_percentage:.2%}")

        # Determine if the page is blank
        is_blank = white_pixel_percentage >= self.pixel_threshold
        print(f"Imagem é em branco: {is_blank}")

        return is_blank, white_pixel_percentage, cropped_image

    def perform_ocr_and_reclassify(self, cropped_image):
        """
        Performs OCR on the cropped image.
        Returns:
            ocr_successful (bool), cleaned_text (str)
        """
        print("Iniciando o processo de OCR e reclassificação...")

        try:
            # Apply median filter to reduce noise
            cropped_image = cropped_image.filter(ImageFilter.MedianFilter(size=3))
            print("Filtro mediano aplicado para reduzir ruído.")

            # Enhance contrast and sharpness
            cropped_image = ImageEnhance.Contrast(cropped_image).enhance(3.0)
            print("Contraste da imagem aumentado.")
            cropped_image = ImageEnhance.Sharpness(cropped_image).enhance(2.5)
            print("Nitidez da imagem aumentada.")

            # Additional histogram equalization
            image_np = np.array(cropped_image.convert('L'))
            image_eq = cv2.equalizeHist(image_np)
            cropped_image = Image.fromarray(image_eq)
            print("Equalização de histograma adicional aplicada.")

            # Adjust DPI and convert to black and white
            with io.BytesIO() as output:
                cropped_image.save(output, format="PNG", dpi=(600, 600))
                output.seek(0)
                with Image.open(output) as image_dpi:
                    image_bw = image_dpi.convert('L')
                    image_bw = ImageEnhance.Contrast(image_bw).enhance(2.0)
                    image_bw = image_bw.point(lambda x: 0 if x < 140 else 255, '1')
                    print("Imagem convertida para preto e branco para OCR.")
                    # Tesseract configuration
                    custom_config = r'--oem 3 --psm 6'
                    text = pytesseract.image_to_string(image_bw, lang='eng+por', config=custom_config)

            # Clean extracted text
            text = re.sub(r'[^A-Za-z0-9À-ÿ]', ' ', text)
            cleaned_text = re.sub(r'\s+', '', text)
            print(f"Texto extraído pelo OCR sem espaços: {cleaned_text[:50]}... (truncado)" if len(
                cleaned_text) > 50 else f"Texto extraído pelo OCR sem espaços: {cleaned_text}")

            ocr_successful = len(cleaned_text) >= self.min_text_length
            return ocr_successful, cleaned_text

        except pytesseract.TesseractError as e:
            print(f"Erro no OCR: {e}")
            return False, ""

    def analyze_page(self, img):
        """
        Analyzes a single PDF page image.
        Returns:
            status (str), white_pixel_percentage (float), ocr_performed (bool), extracted_text (str)
        """
        is_blank, white_pixel_percentage, cropped_img = self.is_blank_or_noisy(img)
        extracted_text = ""
        ocr_performed = False
        status = "OK"  # Define o status padrão como "OK" caso a página não seja em branco

        if is_blank:
            self.pages_blank_count += 1
            ocr_successful, extracted_text = self.perform_ocr_and_reclassify(cropped_img)
            ocr_performed = True

            # Classificação com base nos resultados do OCR
            if (ocr_successful or len(extracted_text) >= 40) and white_pixel_percentage <= self.pixel_threshold:
                status = "Identificado conteúdo após reanálise"
                self.pages_ocr_analyzed_count += 1
            elif len(extracted_text) >= 15 and white_pixel_percentage >= self.pixel_threshold:
                status = "Página em branco com texto irrelevante"
                self.pages_blank_after_ocr_count += 1
            else:
                status = "Página em branco"
                self.pages_blank_after_ocr_count += 1

        return status, white_pixel_percentage, ocr_performed, extracted_text
