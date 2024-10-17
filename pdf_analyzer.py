# analyzer.py
import fitz  # PyMuPDF
import cv2
import numpy as np
import re
from PIL import Image, ImageEnhance, ImageFilter
import io
import pytesseract

class PDFAnalyzer:
    def __init__(self, min_text_length=10, pixel_threshold=0.97):
        # Initialize the minimum text length for OCR success and the pixel threshold for blank detection
        self.min_text_length = min_text_length
        self.pixel_threshold = pixel_threshold
        # Counters to keep track of various page statuses
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

        # Crop 5% from each side to remove potential noisy borders
        width, height = image.size
        crop_percent = 0.05
        left = int(width * crop_percent)
        right = int(width * (1 - crop_percent))
        cropped_image = image.crop((left, 0, right, height))
        print(f"Imagem cortada para remover bordas: {left}px à {right}px")

        # Convert the cropped image to grayscale
        gray_image = cv2.cvtColor(np.array(cropped_image), cv2.COLOR_RGB2GRAY)
        print("Imagem convertida para escala de cinza.")

        # Apply histogram equalization to enhance contrast
        equalized_image = cv2.equalizeHist(gray_image)
        print("Equalização de histograma aplicada para aumentar o contraste.")

        # Apply Gaussian Blur to further reduce noise
        blurred_image = cv2.GaussianBlur(equalized_image, (5, 5), 0)
        print("Desfoque Gaussiano aplicado para reduzir ruídos.")

        # Use morphological operations to remove remaining noise
        kernel = np.ones((3, 3), np.uint8)
        morph_image = cv2.morphologyEx(blurred_image, cv2.MORPH_CLOSE, kernel, iterations=2)
        morph_image = cv2.morphologyEx(morph_image, cv2.MORPH_OPEN, kernel, iterations=1)
        print("Operações morfológicas aplicadas para remover ruídos.")

        # Apply adaptive thresholding to binarize the image
        binary_image = cv2.adaptiveThreshold(
            morph_image, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            blockSize=15,
            C=10
        )
        print("Binarização adaptativa aplicada.")

        # Calculate the percentage of white pixels in the binarized image
        white_pixel_percentage = np.mean(binary_image == 255)
        print(f"Proporção de pixels brancos: {white_pixel_percentage:.2%}")

        # Determine if the page is considered blank based on the white pixel percentage
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
            # Apply median filter to reduce noise in the image
            cropped_image = cropped_image.filter(ImageFilter.MedianFilter(size=3))
            print("Filtro mediano aplicado para reduzir ruído.")

            # Enhance contrast and sharpness to improve OCR accuracy
            cropped_image = ImageEnhance.Contrast(cropped_image).enhance(3.0)
            print("Contraste da imagem aumentado.")
            cropped_image = ImageEnhance.Sharpness(cropped_image).enhance(2.5)
            print("Nitidez da imagem aumentada.")

            # Apply additional histogram equalization for better contrast
            image_np = np.array(cropped_image.convert('L'))
            image_eq = cv2.equalizeHist(image_np)
            cropped_image = Image.fromarray(image_eq)
            print("Equalização de histograma adicional aplicada.")

            # Adjust DPI and convert to black and white for better OCR results
            with io.BytesIO() as output:
                cropped_image.save(output, format="PNG", dpi=(600, 600))
                output.seek(0)
                with Image.open(output) as image_dpi:
                    image_bw = image_dpi.convert('L')
                    image_bw = ImageEnhance.Contrast(image_bw).enhance(2.5)
                    # Convert image to binary (black and white) using a threshold
                    image_bw = image_bw.point(lambda x: 0 if x < 140 else 255, '1')
                    print("Imagem convertida para preto e branco para OCR.")
                    # Tesseract configuration for OCR
                    custom_config = r'--oem 3 --psm 6'
                    text = pytesseract.image_to_string(image_bw, lang='eng+por', config=custom_config)

            # Clean the extracted text by removing non-alphanumeric characters and extra spaces
            text = re.sub(r'[^A-Za-z0-9À-ÿ]', ' ', text)
            cleaned_text = re.sub(r'\s+', '', text)
            print(f"Texto extraído pelo OCR sem espaços: {cleaned_text[:50]}... (truncado)" if len(
                cleaned_text) > 50 else f"Texto extraído pelo OCR sem espaços: {cleaned_text}")

            # Determine if OCR was successful based on the length of the cleaned text
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
        # Check if the page is blank or noisy
        is_blank, white_pixel_percentage, cropped_img = self.is_blank_or_noisy(img)
        extracted_text = ""
        ocr_performed = False
        status = "OK"  # Default status if the page is not blank

        if is_blank:
            # Increment the blank page counter
            self.pages_blank_count += 1
            # Perform OCR on the cropped image to reclassify the page
            ocr_successful, extracted_text = self.perform_ocr_and_reclassify(cropped_img)
            ocr_performed = True

            # Get the number of characters from the extracted text
            quantidade_caracteres = len(extracted_text)

            # Classification based on OCR results
            if quantidade_caracteres >= 50:
                # If there are 50 or more characters, consider the page as having content
                status = "OK"
                self.pages_ocr_analyzed_count += 1
            elif (ocr_successful or quantidade_caracteres >= 20) and white_pixel_percentage <= self.pixel_threshold:
                # If OCR was successful or there are 20 or more characters, classify as having content after reanalysis
                status = "Identificado conteúdo após reanálise"
                self.pages_ocr_analyzed_count += 1
            elif 10 <= quantidade_caracteres < 20 and 0.98 <= white_pixel_percentage < 0.99:
                # If there are at least 10 characters and the percentage of white pixels is between 98% and 99%, mark as needing attention
                status = "Precisa de Atenção"
                # Add logic to highlight in the report, e.g., marking as red
                # This part is for illustration and should be integrated with reporting logic
                print("Marcando coluna como vermelha no relatório.")
            elif quantidade_caracteres <= 15 and white_pixel_percentage >= self.pixel_threshold:
                # If there are 15 or fewer characters and a high percentage of white pixels, classify as blank with irrelevant text
                status = "Página em branco após reanálise"
                self.pages_blank_after_ocr_count += 1
            else:
                # Otherwise, classify as blank
                status = "Página em branco"
                self.pages_blank_after_ocr_count += 1

        return status, white_pixel_percentage, ocr_performed, extracted_text
