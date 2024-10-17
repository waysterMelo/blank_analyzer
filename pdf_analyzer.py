import cv2
import numpy as np
import re
from PIL import Image, ImageEnhance, ImageFilter
import io
import pytesseract
from spellchecker import SpellChecker


class PDFAnalyzer:
    def __init__(self, min_text_length=10, pixel_threshold=0.985, language='eng+por'):
        """
        Inicializa o analisador com parâmetros para OCR e métricas.

        Args:
            min_text_length (int): Comprimento mínimo do texto para considerar OCR bem-sucedido.
            pixel_threshold (float): Limite de proporção de pixels brancos para considerar a página em branco.
            language (str): Idiomas para o Tesseract OCR.
        """
        print("Inicializando PDFAnalyzer...")
        self.min_text_length = min_text_length
        self.pixel_threshold = pixel_threshold
        self.language = language
        # Contadores para rastrear vários status de página
        self.pages_blank_count = 0
        self.pages_blank_after_ocr_count = 0
        self.pages_ocr_analyzed_count = 0
        self.total_characters = 0
        self.correct_characters = 0
        self.total_words = 0
        self.correct_words = 0
        # Inicializa o corretor ortográfico
        self.spell = SpellChecker(language='pt')  # Ajuste o idioma conforme necessário
        print("PDFAnalyzer inicializado com sucesso.")

    def is_blank_or_noisy(self, image):
        """
        Determina se a imagem é em branco ou ruidosa.
        Retorna:
            is_blank (bool), white_pixel_percentage (float), cropped_image (PIL.Image)
        """
        print("Verificando se a imagem é em branco ou ruidosa...")

        # Recorta 5% de cada lado para remover bordas potencialmente ruidosas
        width, height = image.size
        crop_percent = 0.05
        left = int(width * crop_percent)
        right = int(width * (1 - crop_percent))
        cropped_image = image.crop((left, 0, right, height))
        print(f"Imagem cortada para remover bordas: {left}px à {right}px")

        # Converte a imagem recortada para escala de cinza
        gray_image = cv2.cvtColor(np.array(cropped_image), cv2.COLOR_RGB2GRAY)
        print("Imagem convertida para escala de cinza.")

        # Aplica limiarização adaptativa para binarizar a imagem
        binary_image = cv2.adaptiveThreshold(
            gray_image, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            blockSize=15,
            C=10
        )
        print("Binarização adaptativa aplicada.")

        # Calcula a porcentagem de pixels brancos na imagem binarizada
        white_pixel_percentage = np.mean(binary_image == 255)
        print(f"Proporção de pixels brancos: {white_pixel_percentage:.2%}")

        # Determina se a página é considerada em branco com base na porcentagem de pixels brancos
        is_blank = white_pixel_percentage >= self.pixel_threshold
        print(f"Imagem é em branco: {is_blank}")

        return is_blank, white_pixel_percentage, cropped_image

    def perform_ocr_and_reclassify(self, cropped_image):
        """
        Realiza OCR na imagem recortada.
        Retorna:
            ocr_successful (bool), cleaned_text (str)
        """
        print("Iniciando o processo de OCR e reclassificação...")

        try:
            # Aplica filtro mediano para reduzir o ruído na imagem
            cropped_image = cropped_image.filter(ImageFilter.MedianFilter(size=3))
            print("Filtro mediano aplicado para reduzir ruído.")

            # Aumenta o contraste e a nitidez para melhorar a precisão do OCR
            cropped_image = ImageEnhance.Contrast(cropped_image).enhance(3.0)
            print("Contraste da imagem aumentado.")
            cropped_image = ImageEnhance.Sharpness(cropped_image).enhance(2.5)
            print("Nitidez da imagem aumentada.")

            # Ajusta o DPI e converte para preto e branco para melhores resultados de OCR
            with io.BytesIO() as output:
                cropped_image.save(output, format="PNG", dpi=(600, 600))
                output.seek(0)
                with Image.open(output) as image_dpi:
                    image_bw = image_dpi.convert('L')
                    image_bw = ImageEnhance.Contrast(image_bw).enhance(2.0)
                    # Converte a imagem para binária (preto e branco) usando um limiar
                    image_bw = image_bw.point(lambda x: 0 if x < 140 else 255, '1')
                    print("Imagem convertida para preto e branco para OCR.")
                    # Configuração do Tesseract para OCR
                    custom_config = r'--oem 3 --psm 6'
                    text = pytesseract.image_to_string(image_bw, lang=self.language, config=custom_config)
                    print("OCR realizado com Tesseract.")

            # Limpa o texto extraído removendo caracteres indesejados, mas preserva espaços para correção
            text = re.sub(r'[^A-Za-z0-9À-ÿ\s]', ' ', text)
            # Substitui múltiplos espaços por um único espaço
            text = re.sub(r'\s+', ' ', text).strip()
            print(f"Texto extraído pelo OCR: {text[:50]}... (truncado)" if len(
                text) > 50 else f"Texto extraído pelo OCR: {text}")

            # Remove linhas, manchas e ruídos
            lines = text.splitlines()
            cleaned_lines = [line for line in lines if len(line.strip()) > 1 and not re.match(r'^[\W_]+$', line)]
            cleaned_text = ' '.join(cleaned_lines)
            print(f"Texto limpo após remoção de linhas e ruídos: {cleaned_text[:50]}... (truncado)" if len(
                cleaned_text) > 50 else f"Texto limpo após remoção de linhas e ruídos: {cleaned_text}")

            # Realiza correção ortográfica
            corrected_text = self.correct_spelling(cleaned_text)
            print(f"Texto após correção ortográfica: {corrected_text[:50]}... (truncado)" if len(
                corrected_text) > 50 else f"Texto após correção ortográfica: {corrected_text}")

            # Determina se o OCR foi bem-sucedido com base no comprimento do texto limpo
            ocr_successful = len(corrected_text) >= self.min_text_length
            print(f"OCR foi bem-sucedido: {ocr_successful}")
            return ocr_successful, corrected_text

        except pytesseract.TesseractError as e:
            print(f"Erro no OCR: {e}")
            return False, ""

    def correct_spelling(self, text):
        """
        Corrige erros ortográficos no texto utilizando SpellChecker.
        """
        print("Iniciando correção ortográfica...")
        words = text.split()
        corrected_words = []
        for word in words:
            # Verifica se a palavra está correta
            if word.lower() in self.spell:
                corrected_words.append(word)
            else:
                # Sugere correções
                correction = self.spell.correction(word)
                if correction:
                    corrected_words.append(correction)
                    print(f"Corrigido '{word}' para '{correction}'")
                else:
                    corrected_words.append(word)
                    print(f"Nenhuma correção encontrada para '{word}', mantendo original.")
        corrected_text = ' '.join(corrected_words)
        print(f"Texto após correção ortográfica: {corrected_text[:50]}... (truncado)" if len(
            corrected_text) > 50 else f"Texto após correção ortográfica: {corrected_text}")
        return corrected_text

    def analyze_page(self, img, ground_truth_text=None):
        """
        Analisa a imagem de uma única página do PDF.
        Retorna:
            status (str), white_pixel_percentage (float), ocr_performed (bool), extracted_text (str)
        """
        print("Iniciando análise da página...")
        # Verifica se a página é em branco ou ruidosa
        is_blank, white_pixel_percentage, cropped_img = self.is_blank_or_noisy(img)
        extracted_text = ""
        ocr_performed = False
        status = "OK"  # Status padrão se a página não for em branco

        if white_pixel_percentage > 0.98:
            # Se a proporção de pixels brancos for maior que 98%, marca como "Blank"
            status = "Blank"
            print("Página marcada como em branco devido à alta proporção de pixels brancos.")
            self.pages_blank_count += 1

            # Realiza OCR na imagem recortada para tentar reclassificar a página
            ocr_successful, extracted_text = self.perform_ocr_and_reclassify(cropped_img)
            ocr_performed = True

            # Se o OCR detectar texto relevante, marca como "OK after OCR"
            if ocr_successful:
                status = "OK after reanalysis"
                print("Texto relevante detectado após OCR, página reclassificada como 'OK after reanalysis'.")
                self.pages_ocr_analyzed_count += 1
            else:
                # Caso contrário, marca como "Blank after OCR"
                status = "Blank after reanalysis"
                print("Nenhum texto relevante detectado após OCR, página mantida como 'Blank after reanalysis'.")
                self.pages_blank_after_ocr_count += 1

        print(f"Status da página: {status}")
        return status, white_pixel_percentage, ocr_performed, extracted_text