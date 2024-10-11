import fitz  # PyMuPDF
import cv2
import numpy as np
import os
from tkinter import Tk, Label, Button, filedialog, messagebox, StringVar, ttk, Canvas, PhotoImage, CENTER
import threading
import subprocess
import platform
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from ttkthemes import ThemedTk
from PIL import Image, ImageTk, ImageEnhance
import io
import sys

# Definir o TESSDATA_PREFIX corretamente antes de importar o pytesseract
tessdata_prefix = r'C:/Program Files/Tesseract-OCR/tessdata/'
os.environ['TESSDATA_PREFIX'] = tessdata_prefix

import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

def test_tesseract_setup():
    print("Testando configuração do Tesseract...")
    try:
        tessdata_prefix_env = os.environ.get('TESSDATA_PREFIX')
        if not tessdata_prefix_env:
            raise Exception("A variável de ambiente TESSDATA_PREFIX não está definida.")
        if not os.path.isdir(tessdata_prefix_env):
            raise Exception(f"O diretório especificado em TESSDATA_PREFIX não existe: {tessdata_prefix_env}")

        print("Tesseract OCR configurado corretamente.")
        tesseract_version = pytesseract.get_tesseract_version()
        print(f"Versão: {tesseract_version}")
        messagebox.showinfo("Tesseract OCR", f"Tesseract OCR está instalado corretamente.\nVersão: {tesseract_version}")
    except Exception as e:
        print(f"Erro ao inicializar o Tesseract OCR: {str(e)}")
        messagebox.showerror("Erro Tesseract OCR", f"Erro ao inicializar o Tesseract OCR.\n{str(e)}")
        sys.exit(1)

def sanitize_filename(filename):
    print(f"Sanitizando nome do arquivo: {filename}")
    normalized_filename = unicodedata.normalize('NFKD', filename).encode('ascii', 'ignore').decode('ascii')
    print(f"Nome do arquivo sanitizado: {normalized_filename}")
    return normalized_filename

def is_blank_or_noisy(image, white_threshold=215, pixel_threshold=0.98):
    print("Determinando se a página está em branco ou ruidosa...")
    if image is None:
        print("Imagem é None.")
        return False, 0

    resized_image = cv2.resize(image, (image.shape[1] // 2, image.shape[0] // 2))
    white_pixel_percentage = np.mean(resized_image > white_threshold)
    print(f"Porcentagem de pixels brancos: {white_pixel_percentage}")
    return white_pixel_percentage > pixel_threshold, white_pixel_percentage

def perform_ocr_and_reclassify(image, min_text_length=2, is_blank=False):
    print("Executando OCR para verificar texto relevante...")
    try:
        with io.BytesIO() as output:
            if is_blank:
                image.save(output, format="PNG", dpi=(1200, 1200))
            output.seek(0)
            with Image.open(output) as image_dpi:
                print("Aumentando DPI da imagem para melhor análise OCR.")
                enhancer = ImageEnhance.Contrast(image_dpi)
                image_contrast = enhancer.enhance(2.0)
                sharpness = ImageEnhance.Sharpness(image_contrast)
                image_sharp = sharpness.enhance(2.0)
                image_bw = image_sharp.convert('L').point(lambda x: 0 if x < 128 else 255, '1')
                text = pytesseract.image_to_string(image_bw, lang='eng+por')
        cleaned_text = ''.join(text.split())
        print(f"Comprimento do texto extraído: {len(cleaned_text)}")
        return (True, "precisa de revisão") if len(cleaned_text) > min_text_length else (False, "Sem Texto")
    except pytesseract.TesseractError as e:
        print(f"Erro OCR: {str(e)}")
        return False, "Erro OCR"

def analyze_single_pdf(pdf_path, ws, total_pages_processed, total_pages, progress_var, progress_label, progress_bar):
    print(f"Analisando PDF: {pdf_path}")
    with fitz.open(pdf_path) as pdf_document:
        total_pdf_pages = pdf_document.page_count
        pdf_name = sanitize_filename(os.path.basename(pdf_path))
        print(f"Total de páginas no PDF: {total_pdf_pages}")

        for page_num in range(total_pdf_pages):
            print(f"Processando página {page_num + 1}/{total_pdf_pages}")
            try:
                page = pdf_document.load_page(page_num)
                pix = page.get_pixmap()
                image_bytes = pix.tobytes("png")

                with Image.open(io.BytesIO(image_bytes)) as img:
                    width, height = img.size
                    left = width * 0.10
                    top = height * 0.05
                    right = width * 0.90
                    bottom = height * 0.90
                    img = img.crop((left, top, right, bottom))
                    print(f"Imagem recortada para remover 10% das bordas em todos os lados para a página {page_num + 1}")

                    image_gray = img.convert('L')
                    image_np = np.array(image_gray)
                    print(f"Imagem convertida para escala de cinza e carregada em um array numpy para a página {page_num + 1}")

                is_blank, white_pixel_percentage = is_blank_or_noisy(image_np)
                status = "Em branco ou ruidosa" if is_blank else "OK"
                ocr_text = "Não"

                if is_blank:
                    with Image.open(io.BytesIO(image_bytes)) as img:
                        img = img.crop((left, top, right, bottom))
                        relevant_text_found, text_or_revision = perform_ocr_and_reclassify(img, is_blank=is_blank)
                        status = text_or_revision if relevant_text_found else "Em branco após OCR"
                        ocr_text = "Sim" if relevant_text_found else "Não"

                print(f"Status da página {page_num + 1}: {status}, Porcentagem de Pixels Brancos: {white_pixel_percentage:.2%}")
                ws.append([pdf_name, page_num + 1, status, f"{white_pixel_percentage:.2%}", ocr_text])

            except (FileNotFoundError, ValueError, IOError) as e:
                print(f"Erro ao processar a página {page_num + 1} do arquivo {pdf_name}: {str(e)}")
                ws.append([pdf_name, page_num + 1, f"Erro: {str(e)}", None, None])

            total_pages_processed[0] += 1
            progress_percentage = (total_pages_processed[0] / total_pages) * 100
            print(f"Progresso: {progress_percentage:.2f}%")
            progress_var.set(progress_percentage)
            progress_label.after(0, lambda: progress_label.config(text=f"Analisando... {int(float(progress_var.get()))}%"))
            progress_bar.after(0, progress_bar.update_idletasks)

def select_directory():
    global directory
    directory = filedialog.askdirectory()
    if directory:
        print(f"Diretório selecionado: {directory}")
        select_label.config(text=f"Diretório Selecionado: {directory}")
        analyze_button.config(state="normal")

def start_analysis():
    if directory:
        print("Iniciando análise...")
        analyze_button.config(state="disabled")
        analysis_thread = threading.Thread(target=run_analysis_thread, args=(directory, progress_var, progress_label, progress_bar))
        analysis_thread.start()

def run_analysis_thread(directory, progress_var, progress_label, progress_bar):
    print(f"Iniciando thread de análise para o diretório: {directory}")
    output_xlsx = os.path.join(directory, "analysis_report.xlsx")
    analyze_pdfs_in_directory(directory, progress_var, progress_label, output_xlsx, progress_bar)

def open_folder():
    if directory:
        print(f"Abrindo pasta: {directory}")
        if platform.system() == "Windows":
            os.startfile(directory)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", directory])
        elif platform.system() == "Linux":
            subprocess.Popen(["xdg-open", directory])

def analyze_pdfs_in_directory(pdf_dir, progress_var, progress_label, output_xlsx, progress_bar):
    print(f"Analisando PDFs no diretório: {pdf_dir}")

    wb = Workbook()
    ws = wb.active
    ws.title = "PDF Analysis Report"
    ws.append(["Arquivo PDF", "Página", "Status", "Porcentagem de Pixels Brancos", "Texto Extraído"])

    pdf_files = [os.path.join(pdf_dir, f) for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
    print(f"Total de arquivos PDF encontrados: {len(pdf_files)}")

    total_pages = 0
    for pdf_file in pdf_files:
        pdf_document = fitz.open(pdf_file)
        total_pages += pdf_document.page_count
        pdf_document.close()
    print(f"Total de páginas a serem processadas: {total_pages}")

    total_pages_processed = [0]

    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(analyze_single_pdf, pdf_file, ws, total_pages_processed, total_pages, progress_var, progress_label, progress_bar) for pdf_file in pdf_files]
        for future in futures:
            result = future.result()
            print(f"Thread concluída com resultado: {result}")

    tab = Table(displayName="PDFAnalysisTable", ref=f"A1:E{ws.max_row}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showRowStripes=True)
    tab.tableStyleInfo = style
    ws.add_table(tab)

    for row in ws.iter_rows(min_row=2, min_col=4, max_col=4, max_row=ws.max_row):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.endswith('%'):
                cell.number_format = '0.00%'

    adjust_column_width(ws)

    try:
        wb.save(output_xlsx)
    except PermissionError as e:
        print(f"Permissão negada ao salvar o relatório: {str(e)}")
        messagebox.showwarning("Permissão Negada", "O arquivo de relatório está aberto. Por favor, feche o arquivo e clique em OK para continuar.")
        wb.save(output_xlsx)
    print(f"Análise concluída. Relatório salvo em {output_xlsx}")

    messagebox.showinfo("Sucesso", f"Análise concluída. Relatório salvo em {output_xlsx}")

def adjust_column_width(ws):
    print("Ajustando a largura das colunas...")
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width
    print("Largura das colunas ajustada.")

def create_gui():
    print("Criando GUI...")
    global window, select_label, analyze_button, progress_var, progress_label, progress_bar, open_folder_button, canvas, directory
    directory = None
    window = ThemedTk(theme="breeze")
    window.title("PDF Analyzer")
    window.geometry("800x600")
    window.configure(bg="blue")

    main_frame = ttk.Frame(window, padding="20")
    main_frame.pack(fill='both', expand=True)

    label = ttk.Label(main_frame, text="Selecione um diretório para analisar PDFs:")
    label.pack(pady=10)

    select_button = ttk.Button(main_frame, text="Selecionar Diretório", command=select_directory)
    select_button.pack(pady=10)

    select_label = ttk.Label(main_frame, text="Nenhum diretório selecionado", foreground="gray")
    select_label.pack(pady=5)

    analyze_button = ttk.Button(main_frame, text="Analisar", state="disabled", command=start_analysis)
    analyze_button.pack(pady=15)

    progress_var = StringVar()
    progress_var.set("0")
    progress_label = ttk.Label(main_frame, text="Progresso: 0%")
    progress_label.pack(pady=10)
    progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=400, mode="determinate", variable=progress_var)
    progress_bar.pack(pady=10)

    open_folder_button = ttk.Button(main_frame, text="Abrir Pasta do Relatório", command=open_folder, state="normal")
    open_folder_button.pack(pady=15)

    canvas = Canvas(window, width=600, height=400, bg="white", scrollregion=(0, 0, 800, 600))
    canvas.pack(pady=20, expand=True, fill="both")
    window.mainloop()
    print("GUI criada.")

if __name__ == "__main__":
    print("Inicializando Tkinter...")
    root = Tk()
    root.withdraw()
    test_tesseract_setup()
    root.destroy()
    create_gui()