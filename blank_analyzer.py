import fitz  # PyMuPDF
import cv2
import numpy as np
import os
from tkinter import Tk, Label, Button, filedialog, messagebox, StringVar, ttk, Canvas, PhotoImage, NW, CENTER
import threading
import subprocess
import platform
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from ttkthemes import ThemedTk
import time
from PIL import Image, ImageTk
import io
import sys

# Definir o TESSDATA_PREFIX corretamente antes de importar o pytesseract
tessdata_prefix = r'C:/Program Files/Tesseract-OCR/tessdata/'
os.environ['TESSDATA_PREFIX'] = tessdata_prefix

import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

def test_tesseract_setup():
    print("Testing Tesseract setup...")
    try:
        tessdata_prefix_env = os.environ.get('TESSDATA_PREFIX')
        if not tessdata_prefix_env:
            raise Exception("A variável de ambiente TESSDATA_PREFIX não está definida.")
        else:
            print(f"TESSDATA_PREFIX encontrado: {tessdata_prefix_env}")
            if not os.path.isdir(tessdata_prefix_env):
                raise Exception(f"O diretório especificado em TESSDATA_PREFIX não existe: {tessdata_prefix_env}")
            else:
                print("Diretório TESSDATA_PREFIX encontrado.")
                lang_files = ['eng.traineddata', 'por.traineddata']
                missing_files = []
                for lang_file in lang_files:
                    lang_file_path = os.path.join(tessdata_prefix_env, lang_file)
                    if not os.path.isfile(lang_file_path):
                        missing_files.append(lang_file)
                if missing_files:
                    raise Exception(f"Os seguintes arquivos de idioma estão faltando em '{tessdata_prefix_env}': {', '.join(missing_files)}")
        tesseract_version = pytesseract.get_tesseract_version()
        print(f"Tesseract OCR instalado corretamente. Versão: {tesseract_version}")
        messagebox.showinfo("Tesseract OCR", f"Tesseract OCR está instalado corretamente.\nVersão: {tesseract_version}\nTESSDATA_PREFIX: {tessdata_prefix_env}")
    except Exception as e:
        print(f"Erro ao inicializar o Tesseract OCR: {str(e)}")
        messagebox.showerror("Erro Tesseract OCR", f"Erro ao inicializar o Tesseract OCR.\n{str(e)}\nPor favor, verifique se o Tesseract está instalado e o TESSDATA_PREFIX está configurado corretamente.")
        sys.exit(1)

def sanitize_filename(filename):
    print(f"Sanitizing filename: {filename}")
    normalized_filename = unicodedata.normalize('NFKD', filename).encode('ascii', 'ignore').decode('ascii')
    print(f"Sanitized filename: {normalized_filename}")
    return normalized_filename

# Função para determinar se a página está em branco ou ruidosa
def is_blank_or_noisy(image, white_threshold=210, pixel_threshold=0.98):
    print("Determining if the page is blank or noisy...")
    if image is None:
        print("Image is None.")
        return False, 0

    resized_image = cv2.resize(image, (image.shape[1] // 2, image.shape[0] // 2))
    white_pixel_percentage = np.mean(resized_image > white_threshold)
    print(f"White pixel percentage: {white_pixel_percentage}")
    return white_pixel_percentage > pixel_threshold, white_pixel_percentage

# Função para executar OCR em uma imagem e verificar se há texto relevante
def perform_ocr_and_reclassify(image, min_text_length=2):
    print("Performing OCR to check for relevant text...")
    try:
        # Aumentar o DPI da imagem para 600 antes de realizar o OCR
        image = image.resize((image.width * 2, image.height * 2))
        image = image.convert('RGB')
        print("Image DPI increased to 600 for better OCR analysis.")

        # Usar ambos os idiomas, inglês e português
        text = pytesseract.image_to_string(image, lang='eng+por')

        # Remover espaços em branco e quebras de linha
        cleaned_text = ''.join(text.split())
        print(f"Extracted text length: {len(cleaned_text)}")
        
        # Retornar True e o texto se mais de 2 caracteres foram extraídos para revisão
        if len(cleaned_text) > min_text_length:
            print("Text needs revisation.")
            return True, "need revisation"
        
        elif len(cleaned_text) >= 10:  # Verificar o comprimento mínimo para conteúdo relevante
            print("Relevant text found.")
            return True, cleaned_text
        else:
            print("No relevant text found.")
            return False, "No Text"
    except pytesseract.TesseractError as e:
        print(f"OCR Error: {str(e)}")
        return False, "OCR Error"

# Função para analisar um PDF único
def analyze_single_pdf(pdf_path, ws, total_pages_processed, total_pages, progress_var, progress_label, progress_bar):
    print(f"Analyzing PDF: {pdf_path}")
    pdf_document = fitz.open(pdf_path)
    total_pdf_pages = pdf_document.page_count
    pdf_name = sanitize_filename(os.path.basename(pdf_path))
    print(f"Total pages in PDF: {total_pdf_pages}")

    for page_num in range(total_pdf_pages):
        print(f"Processing page {page_num + 1}/{total_pdf_pages}")
        try:
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap()  # DPI padrão
            print(f"Pixmap generated for page {page_num + 1}")

            # Obter os dados da imagem em bytes
            image_bytes = pix.tobytes("png")
            print(f"Image bytes obtained for page {page_num + 1}")

            # Usar PIL para carregar a imagem a partir dos bytes
            with Image.open(io.BytesIO(image_bytes)) as img:
                # Retirar 10% das bordas da imagem apenas nos lados
                width, height = img.size
                left = width * 0.10
                right = width * 0.90
                img = img.crop((left, 0, right, height))
                print(f"Cropped image to remove 10% border from each side for page {page_num + 1}")

                image_gray = img.convert('L')  # Converter para escala de cinza
                image_np = np.array(image_gray)
                print(f"Image converted to grayscale and loaded into numpy array for page {page_num + 1}")

            is_blank, white_pixel_percentage = is_blank_or_noisy(image_np)
            status = "Blank or Noisy" if is_blank else "OK"
            ocr_text = "Não"

            if is_blank:
                # Se a página for considerada "Blank or Noisy", executar OCR com aumento de DPI
                with Image.open(io.BytesIO(image_bytes)) as img:
                    # Retirar 10% das bordas da imagem antes do OCR
                    img = img.crop((left, 0, right, height))
                    relevant_text_found, text_or_revision = perform_ocr_and_reclassify(img)
                    if relevant_text_found:
                        status = text_or_revision
                        ocr_text = "Sim"
                    else:
                        status = "Blank after OCR"

            print(f"Page {page_num + 1} status: {status}, White Pixel Percentage: {white_pixel_percentage:.2%}")
            ws.append([pdf_name, page_num + 1, status, f"{white_pixel_percentage:.2%}", ocr_text])

        except (FileNotFoundError, ValueError) as e:
            print(f"Error processing page {page_num + 1} of file {pdf_name}: {str(e)}")
            ws.append([pdf_name, page_num + 1, f"Error: {str(e)}", None, None])

        total_pages_processed[0] += 1
        progress_percentage = (total_pages_processed[0] / total_pages) * 100
        print(f"Progress: {progress_percentage:.2f}%")
        progress_var.set(progress_percentage)
        progress_label.after(0, lambda: progress_label.config(text=f"Analisando... {int(float(progress_var.get()))}%"))
        progress_bar.after(0, progress_bar.update_idletasks)

    pdf_document.close()
    print(f"Finished analyzing PDF: {pdf_path}")

# Função para selecionar o diretório
def select_directory():
    global directory
    directory = filedialog.askdirectory()
    if directory:
        print(f"Selected directory: {directory}")
        select_label.config(text=f"Selected Directory: {directory}")
        analyze_button.config(state="normal")

# Função para iniciar a análise
def start_analysis():
    if directory:
        print("Starting analysis...")
        analyze_button.config(state="disabled")
        analysis_thread = threading.Thread(target=run_analysis_thread, args=(directory, progress_var, progress_label, progress_bar))
        analysis_thread.start()

# Função para executar a análise em uma thread separada
def run_analysis_thread(directory, progress_var, progress_label, progress_bar):
    print(f"Starting analysis thread for directory: {directory}")
    output_xlsx = os.path.join(directory, "analysis_report.xlsx")
    analyze_pdfs_in_directory(directory, progress_var, progress_label, output_xlsx, progress_bar)

# Função para abrir a pasta com o relatório
def open_folder():
    if directory:
        print(f"Opening folder: {directory}")
        if platform.system() == "Windows":
            os.startfile(directory)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", directory])
        elif platform.system() == "Linux":
            subprocess.Popen(["xdg-open", directory])

# Função para analisar PDFs em um diretório
def analyze_pdfs_in_directory(pdf_dir, progress_var, progress_label, output_xlsx, progress_bar):
    print(f"Analyzing PDFs in directory: {pdf_dir}")

    wb = Workbook()
    ws = wb.active
    ws.title = "PDF Analysis Report"
    ws.append(["PDF File", "Page", "Status", "White Pixel Percentage", "Extracted Text"])

    pdf_files = [os.path.join(pdf_dir, f) for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
    print(f"Total PDF files found: {len(pdf_files)}")

    total_pages = 0
    for pdf_file in pdf_files:
        pdf_document = fitz.open(pdf_file)
        total_pages += pdf_document.page_count
        pdf_document.close()
    print(f"Total pages to be processed: {total_pages}")

    total_pages_processed = [0]

    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(analyze_single_pdf, pdf_file, ws, total_pages_processed, total_pages, progress_var, progress_label, progress_bar) for pdf_file in pdf_files]
        for future in futures:
            result = future.result()
            print(f"Thread completed with result: {result}")

    tab = Table(displayName="PDFAnalysisTable", ref=f"A1:E{ws.max_row}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style
    ws.add_table(tab)

    for row in ws.iter_rows(min_row=2, min_col=4, max_col=4, max_row=ws.max_row):
        for cell in row:
            if cell.value is not None and isinstance(cell.value, str) and cell.value.endswith('%'):
                cell.number_format = '0.00%'

    adjust_column_width(ws)

    try:
        wb.save(output_xlsx)
    except PermissionError as e:
        print(f"Permission denied while saving the report: {str(e)}")
        messagebox.showwarning("Permission Denied", "O arquivo de relatório está aberto. Por favor, feche o arquivo e clique em OK para continuar.")
        wb.save(output_xlsx)
    print(f"Analysis completed. Report saved at {output_xlsx}")

    messagebox.showinfo("Success", f"Analysis completed. Report saved at {output_xlsx}")

    if platform.system() == "Windows":
        os.startfile(pdf_dir)
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", pdf_dir])
    elif platform.system() == "Linux":
        subprocess.Popen(["xdg-open", pdf_dir])

# Ajustar a largura das colunas do relatório
def adjust_column_width(ws):
    print("Adjusting column widths...")
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width
    print("Column widths adjusted.")

# Função para criar a GUI
def create_gui():
    print("Creating GUI...")
    global window, select_label, analyze_button, progress_var, progress_label, progress_bar, open_folder_button, canvas, pdf_files, current_pdf_index, current_page_num
    global directory
    directory = None
    pdf_files = []
    current_pdf_index = 0
    current_page_num = 0

    window = ThemedTk(theme="breeze")
    window.title("PDF Analyzer")
    window.geometry("800x600")
    window.resizable(True, True)
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

    # Canvas para exibir a página do PDF
    canvas = Canvas(window, width=600, height=400, bg="white", scrollregion=(0, 0, 800, 600))
    hbar = ttk.Scrollbar(window, orient="horizontal", command=canvas.xview)
    hbar.pack(side="bottom", fill="x")
    vbar = ttk.Scrollbar(window, orient="vertical", command=canvas.yview)
    vbar.pack(side="right", fill="y")
    canvas.config(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
    canvas.pack(pady=20, expand=True, fill="both")

    window.mainloop()
    print("GUI created.")

if __name__ == "__main__":
    # Inicializar o Tkinter para poder usar messagebox
    print("Initializing Tkinter...")
    root = Tk()
    root.withdraw()  # Ocultar a janela raiz
    test_tesseract_setup()
    root.destroy()
    create_gui()
