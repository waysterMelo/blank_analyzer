import os
import platform
import subprocess
import threading
from tkinter import filedialog, messagebox, StringVar, Canvas, Label
import fitz
from ttkthemes import ThemedTk
from tkinter import ttk
from PIL import Image, ImageTk
from datetime import datetime
import io
import queue
from pdf_analyzer import PDFAnalyzer
from report_generator import ReportGenerator
import time


class PDFAnalyzerGUI:
    def __init__(self):
        self.canvas = None
        self.open_folder_button = None
        self.pages_blank_after_ocr_label = None
        self.pages_total_checked_label = None  # Nova label para total de páginas verificadas
        self.progress_bar = None
        self.progress_label = None
        self.progress_var = None
        self.analyze_button = None
        self.select_label = None
        self.timer_label = None  # Label do timer
        self.window = ThemedTk(theme="arc")
        self.window.title("Analisador de PDFs - Digitalizados")
        self.window.geometry("950x750")  # Aumentei a altura para acomodar melhor o canvas

        self.directory = None
        self.analyzer = PDFAnalyzer()
        self.report_generator = ReportGenerator()

        self.progress_queue = queue.Queue()

        self.setup_style()
        self.create_widgets()

        self.process_queue()

        self.window.mainloop()

    def setup_style(self):
        style = ttk.Style(self.window)
        style.configure("TFrame", background="#2B3E50")
        style.configure("TLabel", background="#2B3E50", foreground="white", font=("Segoe UI", 10))
        style.configure("Header.TLabel", foreground="#ECECEC", font=("Segoe UI", 12, "bold"))
        style.configure("TButton", background="#374956", foreground="#000000", font=("Segoe UI", 10, "bold"),
                        padding=10)
        style.map("TButton", background=[("active", "#516B7F")], relief=[("pressed", "ridge"), ("!pressed", "flat")])

    def create_widgets(self):
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill='both', expand=True)

        header_label = ttk.Label(main_frame, text="Selecione um diretório para análise de PDFs:", style="Header.TLabel")
        header_label.pack(pady=(0, 20))

        select_button = ttk.Button(main_frame, text="Selecionar Diretório", command=self.select_directory, width=25)
        select_button.pack(pady=10)

        self.select_label = ttk.Label(main_frame, text="Nenhum diretório selecionado", foreground="#a6a6a6")
        self.select_label.pack(pady=(5, 15))

        self.analyze_button = ttk.Button(main_frame, text="Iniciar Análise", state="disabled",
                                         command=self.start_analysis, width=25)
        self.analyze_button.pack(pady=10)

        self.progress_var = StringVar()
        self.progress_var.set("0")
        self.progress_label = ttk.Label(main_frame, text="Progresso: 0%", style="TLabel")
        self.progress_label.pack(pady=10)
        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=500, mode="determinate",
                                            variable=self.progress_var)
        self.progress_bar.pack(pady=(5, 20))

        # Labels para exibir contagens
        self.pages_blank_after_ocr_label = ttk.Label(main_frame, text="Página em Branco: 0")
        self.pages_blank_after_ocr_label.pack(pady=5)
        self.pages_total_checked_label = ttk.Label(main_frame,
                                                   text="Total de Páginas Verificadas: 0")  # Nova label para total de páginas
        self.pages_total_checked_label.pack(pady=5)
        self.open_folder_button = ttk.Button(main_frame, text="Abrir Pasta do Relatório", command=self.open_folder,
                                             width=25)
        self.open_folder_button.pack(pady=15)

        # Timer Label

        canvas_frame = ttk.Frame(main_frame, padding="10", borderwidth=2, relief="ridge", style="TFrame")
        canvas_frame.pack(pady=20, expand=True, fill="both")

        self.canvas = Canvas(canvas_frame, bg="white", width=800, height=600)
        self.canvas.pack(expand=True, fill="both")

    def select_directory(self):
        self.directory = filedialog.askdirectory()
        if self.directory:
            self.select_label.config(text=f"Diretório Selecionado: {self.directory}")
            self.analyze_button.config(state="normal")

    def start_analysis(self):
        if self.directory:
            self.analyze_button.config(state="disabled")
            threading.Thread(target=self.run_analysis_thread, daemon=True).start()

    def run_analysis_thread(self):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_xlsx = os.path.join(self.directory, f"analysis_report_{timestamp}.xlsx")
        self.analyze_pdfs_in_directory(output_xlsx)

    def analyze_pdfs_in_directory(self, output_xlsx):
        pdf_files = [os.path.join(self.directory, f) for f in os.listdir(self.directory) if f.lower().endswith('.pdf')]
        total_pages = sum(fitz.open(pdf_file).page_count for pdf_file in pdf_files)
        total_pages_processed = 0

        start_time = time.time()
        estimated_time_per_page = 2  # Tempo estimado por página, em segundos (ajuste conforme necessário)

        for pdf_file in pdf_files:
            pdf_name = os.path.basename(pdf_file)
            with fitz.open(pdf_file) as pdf_document:
                for page_num in range(pdf_document.page_count):
                    page = pdf_document.load_page(page_num)
                    pix = page.get_pixmap()
                    img = Image.open(io.BytesIO(pix.tobytes("png")))

                    status, white_pixel_percentage, ocr_performed, extracted_text = self.analyzer.analyze_page(img)
                    self.report_generator.add_record(pdf_name, page_num + 1, status, white_pixel_percentage,
                                                     ocr_performed, extracted_text)

                    # Atualizar labels e progresso
                    total_pages_processed += 1
                    self.update_labels(total_pages_processed)
                    progress_percentage = (total_pages_processed / total_pages) * 100
                    self.progress_queue.put(progress_percentage)

                    # Enviar a imagem para ser exibida no canvas
                    self.progress_queue.put(("image", img))

        self.report_generator.finalize(output_xlsx)
        if not os.path.exists(output_xlsx):
            print("Erro: O relatório não foi criado.")
            return
        self.progress_queue.put("DONE")

    def process_queue(self):
        try:
            while True:
                message = self.progress_queue.get_nowait()
                if isinstance(message, tuple) and message[0] == "image":
                    image = message[1]
                    self.display_image_on_canvas(image)
                elif isinstance(message, float) or isinstance(message, int):
                    self.update_progress(message)
                elif message == "DONE":
                    self.analyze_button.config(state="normal")
                    messagebox.showinfo("Análise Concluída",
                                        "A análise foi concluída e o relatório foi gerado com sucesso!")
        except queue.Empty:
            pass
        self.window.after(100, self.process_queue)

    def update_progress(self, percentage):
        self.progress_var.set(f"{int(percentage)}")
        self.progress_label.config(text=f"Analisando... {int(percentage)}%")
        self.progress_bar['value'] = percentage
        self.progress_bar.update_idletasks()

    def update_labels(self, total_pages_checked):
        self.pages_blank_after_ocr_label.config(
            text=f"Página em Branco após análises: {self.analyzer.pages_blank_after_ocr_count}")
        self.pages_total_checked_label.config(
            text=f"Total de Páginas Verificadas: {total_pages_checked}")  # Atualiza a nova label

    def display_image_on_canvas(self, image):
        print("Exibindo imagem no canvas...")

        # Obtenha dimensões do canvas e da imagem
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        img_width, img_height = image.size

        if canvas_width == 1 or canvas_height == 1:
            canvas_width, canvas_height = 400, 600
            self.canvas.config(width=canvas_width, height=canvas_height)

        try:
            resample_filter = Image.Resampling.LANCZOS
        except AttributeError:
            resample_filter = Image.LANCZOS

        ratio = min(canvas_width / img_width, canvas_height / img_height)
        new_size = (int(img_width * ratio), int(img_height * ratio))
        image = image.resize(new_size, resample=resample_filter)

        x = (canvas_width - new_size[0]) // 2
        y = (canvas_height - new_size[1]) // 2
        print(f"Posicionando imagem no canvas: ({x}, {y})")

        tk_image = ImageTk.PhotoImage(image)
        self.canvas.delete("all")
        self.canvas.create_image(x, y, anchor="nw", image=tk_image)
        self.canvas.image = tk_image  # Evita que o Python faça coleta de lixo da imagem

    def open_folder(self):
        if self.directory:
            try:
                if platform.system() == "Windows":
                    os.startfile(self.directory)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", self.directory])
                elif platform.system() == "Linux":
                    subprocess.Popen(["xdg-open", self.directory])
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir a pasta: {str(e)}")

    def open_report(self, file_path):
        print(f"Abrindo relatório: {file_path}")
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", file_path])
            elif platform.system() == "Linux":
                subprocess.Popen(["xdg-open", file_path])
        except Exception as e:
            print(f"Erro ao abrir relatório: {e}")
            messagebox.showerror("Erro ao abrir relatório", f"Não foi possível abrir o relatório: {str(e)}")


if __name__ == "__main__":
    PDFAnalyzerGUI()