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
from analises import AnalysisScreen


class PDFAnalyzerGUI:
    def __init__(self):
        # Inicializa os componentes da interface gráfica e outros atributos
        self.canvas = None
        self.open_folder_button = None
        self.open_analise_button = None
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
        self.window.state("zoomed")  # Maximiza a janela

        # Diretório selecionado e instâncias de classes auxiliares
        self.directory = None
        self.analyzer = PDFAnalyzer()
        self.report_generator = ReportGenerator()

        # Fila para gerenciar progresso de processamento
        self.progress_queue = queue.Queue()

        # Configurações de estilo e criação dos componentes da interface
        self.setup_style()
        self.create_widgets()

        # Inicia o processamento da fila de progresso
        self.process_queue()

        # Inicia a interface gráfica
        self.window.mainloop()

    def setup_style(self):
        # Configura o estilo dos componentes da interface gráfica
        style = ttk.Style(self.window)
        style.configure("TFrame", background="#2B3E50")
        style.configure("TLabel", background="#2B3E50", foreground="white", font=("Segoe UI", 10))
        style.configure("Header.TLabel", foreground="#ECECEC", font=("Segoe UI", 12, "bold"))
        style.configure("TButton", background="#374956", foreground="#000000", font=("Segoe UI", 10, "bold"),
                        padding=10)
        style.map("TButton", background=[("active", "#516B7F")], relief=[("pressed", "ridge"), ("!pressed", "flat")])

    def create_widgets(self):
        # Cria os componentes da interface gráfica
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill='both', expand=True)

        # Cabeçalho
        header_label = ttk.Label(main_frame, text="Selecione um diretório para análise de PDFs:", style="Header.TLabel")
        header_label.pack(pady=(0, 20))

        # Botão para selecionar diretório
        select_button = ttk.Button(main_frame, text="Selecionar Diretório", command=self.select_directory, width=25)
        select_button.pack(pady=10)

        # Label para mostrar diretório selecionado
        self.select_label = ttk.Label(main_frame, text="Nenhum diretório selecionado", foreground="#a6a6a6")
        self.select_label.pack(pady=(5, 15))

        # Botão para iniciar análise
        self.analyze_button = ttk.Button(main_frame, text="Iniciar Análise", state="disabled", command=self.start_analysis, width=25)
        self.analyze_button.pack(pady=10)

        # Barra de progresso
        self.progress_var = StringVar()
        self.progress_var.set("0")
        self.progress_label = ttk.Label(main_frame, text="Progresso: 0%", style="TLabel")
        self.progress_label.pack(pady=10)
        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=500, mode="determinate", variable=self.progress_var)
        self.progress_bar.pack(pady=(5, 20))

        # Labels para exibir contagens de páginas
        self.pages_blank_after_ocr_label = ttk.Label(main_frame, text="Página em Branco: 0")
        self.pages_blank_after_ocr_label.pack(pady=5)
        self.pages_total_checked_label = ttk.Label(main_frame, text="Total de Páginas Verificadas: 0")  # Nova label para total de páginas
        self.pages_total_checked_label.pack(pady=5)

        # Botões para abrir relatórios e análises, lado a lado
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=15)
        self.open_folder_button = ttk.Button(button_frame, text="Abrir Pasta do Relatório", command=self.open_folder, width=25)
        self.open_folder_button.grid(row=0, column=0, padx=10)
        self.open_analise_button = ttk.Button(button_frame, text="Abrir Análises", command=self.open_analysis_screen, width=25)
        self.open_analise_button.grid(row=0, column=1, padx=10)

        # Frame do canvas para exibir imagens
        canvas_frame = ttk.Frame(main_frame, padding="10", borderwidth=2, relief="ridge", style="TFrame")
        canvas_frame.pack(pady=20, expand=True, fill="both")

        # Canvas para exibir páginas dos PDFs
        self.canvas = Canvas(canvas_frame, bg="white", width=800, height=600)
        self.canvas.pack(expand=True, fill="both")

    def select_directory(self):
        # Abre um diálogo para selecionar o diretório contendo os arquivos PDF
        self.directory = filedialog.askdirectory()
        if self.directory:
            self.select_label.config(text=f"Diretório Selecionado: {self.directory}")
            self.analyze_button.config(state="normal")  # Habilita o botão de análise

    def start_analysis(self):
        # Inicia a análise dos PDFs no diretório selecionado em uma nova thread
        if self.directory:
            self.analyze_button.config(state="disabled")  # Desabilita o botão durante a análise
            threading.Thread(target=self.run_analysis_thread, daemon=True).start()

    def run_analysis_thread(self):
        # Executa a análise dos PDFs e gera um relatório
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_xlsx = os.path.join(self.directory, f"analysis_report_{timestamp}.xlsx")
        self.analyze_pdfs_in_directory(output_xlsx)

    def analyze_pdfs_in_directory(self, output_xlsx):
        # Analisa todos os PDFs no diretório selecionado e gera um relatório
        pdf_files = [os.path.join(self.directory, f) for f in os.listdir(self.directory) if f.lower().endswith('.pdf')]
        total_pages = sum(fitz.open(pdf_file).page_count for pdf_file in pdf_files)  # Calcula o total de páginas
        total_pages_processed = 0

        start_time = time.time()
        estimated_time_per_page = 2  # Tempo estimado por página, em segundos (ajuste conforme necessário)

        # Itera sobre cada arquivo PDF no diretório
        for pdf_file in pdf_files:
            pdf_name = os.path.basename(pdf_file)
            with fitz.open(pdf_file) as pdf_document:
                for page_num in range(pdf_document.page_count):
                    # Carrega a página e converte para imagem
                    page = pdf_document.load_page(page_num)
                    pix = page.get_pixmap()
                    img = Image.open(io.BytesIO(pix.tobytes("png")))

                    # Realiza análise na página
                    status, white_pixel_percentage, ocr_performed, extracted_text = self.analyzer.analyze_page(img)
                    # Adiciona os resultados ao gerador de relatórios
                    self.report_generator.add_record(pdf_name, page_num + 1, status, white_pixel_percentage,
                                                     ocr_performed, extracted_text)

                    # Atualizar labels e progresso
                    total_pages_processed += 1
                    self.update_labels(total_pages_processed)
                    progress_percentage = (total_pages_processed / total_pages) * 100
                    self.progress_queue.put(progress_percentage)

                    # Enviar a imagem para ser exibida no canvas
                    self.progress_queue.put(("image", img))

        # Finaliza o relatório após processar todas as páginas
        self.report_generator.finalize(output_xlsx)
        if not os.path.exists(output_xlsx):
            print("Erro: O relatório não foi criado.")
            return
        self.progress_queue.put("DONE")  # Indica que a análise foi concluída

    def process_queue(self):
        # Processa os itens da fila para atualizar a interface em tempo real
        try:
            while True:
                message = self.progress_queue.get_nowait()
                if isinstance(message, tuple) and message[0] == "image":
                    # Exibe a imagem no canvas
                    image = message[1]
                    self.display_image_on_canvas(image)
                elif isinstance(message, float) or isinstance(message, int):
                    # Atualiza a barra de progresso
                    self.update_progress(message)
                elif message == "DONE":
                    # Mostra mensagem de conclusão
                    self.analyze_button.config(state="normal")
                    messagebox.showinfo("Análise Concluída",
                                        "A análise foi concluída e o relatório foi gerado com sucesso!")
        except queue.Empty:
            pass
        # Verifica a fila novamente após 100 ms
        self.window.after(100, self.process_queue)

    def update_progress(self, percentage):
        # Atualiza o valor da barra de progresso e o texto exibido
        self.progress_var.set(f"{int(percentage)}")
        self.progress_label.config(text=f"Analisando... {int(percentage)}%")
        self.progress_bar['value'] = percentage
        self.progress_bar.update_idletasks()

    def update_labels(self, total_pages_checked):
        # Atualiza as labels que exibem informações sobre a análise
        self.pages_blank_after_ocr_label.config(
            text=f"Página em Branco após análises: {self.analyzer.pages_blank_after_ocr_count}")
        self.pages_total_checked_label.config(
            text=f"Total de Páginas Verificadas: {total_pages_checked}")

    def display_image_on_canvas(self, image):
        print("Exibindo imagem no canvas...")

        # Obtenha dimensões do canvas e da imagem
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        img_width, img_height = image.size

        # Define dimensões padrão se o canvas não estiver inicializado corretamente
        if canvas_width == 1 or canvas_height == 1:
            canvas_width, canvas_height = 400, 600
            self.canvas.config(width=canvas_width, height=canvas_height)

        # Define o filtro de reamostragem adequado
        try:
            resample_filter = Image.Resampling.LANCZOS
        except AttributeError:
            resample_filter = Image.LANCZOS

        # Calcula a nova escala da imagem para caber no canvas
        ratio = min(canvas_width / img_width, canvas_height / img_height)
        new_size = (int(img_width * ratio), int(img_height * ratio))
        image = image.resize(new_size, resample=resample_filter)

        # Calcula a posição para centralizar a imagem no canvas
        x = (canvas_width - new_size[0]) // 2
        y = (canvas_height - new_size[1]) // 2
        print(f"Posicionando imagem no canvas: ({x}, {y})")

        # Exibe a imagem no canvas
        tk_image = ImageTk.PhotoImage(image)
        self.canvas.delete("all")
        self.canvas.create_image(x, y, anchor="nw", image=tk_image)
        self.canvas.image = tk_image  # Evita que o Python faça coleta de lixo da imagem

    def open_folder(self):
        # Abre o diretório contendo os PDFs ou relatórios
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

    def open_analysis_screen(self):
        # Verifica se um diretório foi selecionado
        if not self.directory:
            messagebox.showerror("Erro", "Nenhum diretório selecionado.")
            return

        # Encontra o arquivo de relatório mais recente com o padrão analysis_report_YYYYMMDD_HHMMSS.xlsx
        report_files = [f for f in os.listdir(self.directory) if
                        f.startswith("analysis_report_") and f.endswith(".xlsx")]

        if not report_files:
            messagebox.showerror("Erro", "Nenhum relatório de análise encontrado.")
            return

        # Modificado para lidar com diferentes tamanhos de strings antes do timestamp
        report_files.sort(
            key=lambda x: datetime.strptime(x.replace("analysis_report_", "").replace(".xlsx", ""), "%Y%m%d_%H%M%S"), reverse=True)
        analysis_report_path = os.path.join(self.directory, report_files[0])

        # Cria a tela de análises pendentes
        AnalysisScreen(self.window, analysis_report_path)
