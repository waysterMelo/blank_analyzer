import os
from tkinter import Toplevel, Listbox, Button, messagebox, filedialog, ttk, Canvas, Frame, Label
from ttkthemes import ThemedTk
import pandas as pd
import pdfplumber
from PIL import Image, ImageTk
import fitz  # Importando a biblioteca PyMuPDF para manipulação de PDFs
import shutil
import threading


class AnalysisScreen:
    def __init__(self, master, analysis_report_path):
        print("Inicializando AnalysisScreen...")
        self.window = Toplevel(master)
        self.window.title("Análises Pendentes")
        self.window.state('zoomed')
        self.window.configure(bg="#2B3E50")

        style = ttk.Style(self.window)
        style.configure("TFrame", background="#2B3E50")
        style.configure("TLabel", background="#2B3E50", foreground="white", font=("Segoe UI", 10))
        style.configure("Header.TLabel", foreground="#ECECEC", font=("Segoe UI", 12, "bold"))

        self.analysis_report_path = analysis_report_path
        self.selected_directory = os.path.dirname(analysis_report_path)
        print(f"Caminho do relatório de análise: {self.analysis_report_path}")

        self.header_frame = ttk.Frame(self.window, style="TFrame")
        self.header_frame.pack(pady=10, padx=20, fill='x')

        self.report_label = ttk.Label(self.header_frame,
                                      text=f"Relatório: {os.path.basename(self.analysis_report_path)}",
                                      style="Header.TLabel")
        self.report_label.pack(side='left', padx=5)

        self.date_label = ttk.Label(self.header_frame, text=f"Data: {pd.to_datetime('today').strftime('%d/%m/%Y')}",
                                    style="Header.TLabel")
        self.date_label.pack(side='right', padx=5)

        self.pending_files_listbox = Listbox(self.window, width=90, height=25)
        self.pending_files_listbox.bind('<<ListboxSelect>>', self.on_pdf_select)
        self.pending_files_listbox.pack(pady=20, padx=10, side='left', fill='y')

        self.pdf_view_frame = ttk.Frame(self.window, padding="10", borderwidth=2, relief="ridge", style="TFrame")
        self.pdf_view_frame.pack(pady=10, padx=10, side='right', fill='both', expand=True)

        self.pdf_canvas = Canvas(self.pdf_view_frame, bg="#2B3E50")
        self.pdf_canvas.pack(expand=True, fill='both')
        self.pdf_canvas.bind('<Configure>', self.center_image)

        # Botão de deletar dentro do pdf_view_frame
        delete_button = Button(self.pdf_view_frame, text="Deletar Página Selecionada", command=self.delete_selected_pdf,
                               bg="#f44336", fg="white", font=("Segoe UI", 10, "bold"), activebackground="#e53935",
                               padx=10, pady=5)
        delete_button.pack(side='bottom', pady=5)

        # Botão para abrir o diretório dos PDFs analisados
        open_button = Button(self.window, text="Abrir Diretório dos PDFs", command=self.open_pdf_directory)
        open_button.pack(pady=10)

        # Label de loading
        self.loading_label = Label(self.window, text="", bg="#2B3E50", fg="white", font=("Segoe UI", 10, "bold"))
        self.loading_label.pack(pady=5)

        print("Chamando load_pending_files...")
        self.load_pending_files()

    def load_pending_files(self):
        print("Carregando arquivos pendentes...")
        if not os.path.exists(self.analysis_report_path):
            print(f"Erro: Arquivo de relatório não encontrado no caminho: {self.analysis_report_path}")
            messagebox.showerror("Erro", "Arquivo de relatório não encontrado!")
            return

        try:
            df = pd.read_excel(self.analysis_report_path)
            print(f"Relatório carregado com sucesso: {len(df)} registros encontrados.")
            if 'Status' not in df.columns or 'Arquivo PDF' not in df.columns:
                raise ValueError("Formato do relatório inválido: colunas necessárias não encontradas.")
            pending_pages = df[df['Status'] != 'OK'][['Arquivo PDF', 'Página']].drop_duplicates()
            for _, row in pending_pages.iterrows():
                pdf_name = row['Arquivo PDF']
                page_number = row['Página']
                entry = f"{pdf_name} - Página {page_number}"
                self.pending_files_listbox.insert('end', entry)
            print(f"Páginas pendentes encontradas: {len(pending_pages)}")
        except Exception as e:
            print(f"Erro ao carregar o relatório: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao carregar o relatório: {str(e)}")

    def on_pdf_select(self, event):
        print("Item selecionado na Listbox...")
        selected_index = self.pending_files_listbox.curselection()
        if not selected_index:
            print("Nenhum item selecionado.")
            return
        selected_entry = self.pending_files_listbox.get(selected_index)
        print(f"Entrada selecionada: {selected_entry}")
        pdf_name, page_info = selected_entry.split(' - Página ')
        self.selected_pdf = os.path.join(self.selected_directory, pdf_name)
        self.selected_page_index = int(page_info) - 1  # Ajuste para zero-based index
        print(f"PDF selecionado: {self.selected_pdf}, Página: {self.selected_page_index + 1}")
        if os.path.exists(self.selected_pdf):
            try:
                print(f"Renderizando página {self.selected_page_index + 1} do PDF {self.selected_pdf}...")
                self.render_pdf_page(self.selected_pdf, self.selected_page_index + 1)
            except Exception as e:
                print(f"Erro ao abrir o PDF: {str(e)}")
                messagebox.showerror("Erro", f"Não foi possível abrir o PDF: {str(e)}")
        else:
            print("Erro: Arquivo PDF não encontrado!")
            messagebox.showerror("Erro", "Arquivo PDF não encontrado!")

    def open_pdf_directory(self):
        print(f"Abrindo diretório dos PDFs: {self.selected_directory}")
        # Abre o diretório onde estão os PDFs analisados
        if os.path.exists(self.selected_directory):
            os.startfile(self.selected_directory)  # Somente para Windows
        else:
            print("Erro: Diretório dos PDFs não encontrado!")
            messagebox.showerror("Erro", "Diretório dos PDFs não encontrado!")

    def delete_selected_pdf(self):
        print("Deletando página selecionada do PDF...")
        selected_index = self.pending_files_listbox.curselection()
        if not selected_index:
            print("Nenhum PDF selecionado para exclusão.")
            messagebox.showwarning("Aviso", "Nenhum PDF selecionado!")
            return

        # Confirma a exclusão da página
        if messagebox.askyesno("Confirmação", f"Tem certeza de que deseja deletar a página {self.selected_page_index + 1} de {os.path.basename(self.selected_pdf)}?"):
            self.loading_label.config(text="Deletando página, por favor aguarde...")
            delete_thread = threading.Thread(target=self.perform_delete, args=(selected_index,))
            delete_thread.start()

    def perform_delete(self, selected_index):
        try:
            print(f"Abrindo PDF {self.selected_pdf} para deletar a página {self.selected_page_index + 1}...")
            # Abre o PDF e exclui a página
            with fitz.open(self.selected_pdf) as pdf_document:
                if pdf_document.page_count > 1:
                    pdf_document.delete_page(self.selected_page_index)

                    # Salva o novo arquivo sem a página deletada em um arquivo temporário
                    temp_pdf_path = self.selected_pdf.replace('.pdf', '_temp.pdf')
                    pdf_document.save(temp_pdf_path, garbage=4, deflate=True)
                    print(f"Página deletada e arquivo temporário salvo em {temp_pdf_path}")

                    # Substitui o arquivo original pelo temporário
                    shutil.move(temp_pdf_path, self.selected_pdf)
                    print(f"Arquivo original substituído pelo arquivo temporário {temp_pdf_path}")

                    # Atualiza a interface
                    print(f"Página {self.selected_page_index + 1} deletada com sucesso. Atualizando interface...")
                    self.pending_files_listbox.delete(selected_index)
                    self.pdf_canvas.delete("all")  # Limpa o Canvas após a exclusão
                    messagebox.showinfo("Sucesso",
                                        f"A página {self.selected_page_index + 1} de {os.path.basename(self.selected_pdf)} foi deletada com sucesso.")
                else:
                    print("Não é possível deletar a única página de um documento PDF.")
                    messagebox.showwarning("Operação Inválida", "Não é possível deletar a única página de um documento PDF.")
        except Exception as e:
            print(f"Erro ao deletar a página do PDF: {str(e)}")
            messagebox.showerror("Erro", f"Não foi possível deletar a página do PDF: {str(e)}")
        finally:
            self.loading_label.config(text="")

    def center_image(self, event):
        print("Centralizando imagem no Canvas...")
        self.pdf_canvas.delete('all')
        if hasattr(self, 'pdf_image'):
            self.pdf_canvas.create_image(self.pdf_canvas.winfo_width() // 2, self.pdf_canvas.winfo_height() // 2,
                                         anchor='center', image=self.pdf_image, tags='pdf_image')

    def render_pdf_page(self, pdf_path, page_number):
        print(f"Renderizando página {page_number} do PDF {pdf_path}...")
        try:
            with pdfplumber.open(pdf_path) as pdf:
                if page_number - 1 < len(pdf.pages):
                    page = pdf.pages[page_number - 1]
                    pil_image = page.to_image(resolution=100).original
                    canvas_width = self.pdf_canvas.winfo_width()
                    canvas_height = self.pdf_canvas.winfo_height()
                    image_ratio = pil_image.width / pil_image.height
                    canvas_ratio = canvas_width / canvas_height

                    if image_ratio > canvas_ratio:
                        new_width = canvas_width
                        new_height = int(canvas_width / image_ratio)
                    else:
                        new_height = canvas_height
                        new_width = int(canvas_height * image_ratio)

                    pil_image = pil_image.resize((new_width, new_height), Image.LANCZOS)

                    self.pdf_image = ImageTk.PhotoImage(pil_image)
                    print("Imagem da página renderizada com sucesso.")
                    self.center_image(None)
                else:
                    print(f"Erro: Número de página {page_number} está fora do intervalo.")
                    messagebox.showerror("Erro", f"Número de página {page_number} está fora do intervalo.")
        except Exception as e:
            print(f"Erro ao renderizar a página do PDF: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao renderizar a página do PDF: {str(e)}")
