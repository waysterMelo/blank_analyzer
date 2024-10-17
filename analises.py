import os
from tkinter import Toplevel, Listbox, Button, messagebox, filedialog, ttk, Canvas, Frame
from ttkthemes import ThemedTk
import pandas as pd
import pdfplumber
from PIL import Image, ImageTk


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
        delete_button = Button(self.pdf_view_frame, text="Deletar PDF Selecionado", command=self.delete_selected_pdf,
                               bg="#f44336", fg="white", font=("Segoe UI", 10, "bold"), activebackground="#e53935",
                               padx=10, pady=5)
        delete_button.pack(side='bottom', pady=5)

        # Botão para abrir o diretório dos PDFs analisados
        open_button = Button(self.window, text="Abrir Diretório dos PDFs", command=self.open_pdf_directory)
        open_button.pack(pady=10)

        self.load_pending_files()

    def load_pending_files(self):
        print("Carregando arquivos pendentes...")
        if not os.path.exists(self.analysis_report_path):
            print(f"Erro: Arquivo de relatório não encontrado no caminho: {self.analysis_report_path}")
            messagebox.showerror("Erro", "Arquivo de relatório não encontrado!")
            return

        try:
            df = pd.read_excel(self.analysis_report_path)
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
        selected_index = self.pending_files_listbox.curselection()
        if not selected_index:
            return
        selected_entry = self.pending_files_listbox.get(selected_index)
        pdf_name, page_info = selected_entry.split(' - Página ')
        selected_pdf = os.path.join(self.selected_directory, pdf_name)
        page_number = int(page_info)
        if os.path.exists(selected_pdf):
            try:
                self.render_pdf_page(selected_pdf, page_number)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o PDF: {str(e)}")
        else:
            messagebox.showerror("Erro", "Arquivo PDF não encontrado!")

    def open_pdf_directory(self):
        # Abre o diretório onde estão os PDFs analisados
        if os.path.exists(self.selected_directory):
            os.startfile(self.selected_directory)  # Somente para Windows
        else:
            messagebox.showerror("Erro", "Diretório dos PDFs não encontrado!")

    def delete_selected_pdf(self):
        selected_index = self.pending_files_listbox.curselection()
        if not selected_index:
            messagebox.showwarning("Aviso", "Nenhum PDF selecionado!")
            return

        selected_entry = self.pending_files_listbox.get(selected_index)
        pdf_name, _ = selected_entry.split(' - Página ')
        selected_pdf = os.path.join(self.selected_directory, pdf_name)

        # Confirmar a exclusão
        if messagebox.askyesno("Confirmação", f"Tem certeza de que deseja deletar {pdf_name}?"):
            try:
                os.remove(selected_pdf)
                self.pending_files_listbox.delete(selected_index)
                self.pdf_canvas.delete("all")  # Limpar o Canvas após a exclusão
                messagebox.showinfo("Sucesso", f"{pdf_name} foi deletado com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível deletar o PDF: {str(e)}")

    def center_image(self, event):
        self.pdf_canvas.delete('all')
        if hasattr(self, 'pdf_image'):
            self.pdf_canvas.create_image(self.pdf_canvas.winfo_width() // 2, self.pdf_canvas.winfo_height() // 2,
                                         anchor='center', image=self.pdf_image, tags='pdf_image')

    def render_pdf_page(self, pdf_path, page_number):
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
                    self.center_image(None)
                else:
                    messagebox.showerror("Erro", f"Número de página {page_number} está fora do intervalo.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao renderizar a página do PDF: {str(e)}")


if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    root.withdraw()
    analysis_screen = AnalysisScreen(root, os.path.join(filedialog.askdirectory(), "analysis_report.xlsx"))
    root.mainloop()