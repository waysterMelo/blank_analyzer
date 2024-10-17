import os
import platform
import subprocess
from tkinter import Toplevel, Listbox, Button, messagebox, filedialog
from ttkthemes import ThemedTk
import pandas as pd


class AnalysisScreen:
    def __init__(self, master, analysis_report_path):
        print("Inicializando AnalysisScreen...")
        # Cria uma nova janela filha
        self.window = Toplevel(master)
        self.window.title("Análises Pendentes")
        self.window.geometry("600x400")
        self.window.configure(bg="#2B3E50")

        # Caminho do relatório de análise
        self.analysis_report_path = analysis_report_path
        self.selected_directory = os.path.dirname(analysis_report_path)
        print(f"Caminho do relatório de análise: {self.analysis_report_path}")

        # Lista de PDFs que não foram marcados como OK
        self.pending_files_listbox = Listbox(self.window, width=80, height=20)
        self.pending_files_listbox.pack(pady=20)

        # Botão para abrir arquivo selecionado
        open_button = Button(self.window, text="Abrir PDF Selecionado", command=self.open_selected_pdf)
        open_button.pack(pady=10)

        # Carregar lista de PDFs não marcados como OK
        self.load_pending_files()

    def load_pending_files(self):
        print("Carregando arquivos pendentes...")
        # Verifica se o arquivo de relatório existe
        if not os.path.exists(self.analysis_report_path):
            print(f"Erro: Arquivo de relatório não encontrado no caminho: {self.analysis_report_path}")
            messagebox.showerror("Erro", "Arquivo de relatório não encontrado!")
            return

        # Carrega o relatório em um DataFrame
        try:
            df = pd.read_excel(self.analysis_report_path)
            print("Relatório carregado com sucesso.")
            if 'Status' not in df.columns or 'Arquivo PDF' not in df.columns:
                raise ValueError("Formato do relatório inválido: colunas necessárias não encontradas.")
        except Exception as e:
            print(f"Erro ao carregar o relatório: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao carregar o relatório: {str(e)}")
            return

        # Filtra apenas os PDFs que não foram marcados como OK
        pending_pages = df[df['Status'] != 'OK'][['Arquivo PDF', 'Página']].drop_duplicates()
        # Adiciona as páginas pendentes ao Listbox
        for _, row in pending_pages.iterrows():
            pdf_name = row['Arquivo PDF']
            page_number = row['Página']
            entry = f"{pdf_name} - Página {page_number}"
            print(f"Adicionando página pendente: {entry}")
            self.pending_files_listbox.insert('end', entry)
        print(f"Páginas pendentes encontradas: {len(pending_pages)}")

    def open_selected_pdf(self):
        print("Tentando abrir PDF selecionado...")
        # Obtém o PDF selecionado na lista
        selected_index = self.pending_files_listbox.curselection()
        if not selected_index:
            print("Aviso: Nenhum PDF selecionado!")
            messagebox.showwarning("Aviso", "Nenhum PDF selecionado!")
            return

        selected_entry = self.pending_files_listbox.get(selected_index)
        pdf_name, page_info = selected_entry.split(' - Página ')
        selected_pdf = os.path.join(self.selected_directory, pdf_name)
        print(f"PDF selecionado: {selected_pdf}")

        # Constrói o caminho completo do PDF selecionado
        selected_pdf_path = selected_pdf
        print(f"Caminho completo do PDF selecionado: {selected_pdf_path}")

        # Verifica se o arquivo selecionado existe
        if os.path.exists(selected_pdf_path):
            try:
                print(f"Abrindo PDF: {selected_pdf_path}")
                if platform.system() == "Windows":
                    os.startfile(selected_pdf_path)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", selected_pdf_path])
                elif platform.system() == "Linux":
                    subprocess.Popen(["xdg-open", selected_pdf_path])
            except Exception as e:
                print(f"Erro ao abrir o PDF: {str(e)}")
                messagebox.showerror("Erro", f"Não foi possível abrir o PDF: {str(e)}")
        else:
            print("Erro: Arquivo PDF não encontrado!")
            messagebox.showerror("Erro", "Arquivo PDF não encontrado!")


if __name__ == "__main__":
    print("Inicializando a interface de Análise de PDFs...")
    # Exemplo de como usar a tela de análises pendentes
    root = ThemedTk(theme="arc")
    root.withdraw()  # Oculta a janela principal
    analysis_screen = AnalysisScreen(root, os.path.join(filedialog.askdirectory(), "analysis_report.xlsx"))
    root.mainloop()
