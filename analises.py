import os
import platform
import subprocess
from tkinter import Toplevel, Listbox, Button, messagebox
from ttkthemes import ThemedTk
import pandas as pd

class AnalysisScreen:
    def __init__(self, master, analysis_report_path):
        # Cria uma nova janela filha
        self.window = Toplevel(master)
        self.window.title("Análises Pendentes")
        self.window.geometry("600x400")
        self.window.configure(bg="#2B3E50")

        # Caminho do relatório de análise
        self.analysis_report_path = analysis_report_path

        # Lista de PDFs que não foram marcados como OK
        self.pending_files_listbox = Listbox(self.window, width=80, height=20)
        self.pending_files_listbox.pack(pady=20)

        # Botão para abrir arquivo selecionado
        open_button = Button(self.window, text="Abrir PDF Selecionado", command=self.open_selected_pdf)
        open_button.pack(pady=10)

        # Carregar lista de PDFs não marcados como OK
        self.load_pending_files()

    def load_pending_files(self):
        # Verifica se o arquivo de relatório existe
        if not os.path.exists(self.analysis_report_path):
            messagebox.showerror("Erro", "Arquivo de relatório não encontrado!")
            return

        # Carrega o relatório em um DataFrame
        try:
            df = pd.read_excel(self.analysis_report_path)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o relatório: {str(e)}")
            return

        # Filtra apenas os PDFs que não foram marcados como OK
        pending_files = df[df['Status'] != 'OK']['PDF_Name'].unique()

        # Adiciona os PDFs pendentes ao Listbox
        for pdf in pending_files:
            self.pending_files_listbox.insert('end', pdf)

    def open_selected_pdf(self):
        # Obtém o PDF selecionado na lista
        selected_index = self.pending_files_listbox.curselection()
        if not selected_index:
            messagebox.showwarning("Aviso", "Nenhum PDF selecionado!")
            return

        selected_pdf = self.pending_files_listbox.get(selected_index)

        # Verifica se o arquivo selecionado existe
        if os.path.exists(selected_pdf):
            try:
                if platform.system() == "Windows":
                    os.startfile(selected_pdf)
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", selected_pdf])
                elif platform.system() == "Linux":
                    subprocess.Popen(["xdg-open", selected_pdf])
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o PDF: {str(e)}")
        else:
            messagebox.showerror("Erro", "Arquivo PDF não encontrado!")

if __name__ == "__main__":
    # Exemplo de como usar a tela de análises pendentes
    root = ThemedTk(theme="arc")
    root.withdraw()  # Oculta a janela principal
    analysis_screen = AnalysisScreen(root, "caminho_para_o_relatorio.xlsx")
    root.mainloop()
