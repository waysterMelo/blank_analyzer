import tkinter as tk
from PIL import Image, ImageTk
import subprocess
import os

# Caminho para a imagem de fundo
image_path = r'C:\Users\Wayster\Desktop\Projetos\blank_analyzer\logo.webp'

def main():
    # Criando a janela principal
    root = tk.Tk()
    root.title("PDF Analyzer")
    root.geometry("900x600")  # Definindo o tamanho da janela
    root.configure(bg="#1C2833")  # Cor de fundo da janela (um tom mais sóbrio)
    root.resizable(False, False)  # Tornar a janela fixa em tamanho para manter o design consistente

    # Centralizando a janela na tela
    window_width = 900
    window_height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int((screen_height - window_height) / 2)
    position_left = int((screen_width - window_width) / 2)
    root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")

    # Carregando a imagem de fundo
    image = Image.open(image_path)
    image = image.resize((900, 600))  # Redimensionando a imagem para o tamanho da janela
    bg_image = ImageTk.PhotoImage(image)

    # Criando um rótulo para a imagem de fundo
    background_label = tk.Label(root, image=bg_image)
    background_label.place(relwidth=1, relheight=1)  # Preenchendo toda a janela com a imagem de fundo

    # Adicionando o nome da aplicação no centro da tela com um visual mais sério e atraente
    app_name = tk.Label(root, text="PDF Analyzer", font=("Helvetica", 40, "bold"), fg="#FFFFFF", bg="#1E3D59", padx=20, pady=10, relief="raised", bd=10)
    app_name.place(relx=0.5, rely=0.25, anchor='center')

    # Função para iniciar a análise chamando o script blank_analyzer.py
    def iniciar_analise():
        current_directory = os.path.dirname(os.path.abspath(__file__))
        analyzer_path = os.path.join(current_directory, 'blank_analyzer.py')
        subprocess.Popen(['python', analyzer_path])

    # Adicionando um botão para iniciar a análise com um design mais elegante
    start_button = tk.Button(root, text="Iniciar Análise", font=("Helvetica", 18, "bold"), fg="#FFFFFF", bg="#1E3D59", activebackground="#34495E", activeforeground="#FFFFFF", padx=20, pady=10, relief="raised", bd=5, command=iniciar_analise)
    start_button.place(relx=0.5, rely=0.55, anchor='center')

    # Adicionando direitos autorais no final da tela
    copyright_label = tk.Label(root, text="Direitos Autorais © Wayster Cruz de Melo", font=("Helvetica", 12, "italic"), fg="#FFFFFF", bg="#1C2833", padx=5, pady=5)
    copyright_label.place(relx=0.01, rely=0.95, anchor='w')

    # Iniciando o loop da interface gráfica
    root.mainloop()

if __name__ == "__main__":
    main()
