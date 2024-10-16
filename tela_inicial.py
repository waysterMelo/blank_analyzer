import tkinter as tk
from PIL import Image, ImageTk
import subprocess
import os
from main import main as iniciar_analise_main

# Caminho para a imagem de fundo
image_path = 'logo.webp'

def tela_inicial():
    # Criando a janela principal
    print("Iniciando a criação da janela principal")
    root = tk.Tk()
    root.title("PDF Analyzer")
    root.geometry("900x600")  # Definindo o tamanho da janela
    root.configure(bg="#1C2833")  # Cor de fundo da janela (um tom mais sóbrio)
    root.resizable(False, False)  # Tornar a janela fixa em tamanho para manter o design consistente
    print("Janela principal criada com sucesso")

    # Centralizando a janela na tela
    window_width = 900
    window_height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int((screen_height - window_height) / 2)
    position_left = int((screen_width - window_width) / 2)
    root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")
    print(f"Janela centralizada na posição: ({position_left}, {position_top})")

    # Carregando a imagem de fundo
    try:
        image = Image.open(image_path)
        print(f"Imagem carregada com sucesso: {image_path}")
    except Exception as e:
        print(f"Erro ao carregar a imagem de fundo: {e}")
        return
    image = image.resize((900, 600))  # Redimensionando a imagem para o tamanho da janela
    bg_image = ImageTk.PhotoImage(image)
    print("Imagem de fundo redimensionada e convertida para PhotoImage")

    # Criando um rótulo para a imagem de fundo
    background_label = tk.Label(root, image=bg_image)
    background_label.place(relwidth=1, relheight=1)  # Preenchendo toda a janela com a imagem de fundo
    print("Rótulo da imagem de fundo criado e posicionado")

    # Adicionando o nome da aplicação no centro da tela com um visual mais sério e atraente
    app_name = tk.Label(root, text="PDF Analyzer", font=("Helvetica", 40, "bold"), fg="#FFFFFF", bg="#1E3D59", padx=20, pady=10, relief="raised", bd=10)
    app_name.place(relx=0.5, rely=0.25, anchor='center')
    print("Nome da aplicação adicionado ao centro da tela")

    # Função para iniciar a análise chamando o script blank_analyzer.py
    def iniciar_analise():
        print("Botão Iniciar Análise pressionado")
        root.destroy()
        print("Janela principal destruída")
        try:
            iniciar_analise_main()
            print("Função main do main.py chamada com sucesso")
        except Exception as e:
            print(f"Erro ao chamar a função main do main.py: {e}")
        current_directory = os.path.dirname(os.path.abspath(__file__))
        analyzer_path = os.path.join(current_directory, 'blank_analyzer.py')
        try:
            subprocess.Popen(['python', analyzer_path])
            print(f"Script blank_analyzer.py iniciado: {analyzer_path}")
        except Exception as e:
            print(f"Erro ao iniciar o script blank_analyzer.py: {e}")

    # Adicionando um botão para iniciar a análise com um design mais elegante
    start_button = tk.Button(root, text="Iniciar Análise", font=("Helvetica", 18, "bold"), fg="#FFFFFF", bg="#1E3D59", activebackground="#34495E", activeforeground="#FFFFFF", padx=20, pady=10, relief="raised", bd=5, command=iniciar_analise)
    start_button.place(relx=0.5, rely=0.55, anchor='center')
    print("Botão Iniciar Análise adicionado e posicionado")

    # Adicionando direitos autorais no final da tela
    copyright_label = tk.Label(root, text="Direitos Autorais © Wayster Cruz de Melo", font=("Helvetica", 12, "italic"), fg="#FFFFFF", bg="#1C2833", padx=5, pady=5)
    copyright_label.place(relx=0.01, rely=0.95, anchor='w')
    print("Rótulo de direitos autorais adicionado")

    # Iniciando o loop da interface gráfica
    print("Iniciando o loop da interface gráfica")
    root.mainloop()

if __name__ == "__main__":
    print("Executando tela_inicial.py como script principal")
    tela_inicial()