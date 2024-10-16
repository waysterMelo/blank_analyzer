import tkinter as tk
from PIL import Image, ImageTk
from tesseract_config import TesseractConfig
from gui import PDFAnalyzerGUI

# Caminho para a imagem de fundo
image_path = 'img/logo.webp'

def configurar_tesseract():
    # Configuração do Tesseract OCR
    tessdata_prefix = r'C:/Program Files/Tesseract-OCR/tessdata/'
    tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
    tesseract_config = TesseractConfig(tessdata_prefix, tesseract_cmd)
    tesseract_config.test_setup()

def iniciar_interface_principal():
    # Criando a janela principal
    print("Iniciando a criação da janela principal")
    root = tk.Tk()
    root.title("PDF Analyzer Blank")
    root.geometry("900x600")
    root.configure(bg="#1C2833")
    root.resizable(False, False)
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
        image = image.resize((900, 600))  # Redimensiona a imagem para o tamanho da janela
        bg_image = ImageTk.PhotoImage(image)
        print("Imagem de fundo carregada com sucesso.")
    except Exception as e:
        print(f"Erro ao carregar a imagem de fundo: {e}")
        return

    # Rótulo para a imagem de fundo
    background_label = tk.Label(root, image=bg_image)
    background_label.place(relwidth=1, relheight=1)

    # Nome da aplicação no centro
    app_name = tk.Label(root, text="PDF Analyzer Blank", font=("Helvetica", 40, "bold"), fg="#FFFFFF", bg="#1E3D59", padx=20, pady=10, relief="raised", bd=10)
    app_name.place(relx=0.5, rely=0.25, anchor='center')

    # Função para iniciar a análise e carregar GUI
    def iniciar_analise():
        print("Botão Iniciar Análise pressionado")
        root.destroy()
        configurar_tesseract()
        PDFAnalyzerGUI()
        print("Interface principal do PDFAnalyzerGUI iniciada.")

    # Botão para iniciar a análise
    start_button = tk.Button(root, text="Iniciar Análise", font=("Helvetica", 18, "bold"), fg="#FFFFFF", bg="#1E3D59", activebackground="#34495E", activeforeground="#FFFFFF", padx=20, pady=10, relief="raised", bd=5, command=iniciar_analise)
    start_button.place(relx=0.5, rely=0.55, anchor='center')

    # Direitos autorais no final da tela
    copyright_label = tk.Label(root, text="Direitos Autorais © Wayster Cruz de Melo", font=("Helvetica", 12, "italic"), fg="#FFFFFF", bg="#1C2833", padx=5, pady=5)
    copyright_label.place(relx=0.01, rely=0.95, anchor='w')

    # Inicia o loop da interface gráfica
    print("Iniciando o loop da interface gráfica")
    root.mainloop()

if __name__ == "__main__":
    print("Executando tela_inicial.py como script principal")
    iniciar_interface_principal()