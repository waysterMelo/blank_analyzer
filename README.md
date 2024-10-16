# PDFAnalyzer

Um aplicativo GUI para analisar arquivos PDF digitalizados, identificar páginas em branco ou ruidosas e realizar OCR para extrair texto.

## Funcionalidades

- Selecionar um diretório contendo arquivos PDF.
- Analisar cada página dos PDFs para verificar se estão em branco ou possuem conteúdo.
- Realizar OCR nas páginas consideradas brancas para extrair texto.
- Gerar um relatório em Excel com os resultados da análise.
- Interface gráfica amigável para monitorar o progresso da análise.

## Requisitos

- Python 3.7 ou superior
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) instalado e configurado corretamente.

## Instalação

1. **Clone o repositório ou baixe os arquivos do projeto.**

2. **Abra o PyCharm e crie um novo projeto ou abra o diretório existente.**

3. **Configure um ambiente virtual e instale as dependências:**

   Abra o terminal no PyCharm e execute:

   ```bash
   pip install -r requirements.txt
