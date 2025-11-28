

# este codigo funciona apenas com arquivos na mesma pasta que a main #
# NÃO TEM INTERFACE GRAFICA - TE LIGA!!!


import os
import re
from pdf2image import convert_from_path
import pytesseract
from openpyxl import Workbook

os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

# Caminho do Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Caminho do Poppler
CAMINHO_POPPLER = r"C:\poppler-25.11.0\Library\bin"

# Pasta com os PDFs
PASTA = os.path.dirname(os.path.abspath(__file__))

# Regex para localizar o número de série
PADRAO_SERIE = r"(?:Série|Serie|Serial|Nº de Série|Numero de Serie)[: ]+([A-Za-z0-9\-\.]+)"


# FUNÇÕES

def extrair_serie(pdf_path):
    try:
        paginas = convert_from_path(pdf_path, poppler_path=CAMINHO_POPPLER)
        texto_total = ""
        for pagina in paginas:
            texto_total += pytesseract.image_to_string(pagina, lang="por")
        
        encontrado = re.search(PADRAO_SERIE, texto_total, re.IGNORECASE)
        return encontrado.group(1) if encontrado else None
    
    except Exception as e:
        print(f"Erro ao processar {pdf_path}: {e}")
        return None


def renomear_pdfs():
    resultados = []  # Guardará (arquivo_original, serie)

    for arquivo in os.listdir(PASTA):
        if arquivo.lower().endswith(".pdf"):
            caminho = os.path.join(PASTA, arquivo)
            print(f"Lendo: {arquivo}")

            serie = extrair_serie(caminho)

            # Armazena resultado para exportação
            resultados.append((arquivo, serie))

            if serie:
                novo_nome = f"RAT ATESTE 7876 - CIAUS - ALMOXARIFADO PORTO ALEGRE SERIE - {serie}.pdf"
                novo_caminho = os.path.join(PASTA, novo_nome)

                if not os.path.exists(novo_caminho):
                    os.rename(caminho, novo_caminho)
                    print(f"Renomeado para: {novo_nome}")
                else:
                    print(f"⚠ Arquivo {novo_nome} já existe.")
            else:
                print("❌ Número de série não encontrado.")

    exportar_excel(resultados)


def exportar_excel(dados):
    """Cria um Excel com o nome original do PDF e a série encontrada."""
    caminho_excel = os.path.join(PASTA, "resultado.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Séries Encontradas"

    ws.append(["Arquivo Original", "Série Encontrada"])

    for arquivo, serie in dados:
        ws.append([arquivo, serie if serie else "NÃO ENCONTRADO"])

    wb.save(caminho_excel)
    print(f"\n📄 Excel criado: {caminho_excel}")


# EXECUÇÃO
if __name__ == "__main__":
    renomear_pdfs()
