import os
import re
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2image import convert_from_path
import pytesseract
from fpdf import FPDF

# ===============================
# CONFIGURAÇÕES
# ===============================
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
CAMINHO_POPPLER = r"C:\poppler-25.11.0\Library\bin"

# ===============================
# VARIÁVEL DE CONTROLE
# ===============================
cancelar = False

# ===============================
# FUNÇÕES
# ===============================
def extrair_chave_pdf(pdf_path, palavra_chave):
    try:
        paginas = convert_from_path(pdf_path, poppler_path=CAMINHO_POPPLER)
        texto_total = ""
        for pagina in paginas:
            texto_total += pytesseract.image_to_string(pagina, lang="por")
        if palavra_chave:
            padrao = rf"{palavra_chave}[: ]+([A-Za-z0-9\-\.]+)"
            encontrado = re.search(padrao, texto_total, re.IGNORECASE)
            return encontrado.group(1) if encontrado else None
        return None
    except Exception as e:
        return None

def extrair_chave_imagem(img_path, palavra_chave):
    try:
        texto = pytesseract.image_to_string(img_path, lang="por")
        if palavra_chave:
            padrao = rf"{palavra_chave}[: ]+([A-Za-z0-9\-\.]+)"
            encontrado = re.search(padrao, texto, re.IGNORECASE)
            return encontrado.group(1) if encontrado else None
        return None
    except Exception as e:
        return None

def atualizar_lista():
    lista_arquivos.delete(*lista_arquivos.get_children())
    pasta = pasta_var.get()
    ext_filtro = ext_var.get().strip().lower()
    if not pasta:
        return
    arquivos = [f for f in os.listdir(pasta) if os.path.isfile(os.path.join(pasta, f))]
    for arquivo in arquivos:
        if ext_filtro and not arquivo.lower().endswith(ext_filtro):
            continue
        # Apenas mostra o arquivo, sem processar OCR
        lista_arquivos.insert("", "end", values=(arquivo, "Aguardando..."))


def processar_ocr_lista():
    global cancelar
    pasta = pasta_var.get()
    palavra = palavra_chave_var.get()
    nome_base = nome_base_var.get()
    for item in lista_arquivos.get_children():
        if cancelar:
            return
        antigo, _ = lista_arquivos.item(item, "values")
        caminho = os.path.join(pasta, antigo)
        novo_nome = antigo
        chave = None
        if antigo.lower().endswith(".pdf"):
            chave = extrair_chave_pdf(caminho, palavra)
        elif antigo.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
            chave = extrair_chave_imagem(caminho, palavra)
            # Garante que vai ser PDF
            if chave:
                novo_nome = f"{nome_base}_{chave}.pdf" if nome_base else f"{chave}.pdf"
            else:
                novo_nome = f"{nome_base}_{os.path.splitext(antigo)[0]}.pdf" if nome_base else f"{os.path.splitext(antigo)[0]}.pdf"
        elif chave:
            novo_nome = f"{nome_base}_{chave}.pdf" if nome_base else f"{chave}.pdf"
        root.after(0, lambda i=item, a=antigo, n=novo_nome:
                   lista_arquivos.item(i, values=(a, n)))

def renomear_arquivos():
    pasta = pasta_var.get()
    if not pasta:
        messagebox.showwarning("Aviso", "Selecione a pasta.")
        return

    backup = backup_var.get()
    backup_path = os.path.join(pasta, "BACKUP") if backup else None
    if backup and not os.path.exists(backup_path):
        os.makedirs(backup_path)

    for item in lista_arquivos.get_children():
        antigo, novo = lista_arquivos.item(item, "values")
        caminho_antigo = os.path.join(pasta, antigo)
        caminho_novo = os.path.join(pasta, novo)

        if backup:
            shutil.copy2(caminho_antigo, os.path.join(backup_path, antigo))

        # Se o arquivo original é imagem → converte para PDF
        if antigo.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".tiff")):
            pdf = FPDF()
            pdf.add_page()
            pdf.image(caminho_antigo, x=10, y=10, w=180)
            pdf.output(caminho_novo)
            os.remove(caminho_antigo)  # opcional: remove a imagem original
        # Se é PDF → renomeia
        elif antigo.lower().endswith(".pdf"):
            if not os.path.exists(caminho_novo):
                os.rename(caminho_antigo, caminho_novo)
            else:
                messagebox.showwarning("Aviso", f"Arquivo {novo} já existe.")

    messagebox.showinfo("Sucesso", "Arquivos renomeados e salvos!")
    atualizar_lista()

def selecionar_pasta():
    caminho = filedialog.askdirectory()
    if caminho:
        pasta_var.set(caminho)
        atualizar_lista()

def iniciar_processamento():
    global cancelar
    cancelar = False
    threading.Thread(target=processar_ocr_lista, daemon=True).start()

def cancelar_processamento():
    global cancelar
    cancelar = True
    messagebox.showinfo("Cancelado", "Processamento interrompido.")

# ===============================
# INTERFACE GRÁFICA
# ===============================
root = tk.Tk()
root.title("Renomeador Avançado de PDFs com OCR")
root.geometry("900x600")

pasta_var = tk.StringVar()
palavra_chave_var = tk.StringVar()
nome_base_var = tk.StringVar()
ext_var = tk.StringVar()
backup_var = tk.BooleanVar()

frame_top = tk.Frame(root)
frame_top.pack(pady=10)

tk.Label(frame_top, text="Pasta:").grid(row=0, column=0, sticky="w")
tk.Entry(frame_top, textvariable=pasta_var, width=50).grid(row=0, column=1, padx=5)
tk.Button(frame_top, text="Selecionar Pasta", command=selecionar_pasta).grid(row=0, column=2, padx=5)

tk.Label(frame_top, text="Palavra-chave:").grid(row=1, column=0, sticky="w")
tk.Entry(frame_top, textvariable=palavra_chave_var, width=20).grid(row=1, column=1, sticky="w")
tk.Label(frame_top, text="Nome base (opcional):").grid(row=2, column=0, sticky="w")
tk.Entry(frame_top, textvariable=nome_base_var, width=30).grid(row=2, column=1, sticky="w")

tk.Label(frame_top, text="Filtrar Extensão (ex: .pdf):").grid(row=3, column=0, sticky="w")
tk.Entry(frame_top, textvariable=ext_var, width=20).grid(row=3, column=1, sticky="w")
tk.Checkbutton(frame_top, text="Salvar PDFs originais em BACKUP", variable=backup_var).grid(row=4, column=1, sticky="w")
tk.Button(frame_top, text="Atualizar Lista", command=atualizar_lista).grid(row=3, column=2, padx=5)

lista_arquivos = ttk.Treeview(root, columns=("Antigo", "Novo"), show="headings")
lista_arquivos.heading("Antigo", text="Nome Original")
lista_arquivos.heading("Novo", text="Novo Nome")
lista_arquivos.pack(expand=True, fill="both", padx=10, pady=10)

frame_botoes = tk.Frame(root)
frame_botoes.pack(pady=5)
tk.Button(frame_botoes, text="Iniciar Processamento", bg="blue", fg="white",
          command=iniciar_processamento).grid(row=0, column=0, padx=10)
tk.Button(frame_botoes, text="Cancelar", bg="red", fg="white",
          command=cancelar_processamento).grid(row=0, column=1, padx=10)

tk.Button(root, text="Renomear Arquivos", command=renomear_arquivos, bg="green", fg="white").pack(pady=10)

root.mainloop()
