import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def selecionar_pasta():
    caminho = filedialog.askdirectory()
    if caminho:
        pasta_var.set(caminho)
        atualizar_lista()

def atualizar_lista():
    lista_arquivos.delete(*lista_arquivos.get_children())
    pasta = pasta_var.get()
    ext_filtro = ext_var.get().strip().lower()
    if not pasta:
        return
    arquivos = [f for f in os.listdir(pasta) if os.path.isfile(os.path.join(pasta, f))]
    for i, arquivo in enumerate(arquivos, start=1):
        if ext_filtro and not arquivo.lower().endswith(ext_filtro):
            continue
        lista_arquivos.insert("", "end", values=(arquivo, f"{prefixo_var.get()}_{i}{os.path.splitext(arquivo)[1]}"))

def renomear_arquivos():
    pasta = pasta_var.get()
    if not pasta:
        messagebox.showwarning("Aviso", "Selecione a pasta.")
        return
    prefixo = prefixo_var.get()
    if not prefixo:
        messagebox.showwarning("Aviso", "Insira o prefixo.")
        return
    for item in lista_arquivos.get_children():
        antigo, novo = lista_arquivos.item(item, "values")
        os.rename(os.path.join(pasta, antigo), os.path.join(pasta, novo))
    messagebox.showinfo("Sucesso", "Arquivos renomeados!")
    atualizar_lista()

# Janela principal
root = tk.Tk()
root.title("Renomeador de Arquivos Avançado")
root.geometry("600x400")

pasta_var = tk.StringVar()
prefixo_var = tk.StringVar()
ext_var = tk.StringVar()

# Layout
frame_top = tk.Frame(root)
frame_top.pack(pady=10)

tk.Label(frame_top, text="Pasta:").grid(row=0, column=0, sticky="w")
tk.Entry(frame_top, textvariable=pasta_var, width=50).grid(row=0, column=1, padx=5)
tk.Button(frame_top, text="Selecionar Pasta", command=selecionar_pasta).grid(row=0, column=2, padx=5)

tk.Label(frame_top, text="Prefixo:").grid(row=1, column=0, sticky="w")
tk.Entry(frame_top, textvariable=prefixo_var, width=20).grid(row=1, column=1, sticky="w")
tk.Label(frame_top, text="Filtrar Extensão (ex: .txt):").grid(row=2, column=0, sticky="w")
tk.Entry(frame_top, textvariable=ext_var, width=20).grid(row=2, column=1, sticky="w")

tk.Button(frame_top, text="Atualizar Lista", command=atualizar_lista).grid(row=2, column=2, padx=5)

# Tabela de arquivos
lista_arquivos = ttk.Treeview(root, columns=("Antigo", "Novo"), show="headings")
lista_arquivos.heading("Antigo", text="Nome Original")
lista_arquivos.heading("Novo", text="Novo Nome")
lista_arquivos.pack(expand=True, fill="both", padx=10, pady=10)

tk.Button(root, text="Renomear Arquivos", command=renomear_arquivos, bg="green", fg="white").pack(pady=10)

root.mainloop()
