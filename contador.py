import os
import tkinter as tk
from tkinter import filedialog
import fitz  # PyMuPDF
from tkinter.filedialog import asksaveasfilename


def contar_paginas_pdf(caminho):
    try:
        pdf =\
            fitz.open(caminho)

        return len(pdf)
    except Exception as e:
        print(f"Erro ao processar {caminho}: {e}")
        return None

def processar_pdfs_em_pastas(raiz):
    # Converte tamanhos padrão para pontos (1 mm = 2.83465 pontos)
    mm_to_points = 2.83465
    A5 = (148 * mm_to_points) * (210 * mm_to_points)
    A4 = (210 * mm_to_points) * (297 * mm_to_points)
    A3 = (297 * mm_to_points) * (420 * mm_to_points)
    A2 = (420 * mm_to_points) * (594 * mm_to_points)
    A1 = (594 * mm_to_points) * (841 * mm_to_points)
    A0 = (841 * mm_to_points) * (1189 * mm_to_points)

    resultados = {
        "total_paginas": 0,
        "A5": 0,
        "A4": 0,
        "A3": 0,
        "A2": 0,
        "A1": 0,
        "A0": 0,
        "A00": 0,
    }

    for root, _, files in os.walk(raiz):
        for file in files:
            if file.lower().endswith(".pdf"):
                caminho_pdf = os.path.join(root, file)
                try:
                    pdf = fitz.open(caminho_pdf)
                    for page in pdf:
                        # Obtém largura e altura da página em pontos
                        width, height = page.rect.width, page.rect.height
                        area = width * height

                        # Classifica o tamanho da página
                        if area < A5:
                            resultados["A5"] += 1
                        elif area < A4:
                            resultados["A4"] += 1
                        elif area < A3:
                            resultados["A3"] += 1
                        elif area < A2:
                            resultados["A2"] += 1
                        elif area < A1:
                            resultados["A1"] += 1
                        elif area < A0:
                            resultados["A0"] += 1
                        else:
                            resultados["A00"] += 1

                    resultados["total_paginas"] += len(pdf)
                except Exception as e:
                    print(f"Erro ao processar {caminho_pdf}: {e}")

    return resultados

def gerar_relatorio(caminho_raiz):
    total_pastas = 0
    total_documentos = 0
    total_paginas = 0
    resultados = processar_pdfs_em_pastas(caminho_raiz)

    contador_geral = []
    resumo = []
    erros = []

    for raiz, _, arquivos in os.walk(caminho_raiz):
        documentos = 0
        paginas = 0
        pasta_contador = []

        for arquivo in arquivos:
            if arquivo.lower().endswith('.pdf'):
                caminho = os.path.join(raiz, arquivo)
                paginas_pdf = contar_paginas_pdf(caminho)

                if paginas_pdf:
                    documentos += 1
                    paginas += paginas_pdf
                    pasta_contador.append(f"{arquivo}: {paginas_pdf} páginas")

                else:
                    erros.append(f"Erro ao ler {caminho}")

        if documentos > 0:
            total_pastas += 1
            total_documentos += documentos
            total_paginas += paginas

            contador_geral.append(f"Pasta: {raiz}")
            contador_geral.extend(pasta_contador)
            contador_geral.append(f"Resumo da pasta: {raiz}, {documentos} documentos, {paginas} páginas\n")
            resumo.append(f"{raiz}: {documentos} documentos, {paginas} páginas")

    resumo.append(f"\nTotal: {total_pastas} pastas, {total_documentos} documentos, {total_paginas} páginas")
    contador_geral.append(f"\nTotal: {total_pastas} pastas, {total_documentos} documentos, {total_paginas} páginas")

    contador_geral.append("\nTamanhos de Página:")
    resumo.append("\nTamanhos de Página:")
    for key, value in resultados.items():
        contador_geral.append(f"{key}: {value} páginas")
        resumo.append(f"{key}: {value} páginas")

    return "\n".join(resumo), "\n".join(contador_geral), "\n".join(erros)
def salvar_txts():
    caminho_raiz = entrada_caminho.get()
    if not caminho_raiz:
        resultado.set("Por favor, selecione uma pasta.")
        return

    resumo_texto, contador_texto, erros_texto = gerar_relatorio(caminho_raiz)

    try:
        # Salvar Contador Geral
        caminho_contador = asksaveasfilename(
            title="Salvar Contador Geral",
            defaultextension=".txt",
            filetypes=[("Arquivos de texto", "*.txt")],
            initialfile="contador_geral.txt"
        )
        if caminho_contador:
            with open(caminho_contador, 'w') as f:
                f.write(contador_texto)

        # Salvar Resumo
        caminho_resumo = asksaveasfilename(
            title="Salvar Resumo",
            defaultextension=".txt",
            filetypes=[("Arquivos de texto", "*.txt")],
            initialfile="resumo.txt"
        )
        if caminho_resumo:
            with open(caminho_resumo, 'w') as f:
                f.write(resumo_texto)


        resultado.set("Arquivos TXT salvos com sucesso!")
    except Exception as e:
        resultado.set(f"Erro ao salvar os arquivos: {e}")
def selecionar_pasta():
    caminho = filedialog.askdirectory()
    entrada_caminho.delete(0, tk.END)
    entrada_caminho.insert(0, caminho)
    atualizar_arvore(caminho)

def iniciar_contagem():
    caminho_raiz = entrada_caminho.get()
    if not caminho_raiz:
        resultado.set("Por favor, selecionar uma pasta.")
        return
    resumo_texto, contador_texto, erros_texto = gerar_relatorio(caminho_raiz)
    resultado.set("Relatorios gerados com sucesso!")
    exibir_relatorio(resumo_texto)


def exibir_relatorio(conteudo):
    area_texto.config(state=tk.NORMAL)
    area_texto.delete('1.0', tk.END)
    area_texto.insert(tk.END, conteudo)
    area_texto.config(state=tk.DISABLED)


def alternar_relatorio(tipo):
    caminho_raiz = entrada_caminho.get()
    if not caminho_raiz:
        resultado.set("Por favor, selecione uma pasta.")
        return
    resumo_texto, contador_texto, erros_texto = gerar_relatorio(caminho_raiz)
    if tipo == 'resumo':
        exibir_relatorio(resumo_texto)
    elif tipo == 'contador':
        exibir_relatorio(contador_texto)
    elif tipo == 'erros':
        exibir_relatorio(erros_texto)


def atualizar_arvore(caminho_raiz):
    for item in arvore.get_children():
        arvore.delete(item)

    def inserir_pastas(pasta, parent=""):
        for item in os.listdir(pasta):
            caminho_completo = os.path.join(pasta, item)
            is_dir = os.path.isdir(caminho_completo)
            node = arvore.insert(parent, 'end', text=item, open=False)
            if is_dir:
                inserir_pastas(caminho_completo, node)

    inserir_pastas(caminho_raiz)

# Variável global para controlar a página atual
pagina_atual = 0

import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# Configuração da interface gráfica
janela = ttk.Window(themename="darkly")  # Tema moderno
janela.title("Visualizador de PDF")
janela.geometry("800x600")  # Definir tamanho inicial

# Layout da interface
frame_superior = ttk.Frame(janela, padding=10)
frame_superior.pack(fill=X)

ttk.Label(frame_superior, text="Caminho da pasta:", font=("Helvetica", 12)).pack(side=LEFT, padx=5)
entrada_caminho = ttk.Entry(frame_superior, width=50)
entrada_caminho.pack(side=LEFT, padx=5)
ttk.Button(frame_superior, text="Selecionar pasta", command=selecionar_pasta, bootstyle=PRIMARY).pack(side=LEFT, padx=5)
ttk.Button(frame_superior, text="Contar páginas", command=iniciar_contagem, bootstyle=SUCCESS).pack(side=LEFT, padx=5)

resultado = ttk.StringVar()
ttk.Label(janela, textvariable=resultado, font=("Helvetica", 12), padding=10).pack()

# Área da árvore de navegação
frame_principal = ttk.Frame(janela, padding=10)
frame_principal.pack(fill=BOTH, expand=True)

arvore = ttk.Treeview(frame_principal)
arvore.pack(side=LEFT, fill=Y, expand=False, padx=5, pady=5)

# Área de texto para exibição dos relatórios
area_texto = ttk.ScrolledText(frame_principal, width=60, height=20)
area_texto.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)

# Botões para alternar entre os relatórios
frame_botoes = ttk.Frame(janela, padding=10)
frame_botoes.pack(fill=X)

ttk.Button(frame_botoes, text="Mostrar Resumo", command=lambda: alternar_relatorio('resumo'), bootstyle=INFO).pack(
    side=LEFT, padx=5)
ttk.Button(frame_botoes, text="Mostrar Contador Geral", command=lambda: alternar_relatorio('contador'),
           bootstyle=SECONDARY).pack(side=LEFT, padx=5)
ttk.Button(frame_botoes, text="Mostrar Erros", command=lambda: alternar_relatorio('erros'), bootstyle=DANGER).pack(
    side=LEFT, padx=5)
ttk.Button(frame_botoes, text="Salvar Resumos", command=salvar_txts, bootstyle=SUCCESS).pack(side=RIGHT, padx=(0,100))


janela.mainloop()
