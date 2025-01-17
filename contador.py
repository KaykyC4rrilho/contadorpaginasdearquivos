import os
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
import fitz  # PyMuPDF
from PIL import Image, ImageTk


def contar_paginas_pdf(caminho):
    try:
        pdf =\
            fitz.open(caminho)

        return len(pdf)
    except Exception as e:
        print(f"Erro ao processar {caminho}: {e}")
        return None
def contar_por_tamanho_pdf(caminho_pdf):
    try:
        pdf = fitz.open(caminho_pdf)
        tamanho_pagina = {"A0": 0, "A1": 0, "A2": 0, "A3": 0, "A4": 0, "Menor que A4": 0, "Maior que A0": 0}

        for pagina in pdf:
            largura = pagina.rect.width
            altura = pagina.rect.height

            # Considerando orientação retrato ou paisagem
            largura, altura = sorted([largura, altura])

            # Maior que A0 (largura > 2384mm e altura > 3370mm)
            if largura > 2384 and altura > 3370:
                tamanho_pagina["Maior que A0"] += 1
            elif largura == 2384 and altura == 3370:
                tamanho_pagina["A0"] += 1
            elif largura >= 1684 and altura >= 2384:
                tamanho_pagina["A1"] += 1
            elif largura >= 1190 and altura >= 1684:
                tamanho_pagina["A2"] += 1
            elif largura >= 841 and altura >= 1190:
                tamanho_pagina["A3"] += 1
            elif largura >= 595 and altura >= 841:
                tamanho_pagina["A4"] += 1
            elif largura < 595 and altura < 841:
                tamanho_pagina["Menor que A4"] += 1
            elif largura > 2384 or altura > 3370:
                tamanho_pagina["Maior que A0"] += 1

        return tamanho_pagina
    except Exception as e:
        print(f"Erro ao processar o PDF: {e}")
        return None
def gerar_relatorio(caminho_raiz):
    total_pastas = 0
    total_documentos = 0
    total_paginas = 0

    contador_geral = []
    resumo = []
    erros = []

    tamanho_pagina_total = {"A0": 0, "A1": 0, "A2": 0, "A3": 0, "A4": 0, "Menor que A4": 0, "Maior que A0": 0}

    for raiz, diretorios, arquivos in os.walk(caminho_raiz):
        documentos = 0
        paginas = 0
        pasta_contador = []

        for arquivo in arquivos:
            if arquivo.lower().endswith('.pdf'):
                caminho = os.path.join(raiz, arquivo)
                paginas_pdf = contar_paginas_pdf(caminho)

                if paginas_pdf is not None:
                    documentos += 1
                    paginas += paginas_pdf
                    pasta_contador.append(f"{arquivo}: {paginas_pdf}")
                    tamanho_pdf = contar_por_tamanho_pdf(caminho)
                    for key in tamanho_pdf:
                        tamanho_pagina_total[key] += tamanho_pdf[key]
                else:
                    erros.append(f"Erro ao Ler {caminho}")

        if documentos > 0:
            total_pastas += 1
            total_documentos += documentos
            total_paginas += paginas

            contador_geral.append(f"Pasta: {raiz}")
            contador_geral.extend((pasta_contador))
            contador_geral.append(f"Resumo da pasta: {raiz}, {documentos} documentos, {paginas} paginas\n")
            resumo.append(f"{raiz}: {documentos} documentos, {paginas} páginas")

    resumo.append(f"\nTotal: {total_pastas} pastas, {total_documentos} documentos, {total_paginas} páginas")
    contador_geral.append(f"\nTotal: {total_pastas} pastas, {total_documentos} documentos, {total_paginas} páginas")

    contador_geral.append(f"\nTamanhos de Página:")
    resumo.append(f"\nTamanhos de Página:")
    for key, value in tamanho_pagina_total.items():
        contador_geral.append(f"{key}: {value} páginas")
        resumo.append(f"{key}: {value} páginas")

    with open('contador_geral.txt', 'w') as f:
        f.write("\n".join(contador_geral))

    with open('resumo.txt', 'w') as f:
        f.write("\n".join(resumo))

    with open('erros.txt', 'w') as f:
        f.write("\n".join(erros))

    return "\n".join(resumo), "\n".join(contador_geral), "\n".join(erros)



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
    area_texto.delete('1.0', tk.END)
    area_texto.insert(tk.END, conteudo)


def alternar_relatorio(tipo):
    caminho_raiz = entrada_caminho.get()
    if not caminho_raiz:
        resultado.set("Por favor, selecione uma pasta.")
        return
    _, contador_texto, erros_texto = gerar_relatorio(caminho_raiz)

    if tipo == 'resumo':
        resumo_texto, _, _ = gerar_relatorio(caminho_raiz)
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


# Função para exibir o PDF como imagem
def exibir_pdf(caminho_pdf):
    try:
        # Abrir o arquivo PDF usando PyMuPDF
        pdf = fitz.open(caminho_pdf)
        global paginas_pdf  # Variável global para armazenar as páginas
        paginas_pdf = []

        # Converte todas as páginas para imagens e armazena em uma lista
        for i in range(pdf.page_count):
            pagina = pdf[i]
            imagem_pagina = pagina.get_pixmap()
            imagem_tk = ImageTk.PhotoImage(
                Image.frombytes("RGB", (imagem_pagina.width, imagem_pagina.height), imagem_pagina.samples))
            paginas_pdf.append(imagem_tk)

        # Exibe a primeira página
        mostrar_pagina(0)
    except Exception as e:
        exibir_relatorio(f"Erro ao abrir o PDF: {e}")


# Função para mostrar uma página específica
def mostrar_pagina(indice):
    if 0 <= indice < len(paginas_pdf):
        imagem_tk = paginas_pdf[indice]
        label_imagem.configure(image=imagem_tk)
        label_imagem.image = imagem_tk  # Mantém a referência para a imagem


# Função para selecionar o item da árvore e abrir o PDF
def item_selecionado(event):
    item = arvore.selection()
    if item:
        caminho_item = arvore.item(item[0])['text']
        caminho_completo = os.path.join(entrada_caminho.get(), caminho_item)
        caminho_completo = os.path.normpath(caminho_completo)
        if caminho_item.lower().endswith('.pdf'):
            exibir_pdf(caminho_completo)


# Função para navegar para a próxima página
def proxima_pagina():
    global pagina_atual
    if pagina_atual < len(paginas_pdf) - 1:
        pagina_atual += 1
        mostrar_pagina(pagina_atual)


# Função para navegar para a página anterior
def pagina_anterior():
    global pagina_atual
    if pagina_atual > 0:
        pagina_atual -= 1
        mostrar_pagina(pagina_atual)


# Variável global para controlar a página atual
pagina_atual = 0

# Configuração da interface gráfica
janela = tk.Tk()
janela.title("Visualizador de PDF")

# Layout da interface
tk.Label(janela, text="Caminho da pasta:").grid(row=0, column=1, padx=5, pady=5)
entrada_caminho = tk.Entry(janela, width=50)
entrada_caminho.grid(row=0, column=2, padx=5, pady=5)

tk.Button(janela, text="Selecionar pasta", command=selecionar_pasta).grid(row=0, column=3, padx=5, pady=5)

tk.Button(janela, text="Contar páginas", command=iniciar_contagem).grid(row=1, column=2, padx=5, pady=5)

resultado = tk.StringVar()
tk.Label(janela, textvariable=resultado).grid(row=2, column=2, padx=5, pady=5)

# Área da árvore de navegação
arvore = ttk.Treeview(janela)
arvore.grid(row=3, column=0, rowspan=3, padx=5, pady=5, sticky='ns')
arvore.bind("<Double-1>", item_selecionado)  # Associa o clique duplo

# Área de texto para exibição dos relatórios
area_texto = scrolledtext.ScrolledText(janela, width=60, height=20)
area_texto.grid(row=3, column=1, columnspan=3, padx=5, pady=5)

# Botões para alternar entre os relatórios
tk.Button(janela, text="Mostrar Resumo", command=lambda: alternar_relatorio('resumo')).grid(row=4, column=1, padx=5,
                                                                                            pady=5)
tk.Button(janela, text="Mostrar Contador Geral", command=lambda: alternar_relatorio('contador')).grid(row=4, column=2,
                                                                                                      padx=5, pady=5)
tk.Button(janela, text="Mostrar Erros", command=lambda: alternar_relatorio('erros')).grid(row=4, column=3, padx=5,
                                                                                          pady=5)

# Label para exibir o PDF como imagem
#label_imagem = tk.Label(janela)
#label_imagem.grid(row=3, column=4, rowspan=3, padx=5, pady=5, sticky="e")  # Alinha à direita com 'sticky="e"'

# Botões de navegação
#tk.Button(janela, text="Anterior", command=pagina_anterior).grid(row=5, column=4, padx=5, pady=5)
#tk.Button(janela, text="Próxima", command=proxima_pagina).grid(row=6, column=4, padx=5, pady=5)

janela.mainloop()