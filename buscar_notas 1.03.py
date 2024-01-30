import os
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from pdf2image import convert_from_path
import pytesseract
import shutil

def processar_imagem(imagem, caminho_imagem_temporaria):
    imagem.save(caminho_imagem_temporaria, 'PNG', quality=50)

def processar_todas_as_imagens(imagens):
    for i, imagem in enumerate(imagens):
        caminho_imagem_temporaria = f"imagem_temp_{i}.png"
        processar_imagem(imagem, caminho_imagem_temporaria)

def buscar_notas_fiscais(caminho_pasta, expressoes_incluir, expressoes_excluir):
    notas_encontradas = []

    for root, dirs, files in os.walk(caminho_pasta):
        for arquivo in files:
            if arquivo.endswith(".pdf"):
                caminho_arquivo = os.path.join(root, arquivo)

                try:
                    imagens = convert_from_path(caminho_arquivo)
                    processar_todas_as_imagens(imagens)

                    for i, imagem in enumerate(imagens):
                        caminho_imagem_temporaria = f"imagem_temp_{i}.png"
                        texto_extrai = pytesseract.image_to_string(imagem)

                        incluir = any(expr.lower() in texto_extrai.lower() for expr in expressoes_incluir)
                        excluir = any(expr.lower() in texto_extrai.lower() for expr in expressoes_excluir)

                        if incluir and not excluir:
                            notas_encontradas.append(caminho_arquivo)
                            break

                        os.remove(caminho_imagem_temporaria)

                except Exception as e:
                    print(f"Erro ao processar PDF {caminho_arquivo}: {e}")

    return notas_encontradas

def criar_relatorio(notas_encontradas, caminho_relatorio):
    with open(caminho_relatorio, 'w', encoding='utf-8') as relatorio:
        if notas_encontradas:
            relatorio.write("PDFs encontrados com as expressões especificadas:\n")
            for nota in notas_encontradas:
                relatorio.write(f"- {nota}\n")
        else:
            relatorio.write("Nenhum PDF encontrado com as expressões especificadas.\n")

def buscar_pasta():
    pasta = filedialog.askdirectory()
    entry_pasta_var.set(pasta)

def buscar_pasta_destino():
    pasta_destino = filedialog.askdirectory()
    entry_pasta_destino_var.set(pasta_destino)

def processar():
    pasta_pdfs = entry_pasta_var.get()
    pasta_destino = entry_pasta_destino_var.get()
    expressoes_incluir = entry_incluir_var.get().split(',')
    expressoes_excluir = entry_excluir_var.get().split(',')

    try:
        pdfs_encontrados = buscar_notas_fiscais(pasta_pdfs, expressoes_incluir, expressoes_excluir)

        if pdfs_encontrados:
            caminho_relatorio = os.path.join(pasta_pdfs, 'relatorio_encontrados.txt')
            criar_relatorio(pdfs_encontrados, caminho_relatorio)
            resultado_text.config(state=tk.NORMAL)
            resultado_text.delete('1.0', tk.END)
            resultado_text.insert(tk.END, "PDFs encontrados com as expressões especificadas:\n")
            for pdf in pdfs_encontrados:
                resultado_text.insert(tk.END, f"{pdf}\n")
                nome_arquivo = os.path.basename(pdf)
                destino_arquivo = os.path.join(pasta_destino, nome_arquivo)
                shutil.move(pdf, destino_arquivo)
            resultado_text.insert(tk.END, f"\nRelatório salvo em: {caminho_relatorio}\nArquivos movidos para: {pasta_destino}")
            resultado_text.config(state=tk.DISABLED)
        else:
            resultado_text.config(state=tk.NORMAL)
            resultado_text.delete('1.0', tk.END)
            resultado_text.insert(tk.END, "Nenhum PDF encontrado com as expressões especificadas\n")
            resultado_text.insert(tk.END, "\nOperação concluída. Nenhum arquivo foi movido para a pasta de destino.")
            resultado_text.config(state=tk.DISABLED)
    except Exception as e:
        print(f"Erro ao processar documentos: {e}")

# Configuração da interface gráfica
root = tk.Tk()
root.title("Busca de PDFs com Expressões")

# Componentes da interface
lbl_pasta = ttk.Label(root, text="Selecione a pasta contendo os documentos:")
lbl_pasta.grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)

entry_pasta_var = tk.StringVar()
entry_pasta = ttk.Entry(root, textvariable=entry_pasta_var, state="readonly")
entry_pasta.grid(row=0, column=1, padx=10, pady=5)

btn_selecionar_pasta = ttk.Button(root, text="Selecionar Pasta", command=buscar_pasta)
btn_selecionar_pasta.grid(row=0, column=2, padx=10, pady=5)

lbl_pasta_destino = ttk.Label(root, text="Selecione a pasta de destino:")
lbl_pasta_destino.grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)

entry_pasta_destino_var = tk.StringVar()
entry_pasta_destino = ttk.Entry(root, textvariable=entry_pasta_destino_var, state="readonly")
entry_pasta_destino.grid(row=1, column=1, padx=10, pady=5)

btn_selecionar_pasta_destino = ttk.Button(root, text="Selecionar Destino", command=buscar_pasta_destino)
btn_selecionar_pasta_destino.grid(row=1, column=2, padx=10, pady=5)

lbl_incluir = ttk.Label(root, text="Expressões a serem incluídas (separadas por vírgula):")
lbl_incluir.grid(row=2, column=0, sticky=tk.W, padx=10, pady=5)

entry_incluir_var = tk.StringVar()
entry_incluir = ttk.Entry(root, textvariable=entry_incluir_var)
entry_incluir.grid(row=2, column=1, padx=10, pady=5)

lbl_excluir = ttk.Label(root, text="Expressões a serem excluídas (separadas por vírgula):")
lbl_excluir.grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)

entry_excluir_var = tk.StringVar()
entry_excluir = ttk.Entry(root, textvariable=entry_excluir_var)
entry_excluir.grid(row=3, column=1, padx=10, pady=5)

btn_processar = ttk.Button(root, text="Processar Documentos", command=processar)
btn_processar.grid(row=4, column=0, columnspan=3, pady=10)

resultado_text = tk.Text(root, height=10, width=50, state=tk.DISABLED)
resultado_text.grid(row=5, column=0, columnspan=3, pady=5, padx=10)

root.mainloop()