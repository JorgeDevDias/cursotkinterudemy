from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import pandas as pd

def Ler_arquivo(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo)
        name_columns = df.columns.tolist()
        linhas = df.values.tolist()
        # print(name_columns)
        # print(linhas)
        return name_columns, linhas
    except Exception as e:
        print(f"Erro ao ler arquivo: {str(e)}")


def Pegar_arquivo():
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if caminho_arquivo:
        cabecalho, linhas = Ler_arquivo(caminho_arquivo)
        atualizar_tabela(cabecalho, linhas)

def atualizar_tabela(cabecalho, linhas):
    tabela["columns"] = cabecalho
    tabela["show"] = "headings"
    

    for col in cabecalho:
        tabela.heading(col, text=col)

    tabela.delete(*tabela.get_children())

    for linha in linhas:
        tabela.insert("", "end", values=linha)

    #Centralizar
    colunas = tabela["columns"]
    for coluna in colunas:
        tabela.heading(coluna, anchor=CENTER)
        tabela.column(coluna, anchor=CENTER)
    
janela = Tk()
height_janela = janela.winfo_screenheight() 
width_janela = janela.winfo_screenwidth() 
janela.geometry(f"{width_janela}x{height_janela}")
janela.title("Lista de Efetivo")

btn_abrir_arquivo = ttk.Button(janela, text="Abrir Arquivo", command=Pegar_arquivo)
btn_abrir_arquivo.grid(column=0, row=0, columnspan=3)

frm = ttk.Frame(janela, padding=10)
# frm.grid(row=1, column=0, sticky="nsew")
frm.pack(side="top")

estilo_tabela = ttk.Style()
estilo_tabela.theme_use("alt")
estilo_tabela.configure(".",font="Arial 12")
tabela = ttk.Treeview(frm)
tabela.pack(side="top")

#tabela.place(relx=0.01,rely=0.1,relwidth=0.98,relheight=0.98)

frm.columnconfigure(0, weight=1)
frm.rowconfigure(1, weight=1)

# Adiciona uma barra de rolagem vertical
scroll_y = ttk.Scrollbar(janela, orient="vertical", command=tabela.yview)
scroll_y.grid(column=1,row=1,  sticky="nsew")
tabela.configure(yscrollcommand=scroll_y.set)

janela.mainloop()