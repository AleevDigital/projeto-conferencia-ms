import time
import zipfile
import pandas as pd
import win32com.client as win32
import os
import tratar_plan as tp
import totalizador_d2 as td2
import totalizador_d4 as td4
import tkinter as tk
from tkinter import filedialog, simpledialog

# === INTERFACE GRÁFICA PARA SELEÇÃO DE ARQUIVOS ===
root = tk.Tk()
root.withdraw()  # esconde a janela principal

empresa = simpledialog.askstring("Empresa", "Digite o nome da empresa:")

# Seleção dos arquivos com janela de navegação
caminho_d2 = filedialog.askopenfilename(title="Selecione o arquivo D2")
caminho_d4 = filedialog.askopenfilename(title="Selecione o arquivo D4")
caminho_dca = filedialog.askopenfilename(title="Selecione o arquivo DCA")
caminho_planilha_totalizadora = filedialog.askopenfilename(title="Selecione a Planilha de Conferência")

# === TRATAMENTO DAS MATRIZES ===
d2_tratado = tp.tratar_matriz(caminho_d2, 'd2')
d4_tratado = tp.tratar_matriz(caminho_d4, 'd4')

# === TOTALIZAÇÃO NAS PLANILHAS ===
td2.totalizador_d2(d2_tratado, caminho_planilha_totalizadora)
td4.totalizador_d4(d4_tratado, caminho_planilha_totalizadora)


caminho_d2_tratado = fr'output\D2\D2_Tratado.xlsx'
caminho_d4_tratado = fr'output\D4\D4_Tratado.xlsx'

d2_tratado.to_excel(caminho_d2_tratado, index=False)
d4_tratado.to_excel(caminho_d4_tratado, index=False)

time.sleep(3)

# === CRIAÇÃO DO ZIP COM OS ARQUIVOS FINALIZADOS ===
arquivos_para_zipar = [
    caminho_d2_tratado,
    caminho_d4_tratado,
    caminho_planilha_totalizadora
]

output_path = fr'C:\Users\aleal\Downloads\Matriz Saldo_D2_D4\Output\{empresa}.zip'

with zipfile.ZipFile(output_path, 'w') as zipf:
    for caminho_completo in arquivos_para_zipar:
        nome_arquivo = os.path.basename(caminho_completo)
        zipf.write(caminho_completo, arcname=nome_arquivo)
        print(f'Arquivo adicionado: {nome_arquivo}')

print(f'\n✅ Arquivos compactados com sucesso em:\n{output_path}')