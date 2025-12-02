# dados.py
import csv
import os
import random
from datetime import datetime
import config

def garantir_csv(caminho):
    if not os.path.exists(caminho):
        try:
            pasta = os.path.dirname(caminho)
            if not os.path.exists(pasta): return
            with open(caminho, 'w', newline='', encoding='utf-8') as f:
                csv.writer(f).writerow(["Data", "Codigo", "Prefixo", "Numero", "Tipo", "Projeto", "Titulo", "Descricao", "Status", "Caminho"])
        except: pass

def ler_registros(caminho_csv, mostrar_inativos=False, termo_busca=""):
    registros = []
    if os.path.exists(caminho_csv):
        try:
            # Tenta UTF-8
            with open(caminho_csv, 'r', encoding='utf-8') as f:
                reader = csv.reader(f); next(reader, None); registros = list(reader)
        except UnicodeDecodeError:
            try:
                # Fallback para ANSI
                with open(caminho_csv, 'r', encoding='cp1252') as f:
                    reader = csv.reader(f); next(reader, None); registros = list(reader)
            except: pass
    
    # Filtros
    resultados = []
    termo = termo_busca.lower()
    for row in reversed(registros):
        while len(row) < 10: row.append("")
        status = row[8]
        
        if status == "INATIVO" and not mostrar_inativos: continue
        
        txt_busca = (row[1] + row[6] + row[7]).lower()
        if termo in txt_busca:
            resultados.append(row)
    return resultados

def gravar_linha(caminho_csv, dados_linha):
    with open(caminho_csv, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(dados_linha)

def gerar_codigo_unico(caminho_csv, prefixo, sufixo):
    existentes = set()
    if os.path.exists(caminho_csv):
        try:
            with open(caminho_csv, 'r', encoding='utf-8') as f:
                reader = csv.reader(f); next(reader, None)
                for r in reader: existentes.add(r[1])
        except: pass
    
    for _ in range(1000):
        cod = f"{prefixo}-{random.randint(10000, 99999)}{sufixo}"
        if cod not in existentes:
            return cod
    return None

def excluir_logico(caminho_csv, codigo, caminho_arquivo):
    linhas = []
    if os.path.exists(caminho_csv):
        try:
            with open(caminho_csv, 'r', encoding='utf-8') as f:
                reader = csv.reader(f); linhas = list(reader)
        except: return
    
    with open(caminho_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        for row in linhas:
            while len(row) < 10: row.append("")
            if len(row) > 9 and row[1] == codigo and row[9] == caminho_arquivo:
                row[8] = "INATIVO"
            writer.writerow(row)