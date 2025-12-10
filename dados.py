import csv
import os
import shutil
import random
import config

def garantir_csv(caminho):
    if not os.path.exists(caminho):
        try:
            pasta = os.path.dirname(caminho)
            if not os.path.exists(pasta): return
            with open(caminho, 'w', newline='', encoding='utf-8') as f:
                csv.writer(f).writerow(["Data", "Codigo", "Prefixo", "Numero", "Tipo", "Projeto", "Titulo", "Descricao", "Status", "Caminho"])
        except: pass

def ler_registros(caminho_csv, mostrar_inativos=False, termo_busca="", ocultar_desenhos=False):
    registros = []
    if os.path.exists(caminho_csv):
        try:
            with open(caminho_csv, 'r', encoding='utf-8') as f:
                reader = csv.reader(f); next(reader, None); registros = list(reader)
        except UnicodeDecodeError:
            try:
                with open(caminho_csv, 'r', encoding='cp1252') as f:
                    reader = csv.reader(f); next(reader, None); registros = list(reader)
            except: pass
    
    resultados = []
    termo = termo_busca.lower()
    for row in reversed(registros):
        while len(row) < 10: row.append("")
        status = row[8]
        tipo = row[4]
        
        # Filtra inativos
        if status == "INATIVO" and not mostrar_inativos: continue
        # Filtra desenhos
        if ocultar_desenhos and "DESENHO" in tipo.upper(): continue
        
        txt_busca = (row[1] + row[6] + row[7]).lower()
        if termo in txt_busca:
            resultados.append(row)
    return resultados

def gravar_linha(caminho_csv, dados_linha):
    with open(caminho_csv, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f); writer.writerow(dados_linha)

def editar_registro(caminho_csv, codigo, novos_dados):
    """
    Edita os dados e altera o STATUS para 'MODIFICADO' para persistir a cor.
    """
    linhas = []
    alterado = False
    if os.path.exists(caminho_csv):
        try:
            with open(caminho_csv, 'r', encoding='utf-8') as f: reader = csv.reader(f); linhas = list(reader)
        except: return False

    with open(caminho_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        for row in linhas:
            while len(row) < 10: row.append("")
            
            # Se encontrar o código, atualiza dados E status
            if row[1] == codigo:
                if 'titulo' in novos_dados: row[6] = novos_dados['titulo']
                if 'projeto' in novos_dados: row[5] = novos_dados['projeto']
                if 'descricao' in novos_dados: row[7] = novos_dados['descricao']
                
                # --- MUDANÇA: Grava status MODIFICADO no arquivo ---
                if row[8] != "INATIVO": # Não reativa se estiver excluído
                    row[8] = "MODIFICADO"
                
                alterado = True
            
            writer.writerow(row)
    return alterado

def gerar_codigo_unico(caminho_csv, prefixo, sufixo):
    existentes = set()
    if os.path.exists(caminho_csv):
        try:
            with open(caminho_csv, 'r', encoding='utf-8') as f:
                reader = csv.reader(f); next(reader, None); [existentes.add(r[1]) for r in reader]
        except: pass
    for _ in range(1000):
        cod = f"{prefixo}-{random.randint(10000, 99999)}{sufixo}"
        if cod not in existentes: return cod
    return None

def excluir_logico(caminho_csv, codigo, caminho_arquivo):
    linhas = []
    if os.path.exists(caminho_csv):
        try:
            with open(caminho_csv, 'r', encoding='utf-8') as f: reader = csv.reader(f); linhas = list(reader)
        except: return
    with open(caminho_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        for row in linhas:
            while len(row) < 10: row.append("")
            if len(row) > 9 and row[1] == codigo and row[9] == caminho_arquivo: row[8] = "INATIVO"
            writer.writerow(row)

def sincronizar_arquivos(caminho_local_csv, pasta_rede):
    # (Mantido igual à versão anterior funcional)
    relatorio = []
    if not os.path.exists(caminho_local_csv): return "Sem dados locais."
    try:
        with open(caminho_local_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f); next(reader, None); dados_locais = list(reader)
    except: return "Erro ao ler CSV local."

    caminho_rede_csv = os.path.join(pasta_rede, "registro_pecas.csv")
    garantir_csv(caminho_rede_csv)

    codigos_rede = set()
    try:
        with open(caminho_rede_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f); next(reader, None)
            for r in reader: codigos_rede.add(r[1])
    except: pass

    novas_linhas = []
    count = 0

    for row in dados_locais:
        while len(row) < 10: row.append("")
        cod = row[1]; path_local = row[9]; status = row[8]
        
        # Sincroniza se for ATIVO ou MODIFICADO
        if path_local and os.path.exists(path_local) and (status == "ATIVO" or status == "MODIFICADO"):
            if cod not in codigos_rede:
                nome_arq = os.path.basename(path_local)
                subpasta = ""
                if "desenhos" in path_local:
                    if "\\ED" in path_local or "/ED" in path_local: subpasta = os.path.join("desenhos", "ED")
                elif os.path.basename(os.path.dirname(path_local)).lower() == "3d":
                    subpasta = "3d"
                
                dest_pasta = os.path.join(pasta_rede, subpasta) if subpasta else pasta_rede
                os.makedirs(dest_pasta, exist_ok=True)
                dest_final = os.path.join(dest_pasta, nome_arq)
                
                try:
                    shutil.copy2(path_local, dest_final)
                    pasta_img_local = os.path.join(os.path.dirname(path_local), "_IMAGENS")
                    if subpasta: pasta_img_local = os.path.join(os.path.dirname(os.path.dirname(path_local)), "_IMAGENS")
                    img_local = os.path.join(pasta_img_local, f"{cod}.jpg")
                    if os.path.exists(img_local):
                        pasta_img_rede = os.path.join(pasta_rede, "_IMAGENS")
                        os.makedirs(pasta_img_rede, exist_ok=True)
                        shutil.copy2(img_local, os.path.join(pasta_img_rede, f"{cod}.jpg"))

                    row[9] = dest_final
                    novas_linhas.append(row)
                    codigos_rede.add(cod)
                    count += 1
                except Exception as e:
                    relatorio.append(f"Erro ao copiar {cod}: {e}")

    if novas_linhas:
        with open(caminho_rede_csv, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            for r in novas_linhas: writer.writerow(r)
        shutil.copy2(caminho_local_csv, caminho_local_csv + ".bak")
        with open(caminho_local_csv, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerow(["Data", "Codigo", "Prefixo", "Numero", "Tipo", "Projeto", "Titulo", "Descricao", "Status", "Caminho"])

    return f"Sincronização: {count} arquivos enviados.\n" + "\n".join(relatorio)