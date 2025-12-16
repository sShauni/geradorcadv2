import csv
import os
import shutil
import random
import config

# --- FUNÇÕES DE CAMINHO ---
def resolver_caminho(caminho_csv, caminho_registrado, raiz_personalizada=None):
    """
    Inteligência para encontrar o arquivo mesmo que a letra do drive mude.
    """
    if not caminho_registrado: return ""
    
    # Normaliza barras (Windows/Linux)
    path_norm = caminho_registrado.replace("/", os.sep).replace("\\", os.sep)
    
    # 1. Tenta encontrar usando a Pasta Raiz Personalizada (REGRA DO PENDRIVE)
    if raiz_personalizada and os.path.exists(raiz_personalizada):
        partes = path_norm.lower().split(os.sep)
        subcaminho = ""
        
        # Tenta reconstruir o caminho relativo baseado em pastas conhecidas
        if "3d" in partes:
            idx = partes.index("3d")
            subcaminho = os.path.join(*path_norm.split(os.sep)[idx:])
        elif "desenhos" in partes:
            idx = partes.index("desenhos")
            subcaminho = os.path.join(*path_norm.split(os.sep)[idx:])
            
        if subcaminho:
            tentativa_smart = os.path.join(raiz_personalizada, subcaminho)
            if os.path.exists(tentativa_smart): return os.path.abspath(tentativa_smart)
        
        # Força bruta: tenta achar o arquivo direto nas pastas padrões
        nome_arq = os.path.basename(path_norm)
        paths_teste = [
            os.path.join(raiz_personalizada, "3d", nome_arq),
            os.path.join(raiz_personalizada, "desenhos", "ED", nome_arq),
            os.path.join(raiz_personalizada, nome_arq)
        ]
        for p in paths_teste:
            if os.path.exists(p): return os.path.abspath(p)

    # 2. Tenta relativo ao CSV (Padrão)
    pasta_csv = os.path.dirname(os.path.abspath(caminho_csv))
    path_relativo = os.path.join(pasta_csv, path_norm)
    if os.path.exists(path_relativo): return os.path.abspath(path_relativo)

    # 3. Caminho absoluto original (Legado)
    if os.path.exists(path_norm): return os.path.abspath(path_norm)
        
    return os.path.abspath(path_relativo)

def tornar_relativo(caminho_csv, caminho_arquivo):
    try:
        pasta_base = os.path.dirname(os.path.abspath(caminho_csv))
        return os.path.relpath(caminho_arquivo, pasta_base)
    except:
        return caminho_arquivo

def garantir_csv(caminho):
    if not os.path.exists(caminho):
        try:
            pasta = os.path.dirname(caminho)
            if not os.path.exists(pasta): return
            with open(caminho, 'w', newline='', encoding='utf-8') as f:
                csv.writer(f).writerow(["Data", "Codigo", "Prefixo", "Numero", "Tipo", "Projeto", "Titulo", "Descricao", "Status", "Caminho"])
        except: pass

# --- FUNÇÕES DE REGISTRO ---
def ler_registros(caminho_csv, mostrar_inativos=False, termo_busca="", ocultar_desenhos=False, raiz_personalizada=None):
    registros = []
    
    # Se existir um CSV na raiz personalizada (Pendrive), usa ele!
    if raiz_personalizada:
        csv_pendrive = os.path.join(raiz_personalizada, "registro_pecas.csv")
        if os.path.exists(csv_pendrive):
            caminho_csv = csv_pendrive

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
        # --- PREENCHIMENTO DE SEGURANÇA (EVITA O ERRO INDEX ERROR) ---
        while len(row) < 10: 
            row.append("")
        # -------------------------------------------------------------
        
        status = row[8]
        tipo = row[4]
        
        # Resolve o caminho real do arquivo
        caminho_real = resolver_caminho(caminho_csv, row[9], raiz_personalizada)
        row[9] = caminho_real 

        if status == "INATIVO" and not mostrar_inativos: continue
        if ocultar_desenhos and "DESENHO" in tipo.upper(): continue
        
        txt_busca = (row[1] + row[6] + row[7]).lower()
        if termo in txt_busca:
            resultados.append(row)
            
    # Retorna (Lista de Dados, Caminho do CSV usado)
    return resultados, caminho_csv 

def gravar_linha(caminho_csv, dados_linha):
    caminho_abs = dados_linha[9]
    dados_linha[9] = tornar_relativo(caminho_csv, caminho_abs)
    with open(caminho_csv, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f); writer.writerow(dados_linha)

def editar_registro(caminho_csv, codigo, novos_dados):
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
            
            if row[1] == codigo:
                if 'titulo' in novos_dados: row[6] = novos_dados['titulo']
                if 'projeto' in novos_dados: row[5] = novos_dados['projeto']
                if 'descricao' in novos_dados: row[7] = novos_dados['descricao']
                if row[8] != "INATIVO": row[8] = "MODIFICADO"
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
            if row[1] == codigo: row[8] = "INATIVO"
            writer.writerow(row)

def sincronizar_arquivos(caminho_local_csv, pasta_rede):
    relatorio = []
    if not os.path.exists(caminho_local_csv): return "Sem dados locais."
    
    # Usa ler_registros para pegar caminhos já resolvidos (mas ignora o retorno do caminho_csv)
    dados_locais, _ = ler_registros(caminho_local_csv, mostrar_inativos=True) 

    caminho_rede_csv = os.path.join(pasta_rede, "registro_pecas.csv")
    garantir_csv(caminho_rede_csv)

    codigos_rede = set()
    try:
        with open(caminho_rede_csv, 'r', encoding='utf-8') as f:
            reader = csv.reader(f); next(reader, None); [codigos_rede.add(r[1]) for r in reader]
    except: pass

    novas_linhas = []
    count = 0

    for row in dados_locais:
        path_local_abs = row[9]; cod = row[1]; status = row[8]
        
        if path_local_abs and os.path.exists(path_local_abs) and (status == "ATIVO" or status == "MODIFICADO"):
            if cod not in codigos_rede:
                nome_arq = os.path.basename(path_local_abs)
                subpasta = ""
                if "desenhos" in path_local_abs.lower() and "ED" in path_local_abs: 
                    subpasta = os.path.join("desenhos", "ED")
                elif os.path.basename(os.path.dirname(path_local_abs)).lower() == "3d":
                    subpasta = "3d"
                
                dest_pasta = os.path.join(pasta_rede, subpasta) if subpasta else pasta_rede
                os.makedirs(dest_pasta, exist_ok=True)
                dest_final = os.path.join(dest_pasta, nome_arq)
                
                try:
                    shutil.copy2(path_local_abs, dest_final)
                    pasta_img = os.path.join(os.path.dirname(path_local_abs), "_IMAGENS")
                    if subpasta: pasta_img = os.path.join(os.path.dirname(os.path.dirname(path_local_abs)), "_IMAGENS")
                    img = os.path.join(pasta_img, f"{cod}.jpg")
                    if os.path.exists(img):
                        pd = os.path.join(pasta_rede, "_IMAGENS"); os.makedirs(pd, exist_ok=True)
                        shutil.copy2(img, os.path.join(pd, f"{cod}.jpg"))

                    row[9] = os.path.join(subpasta, nome_arq)
                    novas_linhas.append(row)
                    codigos_rede.add(cod)
                    count += 1
                except Exception as e:
                    relatorio.append(f"Erro ao copiar {cod}: {e}")

    if novas_linhas:
        with open(caminho_rede_csv, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f); [writer.writerow(r) for r in novas_linhas]
    return f"Sincronização: {count} arquivos enviados.\n" + "\n".join(relatorio)