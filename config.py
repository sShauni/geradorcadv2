# config.py
import os
import json
import sys

# --- CORREÇÃO DO CAMINHO (CRUCIAL PARA .EXE) ---
if getattr(sys, 'frozen', False):
    # Se for executável (.exe), a pasta é onde o arquivo está
    PASTA_SCRIPT = os.path.dirname(sys.executable)
else:
    # Se for script (.py), a pasta é onde o código está
    PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))

ARQUIVO_CONFIG = os.path.join(PASTA_SCRIPT, "config_usuario.json")
ARQUIVO_CSV_LOCAL = os.path.join(PASTA_SCRIPT, "registro_pecas.csv")

# Opções
TIPOS_FABRICACAO = {
    "Usinagem (Torno/Fresa)": "-USI",
    "Corte Laser/Plasma": "-CRT",
    "Dobra/Estamparia": "-DBR",
    "Soldagem": "-SLD",
    "Impressão 3D": "-3DP",
    "Comercial (Comprado)": "-OEM",
    "Montagem": "-ASM"
}

def carregar():
    try:
        with open(ARQUIVO_CONFIG, 'r') as f:
            return json.load(f)
    except:
        return {}

def salvar(dados_dict):
    try:
        # Garante que as chaves existam antes de salvar
        with open(ARQUIVO_CONFIG, 'w') as f:
            json.dump(dados_dict, f, indent=4) # indent deixa o arquivo legível
    except Exception as e:
        print(f"Erro ao salvar config: {e}")
        
def obter_caminho_recurso(nome_arquivo):
    """Retorna o caminho correto para imagens/ícones, seja no .py ou no .exe"""
    if hasattr(sys, '_MEIPASS'):
        # Se for .exe, o PyInstaller descompacta os arquivos nesta pasta temporária
        return os.path.join(sys._MEIPASS, nome_arquivo)
    else:
        # Se for .py, procura na pasta do script mesmo
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), nome_arquivo)