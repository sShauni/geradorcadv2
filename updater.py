import os
import sys
import subprocess
import requests

# ==============================================================================
# CONFIGURAÇÕES
# ==============================================================================
REPO_OWNER = "sShauni"
REPO_NAME = "geradorcadv2"
# Note que aqui o nome deve ser URL_CHECK para combinar com a classe abaixo
URL_CHECK = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
# ==============================================================================

class Updater:
    def __init__(self, versao_atual):
        self.versao_atual = versao_atual
        self.dados_nova_versao = None
        self.temp_exe_name = "update_temp.exe"

    def verificar_atualizacao(self):
        """
        Retorna (True, tag_name, download_url) se houver update.
        Retorna (False, None, None) se já estiver atualizado ou der erro.
        """
        try:
            # Timeout curto para não travar se a internet estiver ruim
            response = requests.get(URL_CHECK, timeout=5)
            response.raise_for_status()
            data = response.json()
            
            tag_remota = data.get("tag_name", "v0.0").replace("v", "")
            versao_limpa = self.versao_atual.replace("v", "")

            # Comparação simples de string
            if tag_remota > versao_limpa:
                url_download = None
                # Procura o asset que termina com .exe
                for asset in data.get("assets", []):
                    if asset["name"].endswith(".exe"):
                        url_download = asset["browser_download_url"]
                        break
                
                if url_download:
                    self.dados_nova_versao = (tag_remota, url_download)
                    return True, tag_remota, url_download
            
            return False, None, None

        except Exception as e:
            print(f"Erro ao buscar update: {e}")
            return False, None, None

    def baixar_atualizacao(self, url, callback_progresso=None):
        """
        Baixa o arquivo com barra de progresso.
        """
        try:
            response = requests.get(url, stream=True)
            total_length = response.headers.get('content-length')

            with open(self.temp_exe_name, "wb") as f:
                if total_length is None: 
                    f.write(response.content)
                else:
                    dl = 0
                    total_length = int(total_length)
                    for data in response.iter_content(chunk_size=4096):
                        dl += len(data)
                        f.write(data)
                        if callback_progresso:
                            percent = int((dl / total_length) * 100)
                            callback_progresso(percent)
            return True
        except Exception as e:
            print(f"Erro no download: {e}")
            return False

    def aplicar_atualizacao(self):
        """
        Cria script BAT para substituir o executável e reiniciar.
        """
        app_exe = sys.executable
        app_dir = os.path.dirname(os.path.abspath(app_exe))
        
        # Caminho completo do novo arquivo baixado
        novo_arquivo = os.path.abspath(self.temp_exe_name)

        # Script Batch para fazer a troca
        bat_script = f"""
@echo off
timeout /t 2 /nobreak > NUL
move /y "{novo_arquivo}" "{app_exe}"
start "" "{app_exe}"
del "%~f0"
"""
        
        bat_path = os.path.join(app_dir, "update_installer.bat")
        
        with open(bat_path, "w") as bat_file:
            bat_file.write(bat_script)
        
        # Executa o BAT e fecha o Python imediatamente
        subprocess.Popen(bat_path, shell=True)
        sys.exit(0)