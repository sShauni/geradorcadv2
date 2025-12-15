import sys
import os
import ctypes
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from interface import MainWindow 
import config

# Função para garantir que achamos o ícone (seja no código ou no .exe)
def resource_path(relative_path):
    try:
        # PyInstaller cria uma pasta temporária em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Se não estiver compilado, usa o caminho do próprio main.py
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    myappid = 'empresa.invenio.cad.v2' 
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass # Ignora se não for Windows ou der erro

    app = QApplication(sys.argv)
    
    app.setApplicationName("Invenio")
    app.setApplicationVersion("2.2")

    # 2. CARREGAMENTO ROBUSTO DO ÍCONE
    # Usamos a função resource_path para garantir o caminho absoluto
    icone_path = resource_path("icon.ico")
    
    # Debug: Print para você ver se ele achou o arquivo no console
    print(f"Tentando carregar ícone de: {icone_path}")
    
    if os.path.exists(icone_path):
        app_icon = QIcon(icone_path)
        app.setWindowIcon(app_icon)
    else:
        print("ERRO: Arquivo icon.ico não encontrado no caminho acima!")

    janela = MainWindow()
    
    # Força o ícone na janela principal também
    if os.path.exists(icone_path):
        janela.setWindowIcon(QIcon(icone_path))

    janela.show()

    sys.exit(app.exec())