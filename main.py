import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from interface import MainWindow # Agora importamos MainWindow, não AppGUI
import config

if __name__ == "__main__":
    # Cria a instância da aplicação Qt
    app = QApplication(sys.argv)
    
    # Define informações básicas
    app.setApplicationName("Invenio")
    app.setApplicationVersion("2.0")

    # Define ícone da janela e da taskbar
    icone_path = config.obter_caminho_recurso("icon.ico")
    if os.path.exists(icone_path):
        app.setWindowIcon(QIcon(icone_path))

    # Inicia a Interface Principal
    janela = MainWindow()
    janela.show()

    # Loop de execução
    sys.exit(app.exec())