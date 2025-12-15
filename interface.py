import sys
import os
import subprocess
import ctypes 
from datetime import datetime

# --- Importações Qt (PySide6) ---
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                               QTableWidget, QTableWidgetItem, QHeaderView, QFrame,
                               QComboBox, QCheckBox, QMessageBox, QFileDialog, 
<<<<<<< HEAD
                               QMenu, QGroupBox, QDialog, QFormLayout)
from PySide6.QtCore import Qt, QTimer, QRect, QSize, Signal, QThread
=======
                               QMenu, QGroupBox, QDialog, QFormLayout, QProgressDialog)
from PySide6.QtCore import Qt, QTimer, QRect, QSize, QThread, Signal
>>>>>>> b9b3ae538bf17b7da839395c142a2e674acef179
from PySide6.QtGui import QColor, QFont, QPixmap, QImage, QIcon

# --- Importações do Projeto ---
from PIL import Image
import config
import dados
import inventor
import scripts_vb 
import updater

# ### CORREÇÃO 1: RESOLUÇÃO BORRADA (HIGH DPI) ###
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass 

os.environ["QT_FONT_DPI"] = "96" 

# === ESTILO VISUAL LIMPO & CORRIGIDO ===
QSS_INVENTOR = """
/* === RESET GLOBAL === */
QMainWindow { background-color: #454e61; }
QWidget { 
    color: #E0E0E0; 
    font-family: 'Calibri'; 
    font-size: 12px; 
    border: none; 
}

/* === FRAMES E GRUPOS === */
QFrame { border: none; background: transparent; }
QFrame#FrameEsquerdo {
    background-color: #303642;
    border-right: 1px solid #3B4252;
}

QDialog {
    background-color: #2F3642;
    border: 1px solid #454E61;
}

QLabel {
    color: #A0A8B8;
    font-size: 11px;
    font-weight: bold;
    font-family: 'Segoe UI';
}

/* Títulos das Seções */
QLabel#TituloSecao {
    color: #6D7685;
    font-weight: bold;
    padding-bottom: 5px;
    border-bottom: 1px solid #3E4654;
}

QLineEdit {
    background-color: #262B35;
    border: 1px solid #3E4654;
    color: white;
    padding: 6px;
    border-radius: 3px;
    font-weight: bold;
}
QLineEdit:focus { border: 1px solid #88C0D0; }

/* === CHECKBOX === */
QCheckBox {
    color: #E0E0E0;
    font-size: 13px;
    spacing: 8px;
}

QCheckBox::indicator {
    width: 18px;
    height: 18px;
    background-color: transparent;
    border: 1px solid #748596;
    border-radius: 1px;
}

QCheckBox::indicator:hover {
    border: 1px solid #FFFFFF;
}

QCheckBox::indicator:checked {
    background-color: white;
    border: 1px solid white;
    image: url("checker.png");
}

QGroupBox { 
    border: none; 
    margin-top: 20px; 
    font-weight: bold;
}
QGroupBox::title { 
    subcontrol-origin: margin; left: 0px; padding: 0 3px; color: #88C0D0; 
}

/* === INPUTS E COMBOS === */
QLineEdit, QComboBox {
    background-color: #3B4252; 
    border: 1px solid #4C566A; 
    border-radius: 2px;
    padding: 4px; 
    color: white; 
}
QLineEdit:focus, QComboBox:focus { 
    border: 1px solid #00AFFF;
    border-radius: 3px;
}
QLineEdit:read-only { 
    background-color: #303642; color: #88C0D0; font-weight: bold; border: none; border-bottom: 2px solid #748596;
}

/* --- COMBOBOX --- */
QComboBox::drop-down { border: none; width: 20px; }
QComboBox QAbstractItemView {
    background-color: #454F61;
    color: white;
    border: 1px solid #4C566A;
    selection-background-color: #454F61;
    outline: none;
}

/* === TABELA === */
QTableWidget {
    background-color: #475163; 
    border: none; 
    outline: none;
}
QTableWidget::item {
    border-bottom: 1px solid #525E70; 
    padding-left: 5px;
}
QHeaderView::section {
    background-color: #333a49; color: white; padding: 6px;
    border: none; border-bottom: 2px solid #333B49; font-weight: bold;
}
QTableWidget::item:selected {
    background-color: rgba(67 , 92, 116, 100); 
    border: 1px solid #3D84AA; 
    color: white;
}
QTableWidget::item:hover { background-color: #323846; }

/* === BOTÕES === */
QPushButton {
    background-color: transparent; 
    border: none;
    padding: 4px 15px; border-radius: 2px; color: white;
    min-height: 36px; max-height: 40px;
    font-family: 'Calibri';
    font-weight: bold;
<<<<<<< HEAD
    font-size: 14px;
=======
    font-size: 8pt;
>>>>>>> b9b3ae538bf17b7da839395c142a2e674acef179
}
QPushButton:hover { background-color: #3a4252; border: none; }
QPushButton:pressed { background-color: #435C74; border: 1px solid #3D84AA; border-radius: 2px; }
    
QPushButton.BtnAcao { 
    text-align: left; 
    padding-left: 15px;
}

QPushButton#BtnConfig {
    background-color: transparent;
    border: none;
    border-radius: 4px;
    min-width: 32px; max-width: 32px;
}
QPushButton#BtnConfig:hover {
    background-color: #1e2229;
    border: none;
}

QPushButton#BtnGerar {
    background-color: transparent; border: 1px solid #748596; border-style: solid;
    min-height: 40px;
    max-height: 45px;
    font-weight: bold;
    font-size: 15px;
}
QPushButton#BtnGerar:hover { background-color: #2E3440; border: 1px solid white; }
QPushButton#BtnGerar:pressed { background-color: transparent; border: 1px solid #515d6c; border-radius: 2px; }

QPushButton#BtnSalvarEsq {
    background-color: #2E3440; border: 1px solid #748596; border: 1px solid #748596; border-style: solid;
    color: #FFFFFF;
    min-height: 40px;
    max-height: 45px;
    font-weight: bold;
    font-size: 14px;
}
QPushButton#BtnSalvarEsq:hover { background-color: transparent; border: 1px solid white; }
QPushButton#BtnSalvarEsq:pressed { background-color: transparent; border: 1px solid #515d6c; border-radius: 2px; }

QPushButton#BtnSync {
    background-color: transparent; border: 1px solid #748596; border-style: solid;
    font-size: 14px; color: white; border: 1px solid #4C566A;
}
QPushButton#BtnSync:hover { background-color: transparent; border: 1px solid white; }
QPushButton#BtnSync:pressed { background-color: transparent; border: 1px solid #515d6c; border-radius: 2px; }

QScrollBar:vertical { background: #454F61; width: 6px; }
QScrollBar::handle:vertical { background: #5D697E; min-height: 15px; border-radius: 3px; }
"""

<<<<<<< HEAD
class WorkerConexao(QThread):
    # Sinais para comunicar com a interface principal
    finalizado = Signal(bool, str, str) # (Sucesso, Mensagem, CaminhoConfirmado)

    def __init__(self, cfg):
        super().__init__()
        self.cfg = cfg

    def run(self):
        ip = self.cfg.get("ip")
        path = self.cfg.get("path")
        user = self.cfg.get("user")
        senha = self.cfg.get("pass")
        
        unc = f"\\\\{ip}\\{path}"
        
        # 1. Tenta verificar se já existe (rápido)
        if os.path.exists(unc):
            self.finalizado.emit(True, f"SRV: {ip} (OK)", unc)
            return

        # 2. Se não existe, tenta conectar com timeout de 5 segundos
        cmd = f'net use "{unc}" /user:{user} "{senha}"'
        try:
            # O timeout=5 garante que não fique travado para sempre
            subprocess.run(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=5)
            
            # 3. Verifica novamente se conectou
            if os.path.exists(unc):
                self.finalizado.emit(True, f"SRV: {ip} (OK)", unc)
            else:
                self.finalizado.emit(False, "LOCAL (Falha Rede)", "")
                
        except subprocess.TimeoutExpired:
            # Se passar de 5 segundos, aborta
            self.finalizado.emit(False, "LOCAL (Timeout)", "")
        except Exception as e:
            self.finalizado.emit(False, "LOCAL (Erro)", "")
=======
class CheckUpdateWorker(QThread):
    update_encontrado = Signal(str, str) # tag, url

    def __init__(self, versao_atual):
        super().__init__()
        self.upd = updater.Updater(versao_atual)

    def run(self):
        tem_update, tag, url = self.upd.verificar_atualizacao()
        if tem_update:
            self.update_encontrado.emit(tag, url)
>>>>>>> b9b3ae538bf17b7da839395c142a2e674acef179

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Invenio 2.1")
        
        # Carrega config
        self.cfg = config.carregar()
        
        # Define tamanho padrão inicial (Backup)
        self.resize(1280, 720) 

        self.setStyleSheet(QSS_INVENTOR)
        
        self.caminho_db_atual = config.ARQUIVO_CSV_LOCAL
        self.caminho_rede_ativo = None
        
        path_base = os.path.dirname(os.path.abspath(__file__))
        self.icon_ipt = QIcon(os.path.join(path_base, "ipt.ico"))
        self.icon_iam = QIcon(os.path.join(path_base, "iam.ico"))
        self.icon_idw = QIcon(os.path.join(path_base, "idw.ico"))
        
        self.ico_config = QIcon(os.path.join(path_base, "conf.png"))
        self.ico_abrir  = QIcon(os.path.join(path_base, "open.png"))
        self.ico_pasta  = QIcon(os.path.join(path_base, "folder.png"))
        self.ico_ins    = QIcon(os.path.join(path_base, "insert.png"))
        self.ico_edit   = QIcon(os.path.join(path_base, "edit.png"))
        self.ico_del    = QIcon(os.path.join(path_base, "trash.png"))
        
        self.ico_gerar  = QIcon(os.path.join(path_base, "gen.png"))
        self.ico_save   = QIcon(os.path.join(path_base, "save.png"))
        self.ico_sync   = QIcon(os.path.join(path_base, "sync.png"))
        self.ico_laser  = QIcon(os.path.join(path_base, "dxf.png"))
        self.ico_fix    = QIcon(os.path.join(path_base, "fix.png"))
        
        caminho_checker = os.path.join(path_base, "checker.png")
        caminho_checker = caminho_checker.replace("\\", "/")
        qss_final = QSS_INVENTOR.replace("checker.png", caminho_checker)
        self.setStyleSheet(qss_final)
        
        dados.garantir_csv(self.caminho_db_atual)
        self.setup_ui()
        
        try:
            hwnd = int(self.winId())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20 
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 
                DWMWA_USE_IMMERSIVE_DARK_MODE, 
                ctypes.byref(ctypes.c_int(2)), 
                4
            )
        except Exception as e:
            print(f"Theme Error: {e}")
        
        if self.cfg.get("usar_servidor"):
            QTimer.singleShot(100, self.conectar_rede)
        else:
            self.atualizar_lista()
    
        # === SISTEMA DE UPDATE ===
        self.VERSAO_ATUAL = "2.1" 
        self.worker_update = CheckUpdateWorker(self.VERSAO_ATUAL)
        self.worker_update.update_encontrado.connect(self.mostrar_aviso_update)
        self.worker_update.start()

        # === CORREÇÃO DE JANELA ===
        # Executa a restauração da geometria logo após a janela ser criada
        QTimer.singleShot(10, self.restaurar_geometria)

    def restaurar_geometria(self):
        """Aplica a posição de forma segura, evitando bugs visuais."""
        default_rect = [100, 100, 1280, 720]
        rect_salvo = self.cfg.get("janela_geo", default_rect)
        is_maximized = self.cfg.get("janela_max", False)

        # Validação: Se o tamanho salvo for muito pequeno (bug), usa padrão
        if len(rect_salvo) == 4:
            x, y, w, h = rect_salvo
            if w < 200 or h < 200: 
                rect_salvo = default_rect
            
            # Aplica a geometria (posição e tamanho "normal")
            self.setGeometry(*rect_salvo)

        # Se tiver que maximizar, faz isso POR ÚLTIMO
        if is_maximized:
            self.showMaximized()

    def closeEvent(self, event):
        """Salva o estado ao fechar."""
        try:
            # 1. Salva se está maximizado
            esta_maximizada = self.isMaximized()
            self.cfg["janela_max"] = esta_maximizada
            
            # 2. Salva a geometria
            if esta_maximizada:
                # Se está maximizada, salvamos o tamanho "restaurado/normal"
                # para que quando o usuário desmaximizar, ela tenha um tamanho decente.
                geo = self.normalGeometry()
            else:
                # Se não está, salva o tamanho atual
                geo = self.geometry()
            
            # Garante que não salvamos valores negativos loucos (comuns em multi-monitor)
            x = max(0, geo.x())
            y = max(0, geo.y())
            self.cfg["janela_geo"] = [x, y, geo.width(), geo.height()]
            
            config.salvar(self.cfg)
        except Exception as e:
            print(f"Erro ao salvar config de janela: {e}")
        
        event.accept()

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0); main_layout.setSpacing(0)

        # ==========================================
        # 1. PAINEL ESQUERDO
        # ==========================================
        frame_esq = QFrame()
        frame_esq.setObjectName("FrameEsquerdo")
        frame_esq.setFixedWidth(400)
        layout_esq = QVBoxLayout(frame_esq)
        layout_esq.setSpacing(10)
        layout_esq.setContentsMargins(15, 15, 15, 15)
        
        # Status Rede
        self.lbl_rede = QLabel("MODO: LOCAL")
        self.lbl_rede.setStyleSheet("color: #DDD; font-weight: bold; border: none;")
        layout_esq.addWidget(self.lbl_rede)

        self.btn_sync = QPushButton(" Sincronizar")
        self.btn_sync.setObjectName("BtnSync")
        self.btn_sync.setIcon(self.ico_sync)
        self.btn_sync.setIconSize(QSize(16, 16))
        self.btn_sync.setEnabled(False)
        self.btn_sync.clicked.connect(self.acao_sincronizar)
        layout_esq.addWidget(self.btn_sync)
        
        # Grupo: Dados do Projeto
        gb_proj = QGroupBox("DADOS DO PROJETO")
        ly_proj = QFormLayout(gb_proj)
        ly_proj.setSpacing(8)
        self.in_projeto = QLineEdit(self.cfg.get("ultimo_projeto", ""))
        self.in_titulo = QLineEdit()
        self.in_desc = QLineEdit()
        ly_proj.addRow("Projeto:", self.in_projeto)
        ly_proj.addRow("Título:", self.in_titulo)
        ly_proj.addRow("Descrição:", self.in_desc)
        layout_esq.addWidget(gb_proj)

        # Grupo: Gerador
        gb_gen = QGroupBox("GERADOR DE CÓDIGOS")
        ly_gen = QVBoxLayout(gb_gen)
        ly_gen.setSpacing(10)
        row_p = QHBoxLayout()
        self.in_prefixo = QLineEdit(self.cfg.get("ultimo_prefixo", "PRJ")); self.in_prefixo.setFixedWidth(80)
        self.cb_tipo = QComboBox(); self.cb_tipo.addItems(list(config.TIPOS_FABRICACAO.keys()))
        row_p.addWidget(QLabel("Prefixo:")); row_p.addWidget(self.in_prefixo); row_p.addWidget(self.cb_tipo)
        ly_gen.addLayout(row_p)
        
        self.in_codigo_gerado = QLineEdit(); self.in_codigo_gerado.setReadOnly(True)
        self.in_codigo_gerado.setPlaceholderText("Código aparecerá aqui...")
        self.in_codigo_gerado.setAlignment(Qt.AlignCenter)
        self.in_codigo_gerado.setStyleSheet("font-size: 16px; padding: 8px; color: #88C0D0; font-weight: bold;")
        ly_gen.addWidget(self.in_codigo_gerado)
        
        # Botões Maiores
        btn_gerar = QPushButton(" GERAR CÓDIGO"); btn_gerar.setObjectName("BtnGerar")
        btn_gerar.setCursor(Qt.PointingHandCursor)
        btn_gerar.setIcon(self.ico_gerar)
        btn_gerar.setIconSize(QSize(20, 20))
        btn_gerar.clicked.connect(self.acao_gerar_codigo)
        ly_gen.addWidget(btn_gerar)
        
        btn_salvar_esq = QPushButton(" SALVAR"); btn_salvar_esq.setObjectName("BtnSalvarEsq")
        btn_salvar_esq.setCursor(Qt.PointingHandCursor)
        btn_salvar_esq.setIcon(self.ico_save)
        btn_salvar_esq.setIconSize(QSize(20, 20))
        btn_salvar_esq.clicked.connect(lambda: self.acao_salvar(origem="gerado"))
        ly_gen.addWidget(btn_salvar_esq)
        layout_esq.addWidget(gb_gen)
        
        # Imagem Preview
        self.lbl_img = QLabel(""); self.lbl_img.setAlignment(Qt.AlignCenter)
        self.lbl_img.setMinimumHeight(200)
        self.lbl_img.setStyleSheet("background-color: #21252B; border: 1px solid #3B4252; border-radius: 2px; margin-top: 5px;")
        layout_esq.addWidget(self.lbl_img)
        
        layout_esq.addStretch()
        row_cfg = QHBoxLayout()
        btn_cfg = QPushButton()
        btn_cfg.setObjectName("BtnConfig")     # Usa o estilo CSS novo
        btn_cfg.setIcon(self.ico_config)       # Define o ícone
        btn_cfg.setIconSize(QSize(20, 20))     # Tamanho do ícone
        btn_cfg.setToolTip("Configurações do Servidor")
        btn_cfg.setCursor(Qt.PointingHandCursor)
        btn_cfg.clicked.connect(self.janela_servidor)
        
        row_cfg.addWidget(btn_cfg)
        row_cfg.addStretch() # Empurra o botão para a esquerda
        layout_esq.addLayout(row_cfg)

        main_layout.addWidget(frame_esq)

        # ==========================================
        # 2. PAINEL DIREITO
        # ==========================================
        frame_dir = QWidget()
        layout_dir = QVBoxLayout(frame_dir)
        layout_dir.setContentsMargins(15, 15, 15, 15)
        
        # Barra Superior
        top_bar = QHBoxLayout()
        self.in_busca = QLineEdit(); self.in_busca.setPlaceholderText("🔍 Buscar código..."); self.in_busca.setStyleSheet("font-weight: bold;")
        self.in_busca.textChanged.connect(self.atualizar_lista)
        
        # CHECKBOXES (Carregando estado do config)
        self.chk_desenhos = QCheckBox("Ocultar Desenhos")
        self.chk_desenhos.setChecked(self.cfg.get("ocultar_desenhos", False)) # Carrega config
        self.chk_desenhos.toggled.connect(self.ao_alternar_filtro)
        
        self.chk_lixeira = QCheckBox("Lixeira")
        self.chk_lixeira.setChecked(self.cfg.get("mostrar_inativos", False)) # Carrega config
        self.chk_lixeira.toggled.connect(self.ao_alternar_filtro)
        
        top_bar.addWidget(QLabel("<b>ITENS</b>")); top_bar.addStretch()
        top_bar.addWidget(self.in_busca); top_bar.addWidget(self.chk_desenhos); top_bar.addWidget(self.chk_lixeira)
        layout_dir.addLayout(top_bar)
        
        # Tabela + Botões
        area_meio = QHBoxLayout()
        area_meio.setSpacing(10)
        
        self.table = QTableWidget(); self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["CÓDIGO", "TIPO", "TÍTULO", "DESCRIÇÃO", "STATUS"])
        self.table.verticalHeader().setVisible(False); self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers); self.table.setShowGrid(False)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.itemSelectionChanged.connect(self.ao_selecionar)
        self.table.doubleClicked.connect(self.acao_abrir_inventor)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.mostrar_menu_contexto)
        area_meio.addWidget(self.table)
        
        # Coluna Botões Direita
        fr_botoes = QFrame()
        fr_botoes.setFixedWidth(160)
        ly_botoes = QVBoxLayout(fr_botoes)
        ly_botoes.setContentsMargins(5, 0, 0, 0)
        ly_botoes.setAlignment(Qt.AlignTop)
        
        ly_botoes.addWidget(QLabel("<b>ARQUIVO</b>"))
        # Botões de Ação
        b_abrir = QPushButton(" Abrir Inventor"); b_abrir.setProperty("class", "BtnAcao"); b_abrir.setIcon(self.ico_abrir); b_abrir.clicked.connect(self.acao_abrir_inventor)
        b_pasta = QPushButton(" Abrir Pasta"); b_pasta.setProperty("class", "BtnAcao"); b_pasta.setIcon(self.ico_pasta); b_pasta.clicked.connect(self.acao_abrir_local)
        ly_botoes.addWidget(b_abrir); ly_botoes.addWidget(b_pasta)
        
        ly_botoes.addSpacing(15); ly_botoes.addWidget(QLabel("<b>MONTAGEM</b>"))
        b_ins = QPushButton(" Inserir +"); b_ins.setProperty("class", "BtnAcao"); b_ins.setIcon(self.ico_ins); b_ins.clicked.connect(self.acao_inserir_montagem)
        ly_botoes.addWidget(b_ins)
        
        ly_botoes.addSpacing(15); ly_botoes.addWidget(QLabel("<b>DADOS</b>"))
        b_edit = QPushButton(" Editar"); b_edit.setProperty("class", "BtnAcao"); b_edit.setIcon(self.ico_edit); b_edit.clicked.connect(self.acao_editar)
        b_del = QPushButton(" Excluir"); b_del.setProperty("class", "BtnAcao"); b_del.setIcon(self.ico_del); b_del.clicked.connect(self.acao_excluir)
        ly_botoes.addWidget(b_edit); ly_botoes.addWidget(b_del)
        
        area_meio.addWidget(fr_botoes)
        layout_dir.addLayout(area_meio)
        
        # Rodapé
        action_bar = QHBoxLayout()
        action_bar.addWidget(QLabel("<b>Automação: </b>"))
        
        b_laser = QPushButton(" Exportar DXF");
        b_laser.setIcon(self.ico_laser)
        b_laser.setIconSize(QSize(18, 18))
        b_laser.clicked.connect(self.acao_exportar_laser)        
        
        b_fix = QPushButton(" Lista Fixadores");
        b_fix.setIcon(self.ico_fix)
        b_fix.setIconSize(QSize(18, 18))
        b_fix.clicked.connect(self.acao_lista_fixadores)
        
        action_bar.addWidget(b_laser);
        action_bar.addWidget(b_fix);
        action_bar.addStretch()
        
        layout_dir.addLayout(action_bar)
        
        main_layout.addWidget(frame_dir)
        
        self.setup_placeholder()
    
    def focar_inventor(self, app):
        """Traz a janela do Inventor para frente sem bugar o estado."""
        try:
            if not app: return
            hwnd = app.MainFrameHWND
            
            # Verifica se está minimizado (IsIconic)
            if ctypes.windll.user32.IsIconic(hwnd):
                # Se estiver minimizado, restaura (SW_RESTORE = 9)
                ctypes.windll.user32.ShowWindow(hwnd, 9)
            else:
                pass
            
            # Traz para frente e define foco
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            
        except Exception as e:
            print(f"Erro ao focar Inventor: {e}")

    # =================================

    def atualizar_lista(self):
        self.table.setRowCount(0)
        regs = dados.ler_registros(self.caminho_db_atual, self.chk_lixeira.isChecked(), self.in_busca.text(), self.chk_desenhos.isChecked())
        self.table.setRowCount(len(regs))
        for i, row in enumerate(regs):
            # row[9] é o caminho completo do arquivo
            caminho_arquivo = row[9].lower() if len(row) > 9 else ""
            icone_atual = self.icon_ipt # Padrão (Peça)
            
            if caminho_arquivo.endswith(".idw"):
                icone_atual = self.icon_idw
            elif caminho_arquivo.endswith(".iam"):
                icone_atual = self.icon_iam

            # Cria o item da primeira coluna (CÓDIGO) e define o ícone
            item_cod = QTableWidgetItem(row[1])
            item_cod.setIcon(icone_atual) # <--- Aplica o ícone aqui
            item_cod.setData(Qt.UserRole, row[9])
            
            c = QColor("white")
            if row[8] == "INATIVO": c = QColor("gray")
            elif row[8] == "MODIFICADO": c = QColor("#88C0D0"); item_cod.setFont(QFont("Segoe UI", 9, QFont.Bold))
            elif "DESENHO" in row[4].upper(): c = QColor("#81A1C1")
            
            self.table.setItem(i, 0, item_cod)
            self.table.setItem(i, 1, QTableWidgetItem(row[4]))
            self.table.setItem(i, 2, QTableWidgetItem(row[6]))
            self.table.setItem(i, 3, QTableWidgetItem(row[7]))
            self.table.setItem(i, 4, QTableWidgetItem(row[8]))
            for k in range(5): self.table.item(i, k).setForeground(c)
            
        if len(regs) == 0:
            self.widget_vazio.show()
            self.widget_vazio.adjustSize() # Garante que o tamanho está correto
            # Força um evento de resize para calcular a posição central
            self.ao_redimensionar_tabela(None) 
        else:
            self.widget_vazio.hide()

    def ao_selecionar(self):
        sel = self.table.selectedItems()
        if not sel: self.lbl_img.clear(); return
        caminho = self.table.item(sel[0].row(), 0).data(Qt.UserRole)
        cod = self.table.item(sel[0].row(), 0).text()
        if caminho and os.path.exists(os.path.dirname(caminho)):
            pasta = os.path.dirname(caminho)
            if os.path.basename(pasta).lower() == "3d": pasta = os.path.dirname(pasta)
            elif "desenhos" in pasta.lower(): pasta = os.path.dirname(os.path.dirname(pasta))
            img_path = os.path.join(pasta, "_IMAGENS", f"{cod}.jpg")
            if os.path.exists(img_path):
                try:
                    pil_img = Image.open(img_path); pil_img.thumbnail((380, 250))
                    if pil_img.mode == "RGB": r, g, b = pil_img.split(); pil_img = Image.merge("RGB", (b, g, r))
                    im2 = pil_img.convert("RGBA"); data = im2.tobytes("raw", "BGRA")
                    qim = QImage(data, im2.size[0], im2.size[1], QImage.Format_ARGB32)
                    self.lbl_img.setPixmap(QPixmap.fromImage(qim))
                except: self.lbl_img.setText("Erro Imagem")
            else: self.lbl_img.setText("Sem Imagem")
        else: self.lbl_img.setText("Caminho não encontrado")

    def get_selecionado(self):
        sel = self.table.selectedItems()
        if not sel: return None
        row = sel[0].row()
        return self.table.item(row, 0).text(), self.table.item(row, 0).data(Qt.UserRole)

    def acao_gerar_codigo(self):
        p = self.in_prefixo.text().upper(); s = config.TIPOS_FABRICACAO[self.cb_tipo.currentText()]
        if not p: return
        cod = dados.gerar_codigo_unico(self.caminho_db_atual, p, s)
        if cod: self.in_codigo_gerado.setText(cod); QApplication.clipboard().setText(cod)

    def acao_salvar(self, origem="gerado"):
        destino = self.caminho_rede_ativo if self.caminho_rede_ativo else QFileDialog.getExistingDirectory(self, "Salvar em")
        if not destino: return
        app = inventor.obter_app()
        if not app: return QMessageBox.critical(self, "Erro", "Inventor fechado.")
        doc = app.ActiveDocument
        if not doc: return QMessageBox.warning(self, "Aviso", "Nenhum documento aberto.")

        cod_final = ""
        if origem == "gerado":
            cod_final = self.in_codigo_gerado.text()
            if not cod_final: return QMessageBox.warning(self, "Aviso", "Gere um código!")
        else:
            sel = self.get_selecionado()
            if not sel: return QMessageBox.warning(self, "Aviso", "Selecione um item.")
            cod_final = sel[0]
            if QMessageBox.question(self, "Confirmar", f"Salvar como {cod_final}?") != QMessageBox.Yes: return

        try:
            if doc.DocumentType == 12292: # IDW
                nome_base = "Desenho"
                try: 
                    if doc.ReferencedDocuments.Count > 0: 
                        nome_base = os.path.splitext(os.path.basename(doc.ReferencedDocuments.Item(1).FullFileName))[0]
                except: pass
                caminho_final = os.path.join(destino, "desenhos", "ED", f"{nome_base}-DT.idw")
                os.makedirs(os.path.dirname(caminho_final), exist_ok=True)
                inventor.salvar_idw(doc, caminho_final)
                self.registrar_db(nome_base, "DESENHO TÉCNICO", caminho_final, f"[DESENHO] {self.in_titulo.text()}")
                QMessageBox.information(self, "Sucesso", "Desenho salvo!")
            else: # IPT/IAM
                ext = ".iam" if doc.DocumentType == 12291 else ".ipt"
                caminho_final = os.path.join(destino, "3d", cod_final + ext)
                os.makedirs(os.path.dirname(caminho_final), exist_ok=True)
                inventor.salvar_peca(doc, caminho_final, self.in_titulo.text(), self.in_projeto.text(), self.in_desc.text(), cod_final)
                inventor.capturar_print(app, destino, cod_final)
                self.registrar_db(cod_final, self.cb_tipo.currentText(), caminho_final, self.in_titulo.text())
                QMessageBox.information(self, "Sucesso", f"Peça salva: {cod_final}")
        except Exception as e: QMessageBox.critical(self, "Erro", str(e))

    def registrar_db(self, cod, tipo, caminho, titulo):
        linha = [datetime.now().strftime("%Y-%m-%d %H:%M"), cod, self.in_prefixo.text(), "", 
                 tipo, self.in_projeto.text(), titulo, self.in_desc.text(), "ATIVO", caminho]
        dados.gravar_linha(self.caminho_db_atual, linha)
        self.cfg["ultimo_prefixo"] = self.in_prefixo.text(); self.cfg["ultimo_projeto"] = self.in_projeto.text()
        config.salvar(self.cfg); self.atualizar_lista()

    def acao_editar(self):
        sel = self.get_selecionado()
        if not sel: return
        cod, caminho = sel
        row = self.table.currentRow()
        dlg = QDialog(self); dlg.setWindowTitle(f"Editar: {cod}")
        form = QFormLayout(dlg)
        et = QLineEdit(self.table.item(row, 2).text()); ep = QLineEdit(self.in_projeto.text()); ed = QLineEdit(self.table.item(row, 3).text())
        form.addRow("Título:", et); form.addRow("Projeto:", ep); form.addRow("Descrição:", ed)
        btn = QPushButton("Salvar"); btn.clicked.connect(dlg.accept); form.addRow(btn)
        if dlg.exec():
            novos = {'titulo': et.text(), 'projeto': ep.text(), 'descricao': ed.text()}
            app = inventor.obter_app()
            if app and os.path.exists(caminho):
                inventor.atualizar_propriedades(app, caminho, novos)
                base = os.path.splitext(caminho)[0]
                path_dwg = base.replace("3d", "desenhos\\ED") + "-DT.idw"
                if os.path.exists(path_dwg): inventor.atualizar_propriedades(app, path_dwg, novos)
            dados.editar_registro(self.caminho_db_atual, cod, novos); self.atualizar_lista()

    def acao_inserir_montagem(self):
        sel = self.get_selecionado()
        if sel:
            app = inventor.obter_app()
            if app: 
                try: 
                    inventor.inserir_componente_montagem(app, sel[1])
                    self.focar_inventor(app) 
                except Exception as e: QMessageBox.warning(self, "Erro", str(e))

    def acao_abrir_inventor(self):
        sel = self.get_selecionado()
        if sel: 
            app = inventor.obter_app()
            if app and os.path.exists(sel[1]): 
                inventor.abrir_arquivo(app, sel[1])
                self.focar_inventor(app) 

    def acao_abrir_local(self):
        sel = self.get_selecionado()
        if sel: subprocess.run(f'explorer /select,"{sel[1]}"')

    def acao_excluir(self):
        sel = self.get_selecionado()
        if sel and QMessageBox.question(self, "Excluir", f"Inativar {sel[0]}?") == QMessageBox.Yes:
            dados.excluir_logico(self.caminho_db_atual, sel[0], sel[1]); self.atualizar_lista()

    def acao_sincronizar(self):
        if not self.caminho_rede_ativo: return
        if QMessageBox.question(self, "Sync", "Mover arquivos para servidor?") != QMessageBox.Yes: return
        self.btn_sync.setText("..."); self.btn_sync.setEnabled(False); QApplication.processEvents()
        res = dados.sincronizar_arquivos(config.ARQUIVO_CSV_LOCAL, self.caminho_rede_ativo)
        QMessageBox.information(self, "Info", res)
        self.btn_sync.setText(" Sincronizar"); self.btn_sync.setEnabled(True); self.atualizar_lista()

    def acao_exportar_laser(self):
        app = inventor.obter_app()
        if app: inventor.executar_ilogic(app, scripts_vb.SCRIPT_EXPORTAR_LASER)

    def acao_lista_fixadores(self):
        app = inventor.obter_app()
        if app: inventor.executar_ilogic(app, scripts_vb.SCRIPT_LISTA_FIXADORES)

    def mostrar_menu_contexto(self, pos):
        menu = QMenu()
        menu.setStyleSheet("QMenu { background-color: #2E3440; color: white; border: 1px solid #555; } QMenu::item:selected { background-color: #007ACC; }")
        menu.addAction("Abrir", self.acao_abrir_inventor); menu.addAction("Inserir", self.acao_inserir_montagem)
        menu.addAction("Pasta", self.acao_abrir_local); menu.addSeparator()
        menu.addAction("Editar", self.acao_editar); menu.addAction("Excluir", self.acao_excluir)
        menu.exec(self.table.mapToGlobal(pos))

    def ao_alternar_filtro(self):
        # Salva o estado atual no config.json
        self.cfg["ocultar_desenhos"] = self.chk_desenhos.isChecked()
        self.cfg["mostrar_inativos"] = self.chk_lixeira.isChecked() # Salva "Lixeira"
        config.salvar(self.cfg)
        self.atualizar_lista()

    def conectar_rede(self):
        """Inicia o processo de conexão em segundo plano sem travar a tela."""
        # Feedback Visual Imediato (Azul Ciano)
        self.lbl_rede.setText("CONECTANDO...")
        self.lbl_rede.setStyleSheet("color: #00FFFF; font-weight: bold;") # Cyan Brilhante
        
        # Cria a thread de trabalho
        self.thread_rede = WorkerConexao(self.cfg)
        
        # Conecta o sinal da thread ao método que vai lidar com o resultado
        self.thread_rede.finalizado.connect(self.ao_terminar_conexao)
        
        # Inicia a thread (o código roda em paralelo agora)
        self.thread_rede.start()
        
    def ao_terminar_conexao(self, sucesso, msg, caminho_unc):
        """Chamado automaticamente quando a thread termina (após max 5s)."""
        if sucesso:
            # Sucesso (Verde)
            self.caminho_rede_ativo = caminho_unc
            self.caminho_db_atual = os.path.join(caminho_unc, "registro_pecas.csv")
            dados.garantir_csv(self.caminho_db_atual)
            
            self.lbl_rede.setText(msg)
            self.lbl_rede.setStyleSheet("color: #A3BE8C; font-weight: bold;") # Verde Nord
            self.btn_sync.setEnabled(True)
            
            # Tenta configurar Content Center do Inventor
            app = inventor.obter_app()
            if app: inventor.configurar_content_center(app, caminho_unc)
            
        else:
            # Falha (Vermelho)
            self.lbl_rede.setText(msg)
            self.lbl_rede.setStyleSheet("color: #BF616A; font-weight: bold;") # Vermelho Nord
            self.caminho_rede_ativo = None
            # Volta para o DB local
            self.caminho_db_atual = config.ARQUIVO_CSV_LOCAL
            
        # Atualiza a lista independente do resultado (para mostrar local ou rede)
        self.atualizar_lista()

    def janela_servidor(self):
        dlg = QDialog(self); dlg.setWindowTitle("Config"); dlg.resize(300, 200)
        form = QFormLayout(dlg)
        i_ip = QLineEdit(self.cfg.get("ip")); i_path = QLineEdit(self.cfg.get("path"))
        i_user = QLineEdit(self.cfg.get("user")); i_pass = QLineEdit(self.cfg.get("pass")); i_pass.setEchoMode(QLineEdit.Password)
        chk = QCheckBox("Usar Servidor"); chk.setChecked(self.cfg.get("usar_servidor", False))
        form.addRow("IP:", i_ip); form.addRow("Pasta:", i_path); form.addRow("User:", i_user); form.addRow("Pass:", i_pass); form.addRow(chk)
        btn = QPushButton("Salvar"); btn.clicked.connect(dlg.accept); form.addRow(btn)
        if dlg.exec():
            self.cfg.update({"ip": i_ip.text(), "path": i_path.text(), "user": i_user.text(), "pass": i_pass.text(), "usar_servidor": chk.isChecked()})
            config.salvar(self.cfg); QMessageBox.information(self, "Info", "Reinicie.")
<<<<<<< HEAD
            
    def setup_placeholder(self):
        """Cria o widget de aviso quando não há arquivos."""
        self.widget_vazio = QWidget(self.table)
        self.widget_vazio.setObjectName("WidgetVazio")
        self.widget_vazio.setAttribute(Qt.WA_TransparentForMouseEvents) # Deixa clicar através se necessário
        
        # Layout vertical para empilhar Ícone e Texto
        ly_vazio = QVBoxLayout(self.widget_vazio)
        ly_vazio.setAlignment(Qt.AlignCenter)
        ly_vazio.setSpacing(10) # Espaço entre a caixa e o texto
        
        # 1. Ícone da Caixa
        self.lbl_icon_vazio = QLabel()
        self.lbl_icon_vazio.setAlignment(Qt.AlignCenter)
        # Tenta carregar a imagem (salve seu png como 'box_empty.png' na pasta do script)
        path_img = os.path.join(os.path.dirname(os.path.abspath(__file__)), "box_empty.png")
        if os.path.exists(path_img):
            pixmap = QPixmap(path_img)
            # Redimensiona se for muito grande (opcional, ajustando para 64x64 por exemplo)
            pixmap = pixmap.scaled(256, 256, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.lbl_icon_vazio.setPixmap(pixmap)
        else:
            self.lbl_icon_vazio.setText("📦") # Fallback se não achar a imagem
            self.lbl_icon_vazio.setStyleSheet("font-size: 40px;")

        # 2. Texto
        self.lbl_txt_vazio = QLabel("Ainda não há nada aqui")
        self.lbl_txt_vazio.setAlignment(Qt.AlignCenter)
        # Configuração da Fonte Calibri Negrito solicitada
        self.lbl_txt_vazio.setStyleSheet("""
            QLabel {
                font-family: 'Calibri';
                font-weight: bold;
                font-size: 24px;
                color: #E0E0E0; /* Cor clara para contrastar com o fundo escuro */
                background: transparent;
            }
        """)

        ly_vazio.addWidget(self.lbl_icon_vazio)
        ly_vazio.addWidget(self.lbl_txt_vazio)
        
        # Inicialmente escondido
        self.widget_vazio.hide()

        # --- Hook (Gancho) para manter centralizado ao redimensionar a tabela ---
        # Salvamos o resizeEvent original da tabela para não quebrá-lo
        self.original_resize_table = self.table.resizeEvent
        # Substituímos pelo nosso que chama o centralizar
        self.table.resizeEvent = self.ao_redimensionar_tabela

    def ao_redimensionar_tabela(self, event):
        if event: self.original_resize_table(event)
        """Mantém o placeholder centralizado quando a tabela muda de tamanho."""
        # Chama o comportamento original da tabela (importante!)
        self.original_resize_table(event)
        
        # Centraliza o widget de aviso
        if hasattr(self, 'widget_vazio') and self.widget_vazio.isVisible():
            geo = self.table.viewport().geometry()
            w_vazio = self.widget_vazio.sizeHint().width()
            h_vazio = self.widget_vazio.sizeHint().height()
            
            x = (geo.width() - w_vazio) // 2
            y = (geo.height() - h_vazio) // 2
            
            self.widget_vazio.setGeometry(x, y, w_vazio, h_vazio)
=======
    
    # === MÉTODOS DE UPDATE ===
    def mostrar_aviso_update(self, tag, url):
        msg = (f"Nova versão {tag} disponível!\nDeseja atualizar agora?")
        resp = QMessageBox.question(self, "Atualização Disponível", msg)
        
        if resp == QMessageBox.Yes:
            self.iniciar_download_update(url)

    def iniciar_download_update(self, url):
        # Cria diálogo de progresso
        progress = QProgressDialog("Baixando atualização...", "Cancelar", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setAutoClose(False) # Não fecha sozinho no 100% para dar tempo de ver
        progress.show()
        
        upd = updater.Updater(self.VERSAO_ATUAL)
        
        # Função interna para atualizar a barra
        def atualizar_barra(valor):
            progress.setValue(valor)
            QApplication.processEvents() # Mantém a GUI responsiva
        
        # Executa o download (bloqueante para simplificar, mas com processEvents na callback)
        sucesso = upd.baixar_atualizacao(url, callback_progresso=atualizar_barra)
        
        progress.close()
        
        if sucesso:
            QMessageBox.information(self, "Sucesso", "O Invenio será reiniciado para aplicar a atualização.")
            upd.aplicar_atualizacao() # Isso fecha o programa
        else:
            QMessageBox.critical(self, "Erro", "Falha ao baixar atualização. Verifique sua conexão.")
>>>>>>> b9b3ae538bf17b7da839395c142a2e674acef179

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Aplica configurações de High DPI também via Qt (garantia extra)
    app.setAttribute(Qt.AA_EnableHighDpiScaling)
    app.setAttribute(Qt.AA_UseHighDpiPixmaps)
    
    window = MainWindow()
    # A janela é exibida dentro do __init__ com base nas preferências salvas
    if not window.isVisible():
        window.show()
        
    sys.exit(app.exec())