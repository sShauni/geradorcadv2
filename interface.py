import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import threading
import subprocess
from datetime import datetime
from PIL import Image, ImageTk

import config
import dados
import inventor

class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador CAD V22 (Modular)")
        self.root.geometry("1150x850")
        
        # --- NOVO: TROCAR ÍCONE DA JANELA ---
        try:
            # Usa a função nova do config para achar o ícone DENTRO do exe
            caminho_icone = config.obter_caminho_recurso("icon.ico")
            
            # O .iconbitmap do Tkinter às vezes falha se o arquivo não existir
            if os.path.exists(caminho_icone):
                self.root.iconbitmap(caminho_icone)
            else:
                # Fallback: Se não achar, tenta procurar na pasta local (caso de debug)
                pass 
        except Exception as e:
            print(f"Erro ao carregar ícone interno: {e}")
        # ------------------------------------
        
        # Carrega Configs
        self.cfg = config.carregar()
        self.caminho_db_atual = config.ARQUIVO_CSV_LOCAL
        self.caminho_rede_ativo = None
        
        self.var_mostrar_inativos = tk.BooleanVar(value=False)
        self.var_cod = tk.StringVar()
        
        dados.garantir_csv(self.caminho_db_atual)
        
        self.criar_layout()
        
        # Conexão Rede em Background
        if self.cfg.get("usar_servidor"):
            threading.Thread(target=self.conectar_rede, daemon=True).start()
        else:
            self.atualizar_lista()

    def criar_layout(self):
        painel = tk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        painel.pack(fill="both", expand=True, padx=5, pady=5)
        
        # === ESQUERDA ===
        frame_esq = tk.Frame(painel)
        painel.add(frame_esq, minsize=420)
        
        # Menu
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        menu_opt = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Opções", menu=menu_opt)
        menu_opt.add_command(label="Configurar Servidor...", command=self.janela_servidor)
        
        # Status Rede
        self.lbl_rede = tk.Label(frame_esq, text="Modo: LOCAL", bg="#ddd", pady=5)
        self.lbl_rede.pack(fill="x", pady=2)
        
        # Campos Dados
        fr_dados = tk.LabelFrame(frame_esq, text="Dados do Projeto", padx=10, pady=10)
        fr_dados.pack(fill="x", pady=5)
        
        tk.Label(fr_dados, text="Projeto:").grid(row=0, column=0, sticky="w")
        self.ent_proj = tk.Entry(fr_dados, width=35)
        self.ent_proj.insert(0, self.cfg.get("ultimo_projeto", ""))
        self.ent_proj.grid(row=0, column=1, sticky="w", pady=2)
        
        tk.Label(fr_dados, text="Título:").grid(row=1, column=0, sticky="w")
        self.ent_tit = tk.Entry(fr_dados, width=35)
        self.ent_tit.grid(row=1, column=1, sticky="w", pady=2)
        
        tk.Label(fr_dados, text="Descrição:").grid(row=2, column=0, sticky="w")
        self.ent_desc = tk.Entry(fr_dados, width=35)
        self.ent_desc.grid(row=2, column=1, sticky="w", pady=2)
        
        # === GERADOR (Layout Restaurado) ===
        fr_gen = tk.LabelFrame(frame_esq, text="Gerador", padx=10, pady=10)
        fr_gen.pack(fill="x", pady=5)
        
        # Linha 0: Prefixo e Tipo
        tk.Label(fr_gen, text="Prefixo:").grid(row=0, column=0)
        self.ent_prefixo = tk.Entry(fr_gen, width=10)
        self.ent_prefixo.insert(0, self.cfg.get("ultimo_prefixo", "PRJ"))
        self.ent_prefixo.grid(row=0, column=1)
        
        self.cb_tipo = ttk.Combobox(fr_gen, values=list(config.TIPOS_FABRICACAO.keys()), state="readonly", width=18)
        self.cb_tipo.current(0)
        self.cb_tipo.grid(row=0, column=2, padx=5)
        
        # Linha 1: Display do Código
        tk.Entry(fr_gen, textvariable=self.var_cod, font=("Arial", 14, "bold"), justify="center", state="readonly", bg="#f0f0f0").grid(row=1, column=0, columnspan=3, pady=10, sticky="ew")
        
        # Linha 2: Botão GERAR (Azul Grande)
        tk.Button(fr_gen, text="1. GERAR CÓDIGO", bg="#2196F3", fg="white", font=("Arial", 9, "bold"), command=self.acao_gerar_codigo).grid(row=2, column=0, columnspan=3, sticky="ew")
        
        # Linha 3: Botão SALVAR (Laranja Grande)
        tk.Button(fr_gen, text="2. SALVAR (CAD / IDW)", bg="#FF9800", fg="black", font=("Arial", 9, "bold"), command=self.acao_salvar).grid(row=3, column=0, columnspan=3, sticky="ew", pady=5)

        # Imagem
        fr_img = tk.LabelFrame(frame_esq, text="Visualizador", padx=5, pady=5)
        fr_img.pack(fill="both", expand=True)
        self.lbl_img = tk.Label(fr_img, text="Selecione um item...")
        self.lbl_img.pack(fill="both", expand=True)

        # === DIREITA ===
        frame_dir = tk.Frame(painel)
        painel.add(frame_dir, minsize=500)
        
        # Filtros
        fr_top = tk.Frame(frame_dir)
        fr_top.pack(fill="x", pady=5)
        tk.Label(fr_top, text="Buscar:").pack(side="left")
        self.ent_busca = tk.Entry(fr_top)
        self.ent_busca.pack(side="left", fill="x", expand=True, padx=5)
        self.ent_busca.bind("<KeyRelease>", lambda e: self.atualizar_lista())
        
        tk.Checkbutton(fr_top, text="Lixeira", variable=self.var_mostrar_inativos, command=self.atualizar_lista).pack(side="left", padx=5)
        
        # Botões Lista
        fr_btn_lst = tk.Frame(frame_dir)
        fr_btn_lst.pack(fill="x", pady=2)
        
        tk.Button(fr_btn_lst, text="Abrir CAD", command=self.acao_abrir).pack(side="left", padx=2)
        tk.Button(fr_btn_lst, text="Inserir na Montagem (+)", bg="#e3f2fd", command=self.acao_inserir_montagem).pack(side="left", padx=5)
        tk.Button(fr_btn_lst, text="Excluir", fg="red", command=self.acao_excluir).pack(side="right", padx=2)

        # Treeview
        cols = ("Codigo", "Tipo", "Titulo", "Descricao", "Status")
        self.tree = ttk.Treeview(frame_dir, columns=cols, show="headings")
        for c in cols: self.tree.heading(c, text=c)
        self.tree.column("Codigo", width=100); self.tree.column("Status", width=60)
        
        self.tree.tag_configure("INATIVO", foreground="gray")
        self.tree.tag_configure("DESENHO", foreground="blue")
        
        scroll = ttk.Scrollbar(frame_dir, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scroll.set)
        scroll.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.ao_selecionar)
        self.tree.bind("<Double-1>", lambda e: self.acao_abrir())

    # --- LÓGICA DE INTERFACE ---
    
    def acao_gerar_codigo(self):
        prefixo = self.ent_prefixo.get().upper()
        sufixo = config.TIPOS_FABRICACAO[self.cb_tipo.get()]
        if not prefixo: return
        
        cod = dados.gerar_codigo_unico(self.caminho_db_atual, prefixo, sufixo)
        if cod:
            self.var_cod.set(cod)
            self.root.clipboard_clear()
            self.root.clipboard_append(cod)

    def acao_salvar(self):
        destino = self.caminho_rede_ativo if self.caminho_rede_ativo else filedialog.askdirectory()
        if not destino: return

        app = inventor.obter_app()
        if not app:
            return messagebox.showerror("Erro", "Inventor não encontrado.")
        
        doc = app.ActiveDocument
        if not doc:
            return messagebox.showwarning("Aviso", "Nenhum documento aberto.")

        try:
            # 12292 = IDW/DWG
            if doc.DocumentType == 12292:
                # Salvar Desenho
                nome_base = "Desenho"
                try: 
                    if doc.ReferencedDocuments.Count > 0:
                        ref = doc.ReferencedDocuments.Item(1)
                        nome_base = os.path.splitext(os.path.basename(ref.FullFileName))[0]
                except: pass
                
                pasta = os.path.join(destino, "desenhos", "ED")
                os.makedirs(pasta, exist_ok=True)
                caminho_final = os.path.join(pasta, f"{nome_base}-DT.idw")
                
                inventor.salvar_idw(doc, caminho_final)
                self.registrar_db(nome_base, "DESENHO TÉCNICO", caminho_final, f"[DESENHO] {self.ent_tit.get()}")
                messagebox.showinfo("Sucesso", "Desenho Salvo!")
                
            else:
                # Salvar Peça/Montagem
                cod = self.var_cod.get()
                if not cod: return messagebox.showwarning("Aviso", "Gere um código!")
                
                ext = ".ipt" if doc.DocumentType != 12291 else ".iam"
                caminho_final = os.path.join(destino, cod + ext)
                
                inventor.salvar_peca(doc, caminho_final, self.ent_tit.get(), self.ent_proj.get(), self.ent_desc.get(), cod)
                inventor.capturar_print(app, destino, cod)
                
                self.registrar_db(cod, self.cb_tipo.get(), caminho_final, self.ent_tit.get())
                messagebox.showinfo("Sucesso", "Peça Salva!")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar:\n{e}")

    def acao_inserir_montagem(self):
        sel = self.tree.selection()
        if not sel: return messagebox.showwarning("Aviso", "Selecione uma peça na lista.")
        
        caminho = self.tree.item(sel[0])['tags'][1]
        
        if not caminho or not os.path.exists(caminho):
            return messagebox.showerror("Erro", "Arquivo não encontrado no disco.")
            
        app = inventor.obter_app()
        if not app: return messagebox.showerror("Erro", "Inventor fechado.")
        
        try:
            inventor.inserir_componente_montagem(app, caminho)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def registrar_db(self, cod, tipo, caminho, titulo):
        linha = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            cod, self.ent_prefixo.get(), "", tipo,
            self.ent_proj.get(), titulo, self.ent_desc.get(),
            "ATIVO", caminho
        ]
        dados.gravar_linha(self.caminho_db_atual, linha)
        
        # Salva config user
        self.cfg["ultimo_prefixo"] = self.ent_prefixo.get()
        self.cfg["ultimo_projeto"] = self.ent_proj.get()
        config.salvar(self.cfg)
        
        self.atualizar_lista()

    def atualizar_lista(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        
        registros = dados.ler_registros(self.caminho_db_atual, self.var_mostrar_inativos.get(), self.ent_busca.get())
        
        for row in registros:
            tag = "ATIVO"
            if row[8] == "INATIVO": tag = "INATIVO"
            elif "DESENHO" in row[4].upper(): tag = "DESENHO"
            
            self.tree.insert("", "end", values=(row[1], row[4], row[6], row[7], row[8]), tags=(tag, row[9]))

    def ao_selecionar(self, e):
        # Lógica simplificada de imagem
        self.lbl_img.config(image="", text=""); self.lbl_img.img = None
        sel = self.tree.selection()
        if not sel: return
        
        caminho = self.tree.item(sel[0])['tags'][1]
        cod = self.tree.item(sel[0])['values'][0]
        
        # Tenta achar imagem
        if caminho and os.path.exists(os.path.dirname(caminho)):
            pasta = os.path.dirname(caminho)
            if "desenhos" in pasta.lower(): pasta = os.path.dirname(os.path.dirname(pasta))
            
            img_path = os.path.join(pasta, "_IMAGENS", f"{cod}.jpg")
            if os.path.exists(img_path):
                try:
                    img = Image.open(img_path)
                    img.thumbnail((350, 350))
                    tk_img = ImageTk.PhotoImage(img)
                    self.lbl_img.config(image=tk_img)
                    self.lbl_img.img = tk_img
                except: pass
            else:
                self.lbl_img.config(text="Sem Imagem")

    def acao_abrir(self):
        sel = self.tree.selection()
        if sel:
            path = self.tree.item(sel[0])['tags'][1]
            if path and os.path.exists(path): os.startfile(path)

    def acao_excluir(self):
        sel = self.tree.selection()
        if not sel: return
        cod = self.tree.item(sel[0])['values'][0]
        caminho = self.tree.item(sel[0])['tags'][1]
        
        if messagebox.askyesno("Excluir", f"Inativar {cod}?"):
            dados.excluir_logico(self.caminho_db_atual, cod, caminho)
            self.atualizar_lista()

    def conectar_rede(self):
        ip = self.cfg.get("ip"); path = self.cfg.get("path"); user = self.cfg.get("user"); senha = self.cfg.get("pass")
        unc = f"\\\\{ip}\\{path}"
        
        if not os.path.exists(unc):
            try:
                subprocess.run(f'net use "{unc}" /user:{user} "{senha}"', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except: pass
            
        if os.path.exists(unc):
            self.caminho_rede_ativo = unc
            self.caminho_db_atual = os.path.join(unc, "registro_pecas.csv")
            dados.garantir_csv(self.caminho_db_atual)
            self.root.after(0, lambda: self.lbl_rede.config(text=f"Servidor: {ip} (OK)", bg="#ccffcc"))
            self.root.after(0, self.atualizar_lista)
        else:
            self.root.after(0, lambda: self.lbl_rede.config(text="Modo: LOCAL (Falha Rede)", bg="#ffcccc"))

    def janela_servidor(self):
        top = tk.Toplevel(self.root)
        top.title("Config")
        top.geometry("300x350") # Ajustei um pouco o tamanho
        
        tk.Label(top, text="IP:").pack()
        ip = tk.Entry(top)
        ip.pack()
        
        tk.Label(top, text="Pasta:").pack()
        path = tk.Entry(top)
        path.pack()
        
        tk.Label(top, text="User:").pack()
        user = tk.Entry(top)
        user.pack()
        
        tk.Label(top, text="Senha:").pack()
        pw = tk.Entry(top, show="*")
        pw.pack()
        
        # Preenche atuais (Usa valor salvo ou vazio)
        ip.insert(0, self.cfg.get("ip", ""))
        path.insert(0, self.cfg.get("path", ""))
        user.insert(0, self.cfg.get("user", ""))
        pw.insert(0, self.cfg.get("pass", ""))
        
        # --- CORREÇÃO DA VARIÁVEL DE CONTROLE ---
        # Usamos self.var_check_servidor para garantir que o Tkinter não perca a referência
        val_atual = self.cfg.get("usar_servidor", False)
        self.var_check_servidor = tk.BooleanVar(value=val_atual)
        
        tk.Checkbutton(top, text="Usar Servidor", variable=self.var_check_servidor).pack(pady=10)
        
        def salvar_config():
            # Atualiza o dicionário local
            self.cfg["ip"] = ip.get()
            self.cfg["path"] = path.get()
            self.cfg["user"] = user.get()
            self.cfg["pass"] = pw.get()
            self.cfg["usar_servidor"] = self.var_check_servidor.get() # Pega o valor booleano
            
            # Manda salvar no disco
            config.salvar(self.cfg)
            
            top.destroy()
            messagebox.showinfo("Sucesso", "Configurações salvas!\nReinicie o programa para conectar.")

        tk.Button(top, text="Salvar", bg="#4CAF50", fg="white", command=salvar_config).pack(pady=5)