import os
import win32com.client
import ctypes
import tempfile
import subprocess

def obter_app():
    try:
        return win32com.client.GetActiveObject("Inventor.Application")
    except:
        return None

def salvar_peca(doc, caminho_completo, titulo, projeto, descricao, part_number):
    try:
        props = doc.PropertySets
        props["Summary Information"]["Title"].Value = titulo
        props["Summary Information"]["Revision Number"].Value = "01"
        props["Design Tracking Properties"]["Project"].Value = projeto
        props["Design Tracking Properties"]["Description"].Value = descricao
        props["Design Tracking Properties"]["Part Number"].Value = part_number
    except: pass 
    doc.SaveAs(caminho_completo, False)

def salvar_idw(doc, caminho_idw):
    try:
        doc.PropertySets["Summary Information"]["Revision Number"].Value = "01"
    except: pass
    doc.SaveAs(caminho_idw, False)

def atualizar_propriedades(app, caminho_arquivo, novos_dados):
    doc_fechar = False
    doc = None
    for d in app.Documents:
        if d.FullFileName.lower() == caminho_arquivo.lower():
            doc = d; break
    if not doc:
        try:
            doc = app.Documents.Open(caminho_arquivo, False)
            doc_fechar = True
        except: return False

    try:
        props = doc.PropertySets
        if 'titulo' in novos_dados: props["Summary Information"]["Title"].Value = novos_dados['titulo']
        if 'projeto' in novos_dados: props["Design Tracking Properties"]["Project"].Value = novos_dados['projeto']
        if 'descricao' in novos_dados: props["Design Tracking Properties"]["Description"].Value = novos_dados['descricao']
        doc.Save()
        if doc_fechar: doc.Close(True)
        return True
    except:
        if doc_fechar: doc.Close(True)
        return False

def capturar_print(app, pasta_raiz, codigo):
    try:
        pasta_imgs = os.path.join(pasta_raiz, "_IMAGENS")
        os.makedirs(pasta_imgs, exist_ok=True)
        caminho_img = os.path.join(pasta_imgs, f"{codigo}.jpg")
        cam = app.ActiveView.Camera
        cam.Fit(); cam.Apply()
        cam.SaveAsBitmap(caminho_img, 400, 400)
    except: pass

def abrir_arquivo(app, caminho_arquivo):
    if not os.path.exists(caminho_arquivo):
        raise Exception("Arquivo não encontrado.")
    doc = app.Documents.Open(caminho_arquivo, True)
    try:
        app.Visible = True
        hwnd = app.MainFrameHWND
        ctypes.windll.user32.ShowWindow(hwnd, 3) # SW_MAXIMIZE
        ctypes.windll.user32.SetForegroundWindow(hwnd)
    except: pass
    return doc

def inserir_componente_montagem(app, caminho_arquivo):
    if not os.path.exists(caminho_arquivo): raise Exception("Arquivo não encontrado.")
    doc = app.ActiveDocument
    if not doc or doc.DocumentType != 12291: raise Exception("Abra uma montagem (.iam).")

    try:
        trans_geom = app.TransientGeometry
        matrix = trans_geom.CreateMatrix() 
        doc.ComponentDefinition.Occurrences.Add(caminho_arquivo, matrix)
        
        app.Visible = True
        hwnd = app.MainFrameHWND
        ctypes.windll.user32.ShowWindow(hwnd, 3) 
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        
    except Exception as e:
        raise Exception(f"Erro ao inserir componente: {e}")

def executar_ilogic(app, codigo_vb):
    doc = app.ActiveDocument
    if not doc: raise Exception("Nenhum documento aberto.")

    caminho_temp = ""
    try:
        f_temp = tempfile.NamedTemporaryFile(mode='w', suffix='.iLogicVb', delete=False, encoding='utf-8')
        f_temp.write(codigo_vb)
        caminho_temp = f_temp.name
        f_temp.close()

        iLogicAddin = app.ApplicationAddIns.ItemById("{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}")
        automation = iLogicAddin.Automation
        automation.RunExternalRule(doc, caminho_temp)
        
    except Exception as e:
        raise Exception(f"Erro ao executar iLogic: {e}")
    finally:
        if caminho_temp and os.path.exists(caminho_temp):
            try: os.remove(caminho_temp)
            except: pass

def configurar_content_center(app, pasta_raiz):
    """
    Modifica o .ipj ativo para apontar o Content Center para 'pasta_raiz/Content Center Files'.
    """
    if not app: return False, "Inventor fechado."
    
    try:
        # 1. Define qual será o caminho final da biblioteca
        caminho_cc_alvo = os.path.join(pasta_raiz, "Content Center Files")
        
        # 2. Garante que a pasta existe (O Inventor exige isso antes de configurar)
        if not os.path.exists(caminho_cc_alvo):
            try:
                os.makedirs(caminho_cc_alvo)
            except Exception as e:
                return False, f"Erro ao criar pasta CC: {e}"

        # 3. Acessa o Gerenciador de Projetos
        manager = app.DesignProjectManager
        projeto_ativo = manager.ActiveDesignProject
        
        # 4. Verifica se já está configurado (Evita salvar sem necessidade)
        caminho_atual = projeto_ativo.ContentCenterPath
        
        # Normaliza as barras para comparar paths (Windows usa \ mas as vezes vem /)
        path_a = os.path.normpath(caminho_atual).lower()
        path_b = os.path.normpath(caminho_cc_alvo).lower()
        
        if path_a != path_b:
            try:
                # --- AQUI ACONTECE A MÁGICA ---
                projeto_ativo.ContentCenterPath = caminho_cc_alvo
                projeto_ativo.Save() # Salva a alteração no arquivo .ipj
                # ------------------------------
                return True, f"IPJ Atualizado para:\n{caminho_cc_alvo}"
            except Exception as e:
                return False, f"Erro ao salvar .ipj (Verifique se é Somente Leitura): {e}"
        else:
            return True, "IPJ já estava correto."
            
    except Exception as e:
        return False, f"Erro na API do Inventor: {e}"