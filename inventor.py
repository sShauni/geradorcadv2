import os
import win32com.client
import ctypes

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
    
    # Foca no Inventor
    try:
        app.Visible = True
        hwnd = app.MainFrameHWND
        ctypes.windll.user32.ShowWindow(hwnd, 9) # Restore
        ctypes.windll.user32.SetForegroundWindow(hwnd)
    except: pass
    
    return doc

def inserir_componente_montagem(app, caminho_arquivo):
    """
    Método Estável: Insere na origem (0,0,0) e foca a janela.
    """
    if not os.path.exists(caminho_arquivo):
        raise Exception("Arquivo não encontrado.")

    doc = app.ActiveDocument
    if not doc or doc.DocumentType != 12291: # 12291 = IAM
        raise Exception("Abra uma montagem (.iam) para inserir.")

    try:
        # 1. Insere na Matriz Identidade (0,0,0 sem rotação)
        trans_geom = app.TransientGeometry
        matrix = trans_geom.CreateMatrix() 
        doc.ComponentDefinition.Occurrences.Add(caminho_arquivo, matrix)
        
        # 2. Traz o Inventor para frente
        app.Visible = True
        hwnd = app.MainFrameHWND
        ctypes.windll.user32.ShowWindow(hwnd, 9) # 9 = SW_RESTORE (Restaura se minimizado)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        
    except Exception as e:
        raise Exception(f"Erro ao inserir componente: {e}")