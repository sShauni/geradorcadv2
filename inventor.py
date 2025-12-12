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

# --- NOVA FUNÇÃO: CONFIGURAR BIBLIOTECA NO SERVIDOR ---
def configurar_content_center(app, pasta_rede):
    """
    Configura o caminho dos arquivos do Content Center.
    Retorna (Sucesso, Mensagem de Erro/Sucesso)
    """
    try:
        # 1. Tenta criar a pasta na rede (Testa permissão de escrita)
        caminho_cc_rede = os.path.join(pasta_rede, "Content Center Files")
        if not os.path.exists(caminho_cc_rede):
            try:
                os.makedirs(caminho_cc_rede)
            except Exception as e:
                return False, f"Sem permissão para criar pasta no servidor:\n{e}"
            
        # 2. Acessa o Projeto Ativo
        design_proj = app.DesignProjectManager.ActiveDesignProject
        
        # Verifica se o projeto permite edição
        # (Alguns projetos são 'Somente Leitura' ou Single User travados)
        if not design_proj:
            return False, "Nenhum projeto ativo encontrado."

        # 3. Aplica a mudança
        caminho_atual = design_proj.ContentCenterPath
        
        if caminho_atual.lower() != caminho_cc_rede.lower():
            # Tenta definir o caminho
            try:
                design_proj.ContentCenterPath = caminho_cc_rede
                design_proj.Save()
            except Exception as e:
                return False, f"O Inventor recusou o caminho de rede.\nVerifique se o Projeto (.ipj) não é somente leitura.\nErro: {e}"
                
            return True, f"Configurado para:\n{caminho_cc_rede}"
        else:
            return True, "Já configurado corretamente."
            
    except Exception as e:
        return False, f"Erro genérico ao configurar: {e}"