# inventor.py
import os
import win32com.client

def obter_app():
    try:
        return win32com.client.GetActiveObject("Inventor.Application")
    except:
        return None

def salvar_peca(doc, caminho_completo, titulo, projeto, descricao, part_number):
    # Preenche iProperties
    try:
        props = doc.PropertySets
        props["Summary Information"]["Title"].Value = titulo
        props["Design Tracking Properties"]["Project"].Value = projeto
        props["Design Tracking Properties"]["Description"].Value = descricao
        props["Design Tracking Properties"]["Part Number"].Value = part_number
    except: pass # Ignora erro de propriedade se falhar

    # Salva
    doc.SaveAs(caminho_completo, False)

def salvar_idw(doc, caminho_idw):
    doc.SaveAs(caminho_idw, False)

def capturar_print(app, pasta_raiz, codigo):
    try:
        pasta_imgs = os.path.join(pasta_raiz, "_IMAGENS")
        os.makedirs(pasta_imgs, exist_ok=True)
        caminho_img = os.path.join(pasta_imgs, f"{codigo}.jpg")
        
        cam = app.ActiveView.Camera
        cam.Fit()
        cam.Apply()
        cam.SaveAsBitmap(caminho_img, 400, 400)
    except: pass

def inserir_componente_montagem(app, caminho_arquivo):
    doc = app.ActiveDocument
    # 12291 = kAssemblyDocumentObject
    if not doc or doc.DocumentType != 12291:
        raise Exception("O documento ativo não é uma montagem (.iam).")

    trans_geom = app.TransientGeometry
    matrix = trans_geom.CreateMatrix()
    doc.ComponentDefinition.Occurrences.Add(caminho_arquivo, matrix)
    
    try: app.Visible = True
    except: pass