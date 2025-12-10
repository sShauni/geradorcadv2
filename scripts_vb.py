# scripts_vb.py

# Script 1: Exportar DXF e Excel
SCRIPT_EXPORTAR_LASER = r'''
' Regra iLogic: Exportar DXF + Arquivo Excel (.xlsx) Formatado
' Adaptado para injeção via Python

Imports System.Collections.Generic
Imports System.IO

Sub Main()
    Dim oDoc As Document = ThisApplication.ActiveDocument
    If oDoc.DocumentType <> kAssemblyDocumentObject Then
        MessageBox.Show("Execute em uma Montagem (.iam)", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Exit Sub
    End If

    Dim oAsmDoc As AssemblyDocument = oDoc
    
    ' Caminho Desktop
    Dim sPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
    Dim sFolder As String = sPath & "\DXF_Corte_Laser"
    
    If Not System.IO.Directory.Exists(sFolder) Then
        System.IO.Directory.CreateDirectory(sFolder)
    End If

    Dim sExcelFile As String = sFolder & "\Lista_de_Corte.xlsx"
    Dim exportedCount As Integer = 0
    Dim dictQty As New Dictionary(Of String, Integer)
    Dim dictDocs As New Dictionary(Of String, PartDocument)
    
    TraverseAssembly(oAsmDoc.ComponentDefinition.Occurrences, dictQty, dictDocs)
    
    Try
        Dim oApp As Object = CreateObject("Excel.Application")
        oApp.Visible = False
        Dim oBook As Object = oApp.Workbooks.Add
        Dim oSheet As Object = oBook.ActiveSheet
        
        Dim headers() As String = {"Item", "Número da Peça", "Espessura (mm)", "Material", "Quantidade", "Nome do Arquivo Gerado"}
        For i As Integer = 0 To headers.Length - 1
            oSheet.Cells(1, i + 1).Value = headers(i)
        Next
        
        oSheet.Range("A1:F1").Font.Bold = True
        oSheet.Range("A1:F1").Interior.Color = 14277081 
        
        Dim row As Integer = 2
        Dim itemCount As Integer = 1
        
        For Each kvp As KeyValuePair(Of String, PartDocument) In dictDocs
            Dim oPartDoc As PartDocument = kvp.Value
            Dim qty As Integer = dictQty(kvp.Key)
            
            Dim sPartNumber As String = GetPartNumber(oPartDoc)
            Dim sMaterial As String = oPartDoc.ComponentDefinition.Material.Name
            Dim oCompDef As SheetMetalComponentDefinition = oPartDoc.ComponentDefinition
            Dim dThickness As Double = oCompDef.Thickness.Value * 10
            Dim sThickness As String = dThickness.ToString("0.##").Replace(".", ",")
            
            Dim finalFileName As String = FormatFileName(qty, sThickness, sMaterial, sPartNumber)
            
            oSheet.Cells(row, 1).Value = itemCount
            oSheet.Cells(row, 2).Value = sPartNumber
            oSheet.Cells(row, 3).Value = sThickness
            oSheet.Cells(row, 4).Value = sMaterial
            oSheet.Cells(row, 5).Value = qty
            oSheet.Cells(row, 6).Value = finalFileName
            
            If ExportDXF(oPartDoc, sFolder, finalFileName) Then
                exportedCount += 1
            End If
            
            row += 1
            itemCount += 1
        Next
        
        Dim dataRange As Object = oSheet.Range("A1:F" & (row - 1))
        dataRange.HorizontalAlignment = -4108
        dataRange.VerticalAlignment = -4108
        dataRange.Borders.LineStyle = 1
        oSheet.Columns("A:F").AutoFit()
        
        If System.IO.File.Exists(sExcelFile) Then System.IO.File.Delete(sExcelFile)
        
        oBook.SaveAs(sExcelFile)
        oBook.Close()
        oApp.Quit()
        
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oSheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
        
        Process.Start("explorer.exe", sFolder)
        MessageBox.Show("Sucesso! " & exportedCount & " DXFs gerados.", "Concluído")
        
    Catch ex As Exception
        MessageBox.Show("Erro Excel: " & ex.Message)
    End Try
End Sub

Sub TraverseAssembly(Occurrences As ComponentOccurrences, ByRef dictQty As Dictionary(Of String, Integer), ByRef dictDocs As Dictionary(Of String, PartDocument))
    Dim oOcc As ComponentOccurrence
    For Each oOcc In Occurrences
        If oOcc.Suppressed Then Continue For
        If oOcc.DefinitionDocumentType = kAssemblyDocumentObject Then
            TraverseAssembly(oOcc.SubOccurrences, dictQty, dictDocs)
        ElseIf oOcc.DefinitionDocumentType = kPartDocumentObject Then
            Dim oPartDoc As PartDocument
            Try
                oPartDoc = oOcc.Definition.Document
            Catch
                Continue For
            End Try
            ' Verifica se é chapa metálica pelo SubType
            If oPartDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                Dim docName As String = oPartDoc.FullFileName
                If dictQty.ContainsKey(docName) Then
                    dictQty(docName) += 1
                Else
                    dictQty.Add(docName, 1)
                    dictDocs.Add(docName, oPartDoc)
                End If
            End If
        End If
    Next
End Sub

Function GetPartNumber(oDoc As Document) As String
    Try
        Dim pn As String = oDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
        If String.IsNullOrEmpty(pn) Then Return System.IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName)
        Return pn
    Catch
        Return "ERRO_PN"
    End Try
End Function

Function FormatFileName(Qty As Integer, Thickness As String, Material As String, PartNum As String) As String
    Dim name As String = String.Format("{0}X{1}MM-{2} - {3}", Qty, Thickness, Material, PartNum)
    Dim invalidChars As Char() = System.IO.Path.GetInvalidFileNameChars()
    For Each c As Char In invalidChars
        name = name.Replace(c, "_")
    Next
    Return name
End Function

Function ExportDXF(oPartDoc As PartDocument, FolderPath As String, FileName As String) As Boolean
    Try
        Dim oCompDef As SheetMetalComponentDefinition = oPartDoc.ComponentDefinition
        If Not oCompDef.HasFlatPattern Then
            Try
                oCompDef.Unfold()
                oCompDef.FlatPattern.ExitEdit()
            Catch
                Return False
            End Try
        End If
        Dim sOutFile As String = FolderPath & "\" & FileName & ".dxf"
        Dim sOut As String = "FLAT PATTERN DXF?AcadVersion=2004&OuterProfileLayer=CORTE&InnerProfileLayer=CORTE" _
             & "&InvisibleLayers=IV_ARC_CENTERS;IV_TANGENT;IV_ROLL;IV_ROLL_TANGENT;IV_ALTREP_BACK;IV_ALTREP_FRONT;IV_FEATURE_PROFILES_DOWN;IV_FEATURE_PROFILES_UP" _
             & "&SimplifySplines=True&MergeProfiles=True&RebaseGeometry=True"
        oCompDef.DataIO.WriteDataToFile(sOut, sOutFile)
        Return True
    Catch ex As Exception
        Return False
    End Try
End Function
'''

# Script 2: Lista de Fixadores
SCRIPT_LISTA_FIXADORES = r'''
' Regra iLogic: Lista de Fixadores para Excel
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Public Class FixadorInfo
    Public PartNumber As String
    Public Norma As String
    Public Tipo As String
    Public Descricao As String
    Public Diametro As String
    Public Comprimento As String
    Public Quantidade As Integer
    Public Sub New()
        Quantidade = 0
    End Sub
End Class

Sub Main()
    If ThisApplication.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
        MessageBox.Show("Execute em uma Montagem (IAM).", "Erro")
        Exit Sub
    End If

    Dim oAsmDoc As AssemblyDocument = ThisApplication.ActiveDocument
    Dim listaFixadores As New Dictionary(Of String, FixadorInfo)
    
    TraverseAssembly(oAsmDoc.ComponentDefinition.Occurrences, listaFixadores)
    
    If listaFixadores.Count = 0 Then
        MessageBox.Show("Nenhum fixador encontrado.", "Aviso")
        Exit Sub
    End If
    
    ExportarParaExcel(listaFixadores)
End Sub

Sub TraverseAssembly(Occurrences As ComponentOccurrences, ByRef lista As Dictionary(Of String, FixadorInfo))
    For Each occ As ComponentOccurrence In Occurrences
        If occ.Suppressed Then Continue For
        
        If occ.DefinitionDocumentType = kAssemblyDocumentObject Then
            TraverseAssembly(occ.SubOccurrences, lista)
        ElseIf occ.DefinitionDocumentType = kPartDocumentObject Then
            If occ.Definition.IsContentMember Then
                ProcessarPeca(occ, lista)
            End If
        End If
    Next
End Sub

Sub ProcessarPeca(occ As ComponentOccurrence, ByRef lista As Dictionary(Of String, FixadorInfo))
    Try
        Dim doc As Document = occ.Definition.Document
        Dim props As PropertySets = doc.PropertySets
        Dim partNumber As String = ""
        Try
            partNumber = props("Design Tracking Properties").Item("Part Number").Value
        Catch
            partNumber = occ.Name
        End Try

        If lista.ContainsKey(partNumber) Then
            lista(partNumber).Quantidade += 1
        Else
            Dim novoFix As New FixadorInfo
            novoFix.PartNumber = partNumber
            novoFix.Quantidade = 1
            novoFix.Descricao = GetProp(props, "Design Tracking Properties", "Description")
            novoFix.Norma = GetProp(props, "Design Tracking Properties", "Standard")
            If novoFix.Norma = "-" Or novoFix.Norma = "" Then novoFix.Norma = GetProp(props, "Design Tracking Properties", "Authority")
            novoFix.Tipo = GetProp(props, "Summary Information", "Title")
            
            Dim reqParams As Parameters = occ.Definition.Parameters
            Dim sDiam As String = "-"
            Dim paramDnames As String() = {"NND", "d", "D", "Thread1_NominalDiameter", "NominalDiameter", "SIZE", "G"}
            For Each nome In paramDnames
                If TryGetParam(reqParams, nome, sDiam) Then Exit For
            Next
            
            Dim sComp As String = "-"
            Dim paramLnames As String() = {"NLG", "NominalLength", "B_L", "L", "l", "LG", "Length", "OAL", "S_L", "p1"}
            Dim achouComp As Boolean = False
            For Each nome In paramLnames
                If TryGetParam(reqParams, nome, sComp) Then 
                    achouComp = True
                    Exit For
                End If
            Next
            
            If Not achouComp Or sComp = "-" Or sComp = "0" Then
                sComp = ExtrairComprimentoDaDescricao(novoFix.Descricao)
            End If
            
            novoFix.Diametro = sDiam
            novoFix.Comprimento = sComp
            lista.Add(partNumber, novoFix)
        End If
    Catch
    End Try
End Sub

Function ExtrairComprimentoDaDescricao(desc As String) As String
    Try
        Dim padrao As String = "[xX]\s*(\d+)"
        Dim match As Match = Regex.Match(desc, padrao)
        If match.Success Then Return match.Groups(1).Value
        Return "-"
    Catch
        Return "-"
    End Try
End Function

Sub ExportarParaExcel(lista As Dictionary(Of String, FixadorInfo))
    Dim oExcelApp As Object = Nothing
    Dim oWorkbook As Object = Nothing
    Dim oSheet As Object = Nothing
    Try
        oExcelApp = CreateObject("Excel.Application")
        oExcelApp.Visible = True
        oWorkbook = oExcelApp.Workbooks.Add
        oSheet = oWorkbook.Sheets(1)
        
        Dim headers() As String = {"Item", "Norma", "Tipo", "Diâmetro", "Comprimento", "Qtd", "Descrição Completa"}
        For i As Integer = 0 To headers.Length - 1
            oSheet.Cells(1, i + 1).Value = headers(i)
        Next
        
        Dim headerRange As Object = oSheet.Range("A1:G1")
        headerRange.Font.Bold = True
        headerRange.Interior.Color = 14277081 
        
        Dim row As Integer = 2
        Dim itemNum As Integer = 1
        Dim sortedKeys As List(Of String) = New List(Of String)(lista.Keys)
        sortedKeys.Sort()

        For Each key As String In sortedKeys
            Dim fix As FixadorInfo = lista(key)
            oSheet.Cells(row, 1).Value = itemNum
            oSheet.Cells(row, 2).Value = fix.Norma
            oSheet.Cells(row, 3).Value = fix.Tipo
            oSheet.Cells(row, 4).Value = fix.Diametro
            oSheet.Cells(row, 5).Value = fix.Comprimento
            oSheet.Cells(row, 6).Value = fix.Quantidade
            oSheet.Cells(row, 7).Value = fix.Descricao
            row += 1
            itemNum += 1
        Next
        
        Dim lastRow As Integer = row - 1
        If lastRow >= 2 Then
            Dim dataRange As Object = oSheet.Range("A1:G" & lastRow)
            dataRange.Borders.LineStyle = 1 
            oSheet.Columns.AutoFit
        End If
    Catch ex As Exception
        MessageBox.Show("Erro Excel: " & ex.Message)
    End Try
End Sub

Function GetProp(props As PropertySets, setNames As String, propName As String) As String
    Try
        Dim val As Object = props(setNames).Item(propName).Value
        If val Is Nothing Then Return "-"
        Return val.ToString()
    Catch
        Return "-"
    End Try
End Function

Function TryGetParam(params As Parameters, paramName As String, ByRef result As String) As Boolean
    Try
        Dim p As Parameter = params.Item(paramName)
        result = p.Expression 
        If String.IsNullOrEmpty(result) Then result = p.Value.ToString("0.##")
        result = result.Replace("ul", "") 
        Return True
    Catch
        Return False
    End Try
End Function
'''