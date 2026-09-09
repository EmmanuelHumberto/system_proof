Option Explicit

Const vbext_ct_StdModule = 1
Const msoAutomationSecurityLow = 1
Const xlOpenXMLWorkbookMacroEnabled = 52

Dim fso, base, arquivoXlsm, moduloBas, logPath, logFile
Dim xl, wb, vbproj, arquivoTemp, tentativa

Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
arquivoXlsm = base & "\Planilha_Comprovantes_VBA.xlsm"
arquivoTemp = base & "\Planilha_Comprovantes_VBA_vba.xlsm"
moduloBas = base & "\vba\modulos_vba.bas"
logPath = base & "\tools\importar_modulo_excel.log"
Set logFile = fso.OpenTextFile(logPath, 2, True, -1)

On Error Resume Next
Log "Inicio"

Set xl = CreateObject("Excel.Application")
Check "CreateObject Excel.Application"
xl.Visible = False
xl.DisplayAlerts = False
xl.AutomationSecurity = msoAutomationSecurityLow

Set wb = xl.Workbooks.Open(arquivoXlsm, False, False)
Check "Abrir workbook"
Set vbproj = wb.VBProject
Check "Acessar VBProject"

RemoveAllStdModules vbproj
vbproj.VBComponents.Import moduloBas
Call AtualizarEventoDespesas(vbproj)
Call AtualizarEventoControle(vbproj)
Check "Importar ModuloComprovantes"

If fso.FileExists(arquivoTemp) Then fso.DeleteFile arquivoTemp, True
wb.SaveCopyAs arquivoTemp
Check "Salvar copia xlsm"
wb.Close False
xl.Quit
Check "Fechar Excel"

If Not fso.FileExists(arquivoTemp) Then
    Log "ERRO: copia temporaria nao foi criada"
    logFile.Close
    WScript.Quit 1
End If

For tentativa = 1 To 10
    Err.Clear
    fso.CopyFile arquivoTemp, arquivoXlsm, True
    If Err.Number = 0 Then Exit For
    WScript.Sleep 1000
Next
Check "Substituir workbook original"
fso.DeleteFile arquivoTemp, True

Log "Fim OK"
logFile.Close
WScript.Echo "Modulo VBA importado com sucesso em: " & arquivoXlsm


Sub AtualizarEventoDespesas(proj)
    On Error Resume Next
    Dim comp, cm, code, startLine, lineCount
    Set comp = ComponenteDaAba(proj, "Despesas")
    If comp Is Nothing Then Set comp = proj.VBComponents("Plan1")
    Set cm = comp.CodeModule
    startLine = cm.ProcStartLine("Worksheet_FollowHyperlink", 0)
    If Err.Number = 0 Then
        lineCount = cm.ProcCountLines("Worksheet_FollowHyperlink", 0)
        cm.DeleteLines startLine, lineCount
    End If
    Err.Clear
    startLine = cm.ProcStartLine("Worksheet_Activate", 0)
    If Err.Number = 0 Then
        lineCount = cm.ProcCountLines("Worksheet_Activate", 0)
        cm.DeleteLines startLine, lineCount
    End If
    Err.Clear
    code = "Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)" & vbCrLf & _
           "    If (Target.Range.Column = 21 Or Target.Range.Column = 22) And Target.Range.Row >= 5 And Target.Range.Row <= 120 Then" & vbCrLf & _
           "        ModuloComprovantes.AbrirRegistrosDaObservacao Target.Range" & vbCrLf & _
           "    End If" & vbCrLf & _
           "End Sub" & vbCrLf & _
           "Private Sub Worksheet_Activate()" & vbCrLf & _
           "    ModuloComprovantes.CongelarReferenciasDespesas" & vbCrLf & _
           "End Sub"
    cm.AddFromString code
    On Error Resume Next
End Sub

Sub AtualizarEventoControle(proj)
    On Error Resume Next
    Dim comp, cm, code, startLine, lineCount
    Set comp = ComponenteDaAba(proj, "Controle")
    If comp Is Nothing Then Exit Sub
    Set cm = comp.CodeModule
    startLine = cm.ProcStartLine("Worksheet_BeforeRightClick", 0)
    If Err.Number = 0 Then
        lineCount = cm.ProcCountLines("Worksheet_BeforeRightClick", 0)
        cm.DeleteLines startLine, lineCount
    End If
    Err.Clear
    code = "Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)" & vbCrLf & _
           "    If Target.Row >= 4 And Target.Column <= 9 Then" & vbCrLf & _
           "        If Trim(CStr(Me.Cells(Target.Row, 1).Value)) <> """" And LCase(Trim(CStr(Me.Cells(Target.Row, 1).Value))) <> ""total filtrado"" Then" & vbCrLf & _
           "            Cancel = True" & vbCrLf & _
           "            Me.Cells(Target.Row, 1).Select" & vbCrLf & _
           "            ModuloComprovantes.ExcluirRegistroControleSelecionado" & vbCrLf & _
           "        End If" & vbCrLf & _
           "    End If" & vbCrLf & _
           "End Sub"
    cm.AddFromString code
    On Error Resume Next
End Sub

Function ComponenteDaAba(proj, nomeAba)
    Dim sh
    For Each sh In wb.Worksheets
        If sh.Name = nomeAba Then
            Set ComponenteDaAba = proj.VBComponents(sh.CodeName)
            Exit Function
        End If
    Next
    Set ComponenteDaAba = Nothing
End Function
Sub RemoveAllStdModules(proj)
    On Error Resume Next
    Dim i, comp
    For i = proj.VBComponents.Count To 1 Step -1
        Set comp = proj.VBComponents(i)
        If comp.Type = vbext_ct_StdModule Then proj.VBComponents.Remove comp
        Err.Clear
    Next
    On Error Resume Next
End Sub
Sub RemoveComponent(proj, nome)
    On Error Resume Next
    Dim comp
    Set comp = proj.VBComponents(nome)
    If Err.Number = 0 Then proj.VBComponents.Remove comp
    Err.Clear
    On Error Resume Next
End Sub

Sub Log(msg)
    logFile.WriteLine Now & " - " & msg
End Sub

Sub Check(etapa)
    If Err.Number <> 0 Then
        Log "ERRO em " & etapa & ": " & Err.Number & " - " & Err.Description
        On Error Resume Next
        If Not wb Is Nothing Then wb.Close False
        If Not xl Is Nothing Then xl.Quit
        logFile.Close
        WScript.Quit 1
    Else
        Log "OK: " & etapa
    End If
    Err.Clear
End Sub



