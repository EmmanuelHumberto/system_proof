Option Explicit
Const msoAutomationSecurityLow = 1
Const msoShapeRoundedRectangle = 5
Const xlCenter = -4108

Dim xl, wb, ws, fso, base, arquivoXlsm, macroPrefix, i
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
arquivoXlsm = base & "\Planilha_Comprovantes_VBA.xlsm"

On Error Resume Next
Set xl = CreateObject("Excel.Application")
Check "Criar Excel"
xl.Visible = False
xl.DisplayAlerts = False
xl.AutomationSecurity = msoAutomationSecurityLow
Set wb = xl.Workbooks.Open(arquivoXlsm, False, False)
Check "Abrir XLSM"
Set ws = wb.Worksheets("Import")
Check "Abrir aba Import"
macroPrefix = "'" & wb.Name & "'!"

On Error Resume Next
ws.Unprotect "daisy2026"
If Err.Number <> 0 Then Err.Clear: ws.Unprotect ""
On Error GoTo 0
Check "Desproteger Import"

For i = ws.Shapes.Count To 1 Step -1
    ws.Shapes(i).Delete
Next
Check "Limpar botoes antigos"

ws.Range("A1:N28").Interior.Color = RGB(17, 24, 39)
ws.Range("B2:H2").Value = ""
ws.Range("B2:H2").Merge
ws.Range("B2").Value = "Comprovantes"
ws.Range("B2").Font.Name = "Segoe UI"
ws.Range("B2").Font.Size = 18
ws.Range("B2").Font.Bold = True
ws.Range("B2").Font.Color = RGB(243, 244, 246)
ws.Range("B3:H3").Value = ""
ws.Range("B3:H3").Merge
ws.Range("B3").Value = "Importacao e revisao"
ws.Range("B3").Font.Name = "Segoe UI"
ws.Range("B3").Font.Size = 9
ws.Range("B3").Font.Color = RGB(156, 163, 175)

SectionTitle ws, 48, 88, "Operacoes"
Action ws, 48, 120, 162, 30, "Ferramenta", macroPrefix & "AbrirFerramentaExtracao", RGB(31, 41, 55), RGB(229, 231, 235)
Action ws, 224, 120, 162, 30, "Extrair", macroPrefix & "ExtrairDados", RGB(37, 99, 235), RGB(255, 255, 255)
Action ws, 400, 120, 162, 30, "Carregar JSON", macroPrefix & "CarregarFilaJson", RGB(14, 116, 144), RGB(255, 255, 255)

SectionTitle ws, 48, 182, "Importacao"
Action ws, 48, 214, 250, 32, "Importar pendentes", macroPrefix & "ImportarTodosPendentes", RGB(22, 101, 52), RGB(255, 255, 255)
Action ws, 312, 214, 250, 32, "Importar linha", macroPrefix & "ImportarSelecionadoDaFila", RGB(21, 128, 61), RGB(255, 255, 255)

SectionTitle ws, 48, 276, "Manutencao"
Action ws, 48, 308, 162, 30, "Editar registro", macroPrefix & "AbrirFormularioEdicao", RGB(67, 56, 202), RGB(255, 255, 255)
Action ws, 224, 308, 162, 30, "Abrir arquivo", macroPrefix & "AbrirComprovante", RGB(180, 83, 9), RGB(255, 251, 235)
Action ws, 400, 308, 162, 30, "Limpar dados", macroPrefix & "LimparDados", RGB(127, 29, 29), RGB(254, 226, 226)
Action ws, 48, 360, 162, 26, "Destravar edicao", macroPrefix & "DesbloquearEdicao", RGB(31, 41, 55), RGB(209, 213, 219)
Action ws, 224, 360, 162, 26, "Travar edicao", macroPrefix & "BloquearEdicao", RGB(31, 41, 55), RGB(209, 213, 219)

ws.Cells.Locked = True
ws.Protect "daisy2026", True, True, True, True
ws.EnableSelection = 1
ws.Activate
ws.Range("A1").Select
xl.ActiveWindow.FreezePanes = False
xl.ActiveWindow.Split = False
xl.ActiveWindow.ScrollRow = 1
xl.ActiveWindow.ScrollColumn = 1
xl.ActiveWindow.DisplayGridlines = False
xl.ActiveWindow.Zoom = 100
wb.Save
Check "Salvar"
FecharExcel
WScript.Echo "OK: layout Import corrigido."
WScript.Quit 0

Sub SectionTitle(ws, x, y, caption)
    Dim sh
    Set sh = ws.Shapes.AddTextbox(1, x, y, 220, 18)
    sh.Fill.Visible = False
    sh.Line.Visible = False
    sh.TextFrame.Characters.Text = caption
    sh.TextFrame.Characters.Font.Name = "Segoe UI"
    sh.TextFrame.Characters.Font.Size = 9
    sh.TextFrame.Characters.Font.Bold = True
    sh.TextFrame.Characters.Font.Color = RGB(156, 163, 175)
End Sub

Sub Action(ws, x, y, w, h, caption, macroName, fillColor, fontColor)
    Dim sh
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    sh.Fill.ForeColor.RGB = fillColor
    sh.Line.ForeColor.RGB = RGB(55, 65, 81)
    sh.TextFrame.Characters.Text = caption
    sh.TextFrame.HorizontalAlignment = xlCenter
    sh.TextFrame.VerticalAlignment = xlCenter
    sh.TextFrame.Characters.Font.Name = "Segoe UI"
    sh.TextFrame.Characters.Font.Bold = True
    sh.TextFrame.Characters.Font.Size = 9
    sh.TextFrame.Characters.Font.Color = fontColor
    sh.OnAction = macroName
End Sub

Sub StatusBlock(ws, x, y)
    Dim sh
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, 406, 92)
    sh.Fill.ForeColor.RGB = RGB(24, 31, 46)
    sh.Line.ForeColor.RGB = RGB(55, 65, 81)
    sh.TextFrame.MarginLeft = 12
    sh.TextFrame.MarginTop = 8
    sh.TextFrame.Characters.Text = "Status" & Chr(10) & "JSON: comprovantes.json" & Chr(10) & "Pasta: comprovantes"
    sh.TextFrame.Characters.Font.Name = "Segoe UI"
    sh.TextFrame.Characters.Font.Size = 9
    sh.TextFrame.Characters.Font.Color = RGB(209, 213, 219)
End Sub

Sub QuickLink(ws, x, y, w, h, caption, target)
    Dim sh
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    sh.Fill.ForeColor.RGB = RGB(31, 41, 55)
    sh.Line.ForeColor.RGB = RGB(55, 65, 81)
    sh.TextFrame.Characters.Text = caption
    sh.TextFrame.HorizontalAlignment = xlCenter
    sh.TextFrame.VerticalAlignment = xlCenter
    sh.TextFrame.Characters.Font.Name = "Segoe UI"
    sh.TextFrame.Characters.Font.Size = 9
    sh.TextFrame.Characters.Font.Bold = True
    sh.TextFrame.Characters.Font.Color = RGB(125, 211, 252)
    ws.Hyperlinks.Add sh, "", target, "", caption
End Sub

Sub FecharExcel()
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set wb = Nothing
    Set xl = Nothing
End Sub

Sub Check(etapa)
    If Err.Number <> 0 Then
        Dim msg
        msg = "ERRO em " & etapa & ": " & Err.Number & " - " & Err.Description
        Err.Clear
        FecharExcel
        WScript.Echo msg
        WScript.Quit 1
    End If
End Sub
