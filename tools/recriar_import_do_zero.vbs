Option Explicit
Const msoAutomationSecurityLow = 1
Const msoShapeRoundedRectangle = 5
Const msoShapeRectangle = 1
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
Check "Abrir Import"
macroPrefix = "'" & wb.Name & "'!"

Desproteger ws

For i = ws.Shapes.Count To 1 Step -1
    ws.Shapes(i).Delete
Next
Check "Remover shapes"

ws.Cells.Clear
ws.Cells.EntireColumn.Hidden = False
ws.Cells.EntireRow.Hidden = False
Check "Limpar aba Import"
ws.Cells.Locked = True

ws.Range("A1:H24").Interior.Color = RGB(15, 23, 42)
ws.Range("I1:XFD1048576").Interior.ColorIndex = -4142
ws.Range("A25:XFD1048576").Interior.ColorIndex = -4142

ws.Columns("A").ColumnWidth = 3
For i = 2 To 7
    ws.Columns(i).ColumnWidth = 15
Next
ws.Columns("H").ColumnWidth = 3
ws.Columns("I:XFD").ColumnWidth = 8.43
ws.Rows("1:24").RowHeight = 20

ws.Range("B2:G2").Merge
ws.Range("B2").Value = "Central de comprovantes"
ws.Range("B2").Font.Name = "Segoe UI"
ws.Range("B2").Font.Size = 16
ws.Range("B2").Font.Bold = True
ws.Range("B2").Font.Color = RGB(248, 250, 252)
ws.Range("B3:G3").Merge
ws.Range("B3").Value = "Extracao, fila e revisao"
ws.Range("B3").Font.Name = "Segoe UI"
ws.Range("B3").Font.Size = 9
ws.Range("B3").Font.Color = RGB(148, 163, 184)
TopRule ws, 48, 76, 492

SectionTitle ws, 48, 104, "Preparar"
Action ws, 48, 132, 156, 30, "Ferramenta", macroPrefix & "AbrirFerramentaExtracao", RGB(30, 41, 59), RGB(226, 232, 240)
Action ws, 216, 132, 156, 30, "Extrair", macroPrefix & "ExtrairDados", RGB(30, 64, 175), RGB(255, 255, 255)
Action ws, 384, 132, 156, 30, "Carregar JSON", macroPrefix & "CarregarFilaJson", RGB(15, 82, 111), RGB(255, 255, 255)

SectionTitle ws, 48, 196, "Importar"
Action ws, 48, 224, 240, 32, "Importar pendentes", macroPrefix & "ImportarTodosPendentes", RGB(22, 101, 52), RGB(255, 255, 255)
Action ws, 300, 224, 240, 32, "Importar linha", macroPrefix & "ImportarSelecionadoDaFila", RGB(21, 94, 117), RGB(255, 255, 255)

SectionTitle ws, 48, 290, "Manutencao"
Action ws, 48, 318, 156, 30, "Editar registro", macroPrefix & "AbrirFormularioEdicao", RGB(51, 65, 85), RGB(241, 245, 249)
Action ws, 216, 318, 156, 30, "Abrir arquivo", macroPrefix & "AbrirComprovante", RGB(51, 65, 85), RGB(241, 245, 249)
Action ws, 384, 318, 156, 30, "Limpar dados", macroPrefix & "LimparDados", RGB(127, 29, 29), RGB(254, 226, 226)
Action ws, 48, 370, 156, 26, "Destravar edicao", macroPrefix & "DesbloquearEdicao", RGB(30, 41, 59), RGB(203, 213, 225)
Action ws, 216, 370, 156, 26, "Travar edicao", macroPrefix & "BloquearEdicao", RGB(30, 41, 59), RGB(203, 213, 225)

ws.ScrollArea = "A1:H24"
ws.Protect "daisy2026", True, True, True, True
ws.EnableSelection = 1
ws.Activate
ws.Range("A1").Select
With xl.ActiveWindow
    .FreezePanes = False
    .SplitColumn = 0
    .SplitRow = 0
    .ScrollRow = 1
    .ScrollColumn = 1
    .DisplayGridlines = False
    .Zoom = 100
End With

wb.Save
Check "Salvar"
FecharExcel
WScript.Echo "OK: aba Import recriada do zero."
WScript.Quit 0

Sub Desproteger(ws)
    On Error Resume Next
    ws.Unprotect "daisy2026"
    If Err.Number <> 0 Then Err.Clear: ws.Unprotect ""
    Err.Clear
    On Error GoTo 0
End Sub

Sub TopRule(ws, x, y, w)
    Dim sh
    Set sh = ws.Shapes.AddShape(msoShapeRectangle, x, y, w, 1.5)
    sh.Fill.ForeColor.RGB = RGB(51, 65, 85)
    sh.Line.Visible = False
    sh.Placement = 3
End Sub

Sub SectionTitle(ws, x, y, caption)
    Dim sh
    Set sh = ws.Shapes.AddTextbox(1, x, y, 220, 18)
    sh.Fill.Visible = False
    sh.Line.Visible = False
    sh.TextFrame.Characters.Text = caption
    sh.TextFrame.Characters.Font.Name = "Segoe UI"
    sh.TextFrame.Characters.Font.Size = 9
    sh.TextFrame.Characters.Font.Bold = True
    sh.TextFrame.Characters.Font.Color = RGB(148, 163, 184)
    sh.Placement = 3
End Sub

Sub Action(ws, x, y, w, h, caption, macroName, fillColor, fontColor)
    Dim sh
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    sh.Fill.ForeColor.RGB = fillColor
    sh.Line.ForeColor.RGB = RGB(71, 85, 105)
    sh.TextFrame.Characters.Text = caption
    sh.TextFrame.HorizontalAlignment = xlCenter
    sh.TextFrame.VerticalAlignment = xlCenter
    sh.TextFrame.Characters.Font.Name = "Segoe UI"
    sh.TextFrame.Characters.Font.Bold = True
    sh.TextFrame.Characters.Font.Size = 9
    sh.TextFrame.Characters.Font.Color = fontColor
    sh.OnAction = macroName
    sh.Placement = 3
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
