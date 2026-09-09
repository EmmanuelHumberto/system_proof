Option Explicit
Const msoAutomationSecurityLow = 1
Const msoShapeRoundedRectangle = 5
Const msoShapeRectangle = 1
Const xlCenter = -4108
Const xlLeft = -4131
Const xlTop = -4160

Dim xl, wb, ws, fso, base, arquivoXlsm
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
arquivoXlsm = base & "\Planilha_Comprovantes_VBA.xlsm"
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.AutomationSecurity = msoAutomationSecurityLow
Set wb = xl.Workbooks.Open(arquivoXlsm, False, False)
Set ws = wb.Worksheets("Import")
Montar ws, wb.Name
wb.Save
wb.Close False
xl.Quit
WScript.Echo "Aba Import remodelada: clean dark."

Sub Montar(ws, wbName)
    Dim i, macroPrefix
    macroPrefix = "'" & wbName & "'!"
    On Error Resume Next
    ws.Unprotect ""
    On Error GoTo 0
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next
    ws.Range("A1:N28").ClearContents
    ws.Range("A1:N28").ClearFormats
    ws.Activate
    xl.ActiveWindow.DisplayGridlines = False
    xl.ActiveWindow.Zoom = 100

    ws.Range("A1:N28").Interior.Color = RGB(17, 24, 39)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 4
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 18
    ws.Columns("I").ColumnWidth = 4
    ws.Columns("J").ColumnWidth = 17
    ws.Columns("K").ColumnWidth = 17
    ws.Columns("L").ColumnWidth = 17
    ws.Rows("1:28").RowHeight = 20

    With ws.Range("B2:H2")
        .Merge
        .Value = "Comprovantes"
        .Font.Name = "Segoe UI"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(243, 244, 246)
        .Interior.Color = RGB(17, 24, 39)
    End With
    With ws.Range("B3:H3")
        .Merge
        .Value = "Fila, revisao e registros"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Color = RGB(156, 163, 175)
    End With

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

    With ws.Range("B22:L22")
        .Merge
        .Value = ""
        .Interior.Color = RGB(17, 24, 39)
    End With
    ws.Cells.Locked = True
    ws.Protect "", True, True, True, True
    ws.EnableSelection = 1
End Sub

Sub SectionTitle(ws, x, y, caption)
    Dim sh
    Set sh = ws.Shapes.AddTextbox(1, x, y, 300, 22)
    sh.Fill.Visible = False
    sh.Line.Visible = False
    sh.TextFrame.Characters.Text = caption
    sh.TextFrame.Characters.Font.Name = "Segoe UI"
    sh.TextFrame.Characters.Font.Size = 10
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
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, 442, 96)
    sh.Fill.ForeColor.RGB = RGB(24, 31, 46)
    sh.Line.ForeColor.RGB = RGB(55, 65, 81)
    sh.TextFrame.MarginLeft = 14
    sh.TextFrame.MarginTop = 10
    sh.TextFrame.Characters.Text = "Status" & Chr(10) & "JSON: comprovantes.json" & Chr(10) & "Pasta: comprovantes"
    sh.TextFrame.Characters.Font.Name = "Segoe UI"
    sh.TextFrame.Characters.Font.Size = 10
    sh.TextFrame.Characters.Font.Color = RGB(209, 213, 219)
    sh.TextFrame.Characters(1, 6).Font.Bold = True
    sh.TextFrame.Characters(1, 6).Font.Color = RGB(243, 244, 246)
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
