Option Explicit
Const msoAutomationSecurityLow = 1
Const msoShapeRoundedRectangle = 5
Const msoShapeRectangle = 1
Const msoTextOrientationHorizontal = 1
Const msoTrue = -1
Const msoFalse = 0
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
WScript.Echo "Import redesenhada com botoes modernos."

Sub Montar(ws, wbName)
    Dim i, macroPrefix
    macroPrefix = "'" & wbName & "'!"
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next
    ws.Cells.Clear
    ws.Activate
    xl.ActiveWindow.DisplayGridlines = False
    xl.ActiveWindow.Zoom = 90
    ws.Range("A1:L34").Interior.Color = RGB(15, 23, 42)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:D").ColumnWidth = 21
    ws.Columns("E").ColumnWidth = 3
    ws.Columns("F:H").ColumnWidth = 21
    ws.Columns("I").ColumnWidth = 3
    ws.Columns("J:L").ColumnWidth = 19
    ws.Rows("1:34").RowHeight = 22

    With ws.Range("B2:H2")
        .Merge
        .Value = "Central de Importacao"
        .Font.Bold = True
        .Font.Size = 26
        .Font.Color = RGB(248, 250, 252)
    End With
    With ws.Range("B3:H3")
        .Merge
        .Value = "Comprovantes, revisao e registros em um fluxo unico."
        .Font.Size = 11
        .Font.Color = RGB(148, 163, 184)
    End With

    Panel ws, 48, 96, 360, 285, "1", "Preparar fila", "Extraia comprovantes ou gere o JSON completo.", RGB(37, 99, 235)
    Panel ws, 440, 96, 410, 285, "2", "Revisar e importar", "Carregue a fila, confira e grave na planilha.", RGB(22, 163, 74)
    Panel ws, 48, 408, 360, 235, "3", "Manutencao", "Abra documentos e limpe dados quando precisar.", RGB(245, 158, 11)
    Panel ws, 440, 408, 410, 235, "4", "Navegacao", "Acesso rapido as abas de trabalho.", RGB(168, 85, 247)
    Panel ws, 882, 96, 260, 285, "", "Status do processo", "Fonte: comprovantes.json" & Chr(10) & "Escopo: fila carregada" & Chr(10) & "Fluxo: extrair > revisar > importar", RGB(14, 165, 233)

    Action ws, 88, 160, 280, 34, "Abrir ferramenta", macroPrefix & "AbrirFerramentaExtracao", RGB(30, 41, 59), RGB(226, 232, 240)
    Action ws, 88, 206, 280, 34, "Extrair tudo", macroPrefix & "ExtrairDados", RGB(37, 99, 235), RGB(255, 255, 255)
    Action ws, 88, 252, 280, 34, "Carregar fila JSON", macroPrefix & "CarregarFilaJson", RGB(14, 165, 233), RGB(255, 255, 255)

    Action ws, 482, 172, 300, 38, "Importar pendentes", macroPrefix & "ImportarTodosPendentes", RGB(22, 163, 74), RGB(255, 255, 255)
    Action ws, 482, 226, 300, 38, "Importar linha", macroPrefix & "ImportarSelecionadoDaFila", RGB(34, 197, 94), RGB(15, 23, 42)

    Action ws, 88, 496, 280, 36, "Abrir comprovante", macroPrefix & "AbrirComprovante", RGB(245, 158, 11), RGB(15, 23, 42)
    Action ws, 88, 548, 280, 36, "Limpar dados", macroPrefix & "LimparDados", RGB(127, 29, 29), RGB(254, 226, 226)

    LinkCell ws, "F21", "Abrir Fila", "'Fila'!A1"
    LinkCell ws, "G21", "Abrir Controle", "'Controle'!A1"
    LinkCell ws, "F23", "Abrir Despesas", "'Despesas'!A1"

    ws.Range("B29:H29").Merge
    ws.Range("B29").Value = "Na aba Despesas, clique em Observacoes para abrir Controle filtrado pelos numeros vinculados."
    ws.Range("B29").Font.Color = RGB(148, 163, 184)
    ws.Range("B29").Font.Italic = True
End Sub

Sub Panel(ws, x, y, w, h, stepNo, title, body, accent)
    Dim sh, text
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    sh.Fill.ForeColor.RGB = RGB(30, 41, 59)
    sh.Line.ForeColor.RGB = RGB(51, 65, 85)
    text = title & Chr(10) & body
    If stepNo <> "" Then text = stepNo & "  " & text
    sh.TextFrame.Characters.Text = text
    sh.TextFrame.MarginLeft = 14
    sh.TextFrame.MarginTop = 12
    sh.TextFrame.VerticalAlignment = xlTop
    sh.TextFrame.Characters.Font.Color = RGB(226, 232, 240)
    sh.TextFrame.Characters.Font.Size = 11
    sh.TextFrame.Characters(1, Len(title) + Len(stepNo) + 2).Font.Bold = True
    sh.TextFrame.Characters(1, Len(title) + Len(stepNo) + 2).Font.Size = 14
End Sub

Sub Action(ws, x, y, w, h, caption, macroName, fillColor, fontColor)
    Dim sh
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    sh.Fill.ForeColor.RGB = fillColor
    sh.Line.ForeColor.RGB = RGB(71, 85, 105)
    sh.TextFrame.Characters.Text = caption
    sh.TextFrame.HorizontalAlignment = xlCenter
    sh.TextFrame.VerticalAlignment = xlCenter
    sh.TextFrame.Characters.Font.Bold = True
    sh.TextFrame.Characters.Font.Size = 10
    sh.TextFrame.Characters.Font.Color = fontColor
    sh.OnAction = macroName
End Sub

Sub LinkCell(ws, addr, label, target)
    ws.Hyperlinks.Add ws.Range(addr), "", target, "", label
    ws.Range(addr).Font.Bold = True
    ws.Range(addr).Font.Color = RGB(125, 211, 252)
    ws.Range(addr).Interior.Color = RGB(30, 41, 59)
End Sub
