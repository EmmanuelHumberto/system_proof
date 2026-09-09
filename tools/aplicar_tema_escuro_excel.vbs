Option Explicit

Const msoAutomationSecurityLow = 1
Const xlOpenXMLWorkbookMacroEnabled = 52
Const xlCenter = -4108
Const xlLeft = -4131
Const xlTop = -4160
Const xlContinuous = 1
Const xlThin = 2

Dim fso, base, arquivoXlsm, xl, wb, logFile
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
arquivoXlsm = base & "\Planilha_Comprovantes_VBA.xlsm"
Set logFile = fso.OpenTextFile(base & "\tools\aplicar_tema_escuro_excel.log", 2, True, -1)

On Error Resume Next
Log "Inicio"
Set xl = CreateObject("Excel.Application")
Check "CreateObject Excel.Application"
xl.Visible = False
xl.DisplayAlerts = False
xl.AutomationSecurity = msoAutomationSecurityLow
Set wb = xl.Workbooks.Open(arquivoXlsm, False, False)
Check "Abrir workbook"

MontarImport wb.Worksheets("Import"), wb.Name
EstilizarTabela wb.Worksheets("Fila"), "A1:N1", "Fila de revisao", RGB(34, 197, 94)
EstilizarTabela wb.Worksheets("Controle"), "A3:I3", "Registros importados", RGB(59, 130, 246)
EstilizarDespesas wb.Worksheets("Despesas")
EstilizarSimples wb.Worksheets("Resumo e Rateio"), RGB(168, 85, 247)
EstilizarSimples wb.Worksheets("Config"), RGB(245, 158, 11)

wb.Worksheets("Import").Activate
Err.Clear
wb.Save
Check "Salvar workbook"
wb.Close False
xl.Quit
Check "Fechar Excel"
Log "Fim OK"
logFile.Close
WScript.Echo "Tema escuro aplicado em: " & arquivoXlsm

Sub MontarImport(ws, wbName)
    Dim macroPrefix
    macroPrefix = "'" & wbName & "'!"
    ws.Cells.Clear
    LimparBotoes ws
    ws.Activate
    xl.ActiveWindow.DisplayGridlines = False
    xl.ActiveWindow.Zoom = 90

    ws.Range("A1:L34").Interior.Color = RGB(15, 23, 42)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:D").ColumnWidth = 21
    ws.Columns("E").ColumnWidth = 3
    ws.Columns("F:H").ColumnWidth = 21
    ws.Columns("I").ColumnWidth = 3
    ws.Columns("J:L").ColumnWidth = 18
    ws.Rows("1:34").RowHeight = 22

    With ws.Range("B2:H2")
        .Merge
        .Value = "Central de Importacao"
        .Font.Bold = True
        .Font.Size = 26
        .Font.Color = RGB(248, 250, 252)
        .Interior.Color = RGB(15, 23, 42)
    End With
    With ws.Range("B3:H3")
        .Merge
        .Value = "Comprovantes, revisao e registros em um fluxo unico."
        .Font.Size = 11
        .Font.Color = RGB(148, 163, 184)
    End With

    AddPanel ws, "B5:D15", "1", "Preparar fila", "Extraia comprovantes ou gere o JSON completo.", RGB(37, 99, 235)
    AddPanel ws, "F5:H15", "2", "Revisar e importar", "Carregue a fila, confira e grave na planilha.", RGB(22, 163, 74)
    AddPanel ws, "B17:D26", "3", "Manutencao", "Abra documentos e limpe dados quando precisar.", RGB(245, 158, 11)
    AddPanel ws, "F17:H26", "4", "Navegacao", "Acesso rapido as abas de trabalho.", RGB(168, 85, 247)

    AddButton ws, 88, 150, 250, 34, "Abrir ferramenta", macroPrefix & "AbrirFerramentaExtracao", RGB(30, 41, 59), RGB(226, 232, 240)
    AddButton ws, 88, 195, 250, 34, "Extrair tudo", macroPrefix & "ExtrairDados", RGB(37, 99, 235), RGB(255, 255, 255)
    AddButton ws, 88, 240, 250, 34, "Carregar fila JSON", macroPrefix & "CarregarFilaJson", RGB(14, 165, 233), RGB(255, 255, 255)

    AddButton ws, 426, 162, 250, 38, "Importar pendentes", macroPrefix & "ImportarTodosPendentes", RGB(22, 163, 74), RGB(255, 255, 255)
    AddButton ws, 426, 212, 250, 38, "Importar linha", macroPrefix & "ImportarSelecionadoDaFila", RGB(34, 197, 94), RGB(15, 23, 42)

    AddButton ws, 88, 432, 250, 36, "Abrir comprovante", macroPrefix & "AbrirComprovante", RGB(245, 158, 11), RGB(15, 23, 42)
    AddButton ws, 88, 480, 250, 36, "Limpar dados", macroPrefix & "LimparDados", RGB(127, 29, 29), RGB(254, 226, 226)

    With ws.Range("J5:L18")
        .Interior.Color = RGB(30, 41, 59)
        .Borders.Color = RGB(51, 65, 85)
        .Font.Color = RGB(226, 232, 240)
    End With
    ws.Range("J5:L5").Merge
    ws.Range("J5").Value = "Status do processo"
    ws.Range("J5").Font.Bold = True
    ws.Range("J5").Font.Size = 13
    ws.Range("J7").Value = "Fonte"
    ws.Range("K7:L7").Merge
    ws.Range("K7").Value = "comprovantes.json"
    ws.Range("J9").Value = "Escopo"
    ws.Range("K9:L9").Merge
    ws.Range("K9").Value = "Fila carregada"
    ws.Range("J11").Value = "Fluxo"
    ws.Range("K11:L11").Merge
    ws.Range("K11").Value = "1 Extrair  >  2 Revisar  >  3 Importar"
    ws.Range("J13").Value = "Links"
    ws.Range("K13:L13").Merge
    ws.Range("K13").Value = "Observacoes filtram Controle"
    ws.Range("J7:J13").Font.Color = RGB(148, 163, 184)

    ws.Hyperlinks.Add ws.Range("F21"), "", "'Fila'!A1", "", "Abrir Fila"
    ws.Hyperlinks.Add ws.Range("G21"), "", "'Controle'!A1", "", "Abrir Controle"
    ws.Hyperlinks.Add ws.Range("F23"), "", "'Despesas'!A1", "", "Abrir Despesas"
    With ws.Range("F21:G23")
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(125, 211, 252)
    End With

    ws.Range("B29:H29").Merge
    ws.Range("B29").Value = "Dica: na aba Despesas, clique em Observacoes para abrir Controle filtrado pelos numeros vinculados."
    ws.Range("B29").Font.Color = RGB(148, 163, 184)
    ws.Range("B29").Font.Italic = True
End Sub

Sub AddPanel(ws, addr, etapa, titulo, subtitulo, cor)
    Dim rng
    Set rng = ws.Range(addr)
    rng.Merge
    rng.Interior.Color = RGB(30, 41, 59)
    rng.Borders.Color = RGB(51, 65, 85)
    rng.Borders.Weight = xlThin
    rng.VerticalAlignment = xlTop
    rng.HorizontalAlignment = xlLeft
    rng.Cells(1, 1).Value = etapa & "  " & titulo & Chr(10) & subtitulo
    rng.Cells(1, 1).Font.Color = RGB(226, 232, 240)
    rng.Cells(1, 1).Font.Size = 11
    rng.Cells(1, 1).Characters(1, Len(etapa)).Font.Color = RGB(255, 255, 255)
    rng.Cells(1, 1).Characters(1, Len(etapa)).Font.Bold = True
    rng.Cells(1, 1).Characters(4, Len(titulo)).Font.Bold = True
    rng.Cells(1, 1).Characters(4, Len(titulo)).Font.Size = 14
    rng.WrapText = True
End Sub

Sub AddButton(ws, leftPos, topPos, widthVal, heightVal, caption, macroName, fillColor, fontColor)
    Dim btn
    Set btn = ws.Buttons.Add(leftPos, topPos, widthVal, heightVal)
    btn.Caption = caption
    btn.OnAction = macroName
    btn.Font.Bold = True
    btn.Font.Size = 10
    btn.Font.Color = fontColor
    btn.Interior.Color = fillColor
End Sub

Sub EstilizarTabela(ws, headerAddr, titulo, cor)
    ws.Activate
    xl.ActiveWindow.DisplayGridlines = False
    ws.Cells.Interior.Color = RGB(15, 23, 42)
    ws.Cells.Font.Color = RGB(226, 232, 240)
    ws.Range(headerAddr).Interior.Color = RGB(30, 41, 59)
    ws.Range(headerAddr).Font.Color = RGB(248, 250, 252)
    ws.Range(headerAddr).Font.Bold = True
    ws.Range(headerAddr).Borders.Color = RGB(71, 85, 105)
    ws.Rows(1).RowHeight = 24
    ws.Columns.AutoFit
End Sub

Sub EstilizarDespesas(ws)
    ws.Activate
    xl.ActiveWindow.DisplayGridlines = False
    ws.Cells.Interior.Color = RGB(15, 23, 42)
    ws.Cells.Font.Color = RGB(226, 232, 240)
    ws.Range("A1:Q2").Interior.Color = RGB(15, 23, 42)
    ws.Range("A1:Q2").Font.Color = RGB(248, 250, 252)
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A4:Q4").Interior.Color = RGB(30, 41, 59)
    ws.Range("A4:Q4").Font.Color = RGB(248, 250, 252)
    ws.Range("A4:Q4").Font.Bold = True
    ws.Range("A5:Q59").Borders.Color = RGB(51, 65, 85)
    ws.Range("P5:Q57").Interior.Color = RGB(24, 36, 56)
    ws.Columns.AutoFit
End Sub

Sub EstilizarSimples(ws, cor)
    ws.Activate
    xl.ActiveWindow.DisplayGridlines = False
    ws.Cells.Interior.Color = RGB(15, 23, 42)
    ws.Cells.Font.Color = RGB(226, 232, 240)
    ws.Range("A1:Z4").Interior.Color = RGB(30, 41, 59)
    ws.Range("A1:Z4").Font.Color = RGB(248, 250, 252)
    ws.Range("A1:Z4").Font.Bold = True
    ws.Columns.AutoFit
End Sub

Sub LimparBotoes(ws)
    Dim i
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next
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



