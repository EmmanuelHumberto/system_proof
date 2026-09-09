Option Explicit

Const msoAutomationSecurityLow = 1
Const xlOpenXMLWorkbookMacroEnabled = 52

Dim fso, base, arquivoXlsm, logFile
Dim xl, wb, ws, arquivoTemp
Dim tentativa

Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
arquivoXlsm = base & "\Planilha_Comprovantes_VBA.xlsm"
arquivoTemp = base & "\Planilha_Comprovantes_VBA_atualizado.xlsm"
Set logFile = fso.OpenTextFile(base & "\tools\criar_botoes_excel.log", 2, True, -1)

On Error Resume Next
Log "Inicio"

Set xl = CreateObject("Excel.Application")
Check "CreateObject Excel.Application"
xl.Visible = False
xl.DisplayAlerts = False
xl.AutomationSecurity = msoAutomationSecurityLow

Set wb = xl.Workbooks.Open(arquivoXlsm, False, False)
Check "Abrir workbook"

Set ws = wb.Worksheets("Import")
Check "Selecionar aba Import"

LimparBotoes ws
MontarTela ws, wb.Name

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
WScript.Echo "Botoes criados com sucesso em: " & arquivoXlsm

Sub MontarTela(ws, wbName)
    Dim macroPrefix
    macroPrefix = "'" & wbName & "'!"

    ws.Cells.Clear
    ws.Activate
    xl.ActiveWindow.DisplayGridlines = False
    xl.ActiveWindow.Zoom = 90

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 4
    ws.Columns("F").ColumnWidth = 22
    ws.Columns("G").ColumnWidth = 22
    ws.Columns("H").ColumnWidth = 22
    ws.Columns("I").ColumnWidth = 4
    ws.Columns("J").ColumnWidth = 24
    ws.Rows("1:30").RowHeight = 22

    ws.Range("A1:J30").Interior.Color = RGB(246, 248, 250)
    ws.Range("B2:H2").Merge
    ws.Range("B2").Value = "Central de importacao de comprovantes"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 22
    ws.Range("B2").Font.Color = RGB(31, 41, 55)

    ws.Range("B3:H3").Merge
    ws.Range("B3").Value = "Extraia os dados, carregue a fila JSON, revise e importe para a planilha principal."
    ws.Range("B3").Font.Size = 11
    ws.Range("B3").Font.Color = RGB(75, 85, 99)

    AddPanel ws, "B5:D15", "1. Preparar fila", "Abra a ferramenta de extracao ou gere o JSON direto."
    AddPanel ws, "F5:H15", "2. Revisar e importar", "Carrega a fila para conferencia antes de gravar."
    AddPanel ws, "B17:D25", "3. Manutencao", "Acoes auxiliares para a planilha e comprovantes."
    AddPanel ws, "F17:H25", "Atalhos", "Navegue pelas abas de trabalho."

    AddButton ws, 72, 138, 260, 36, "Abrir Ferramenta de Extracao", macroPrefix & "AbrirFerramentaExtracao"
    AddButton ws, 72, 184, 260, 36, "Extrair / Enfileirar JSON", macroPrefix & "ExtrairDados"
    AddButton ws, 72, 230, 260, 36, "Carregar Fila JSON", macroPrefix & "CarregarFilaJson"

    AddButton ws, 410, 150, 260, 38, "Importar Todos Pendentes", macroPrefix & "ImportarTodosPendentes"
    AddButton ws, 410, 200, 260, 38, "Importar Linha Selecionada", macroPrefix & "ImportarSelecionadoDaFila"

    AddButton ws, 72, 420, 260, 38, "Abrir Comprovante", macroPrefix & "AbrirComprovante"
    AddButton ws, 72, 470, 260, 38, "Limpar Dados Importados", macroPrefix & "LimparDados"

    ws.Hyperlinks.Add ws.Range("F20"), "", "'Fila'!A1", "", "Abrir Fila"
    ws.Hyperlinks.Add ws.Range("G20"), "", "'Controle'!A1", "", "Abrir Controle"
    ws.Hyperlinks.Add ws.Range("F22"), "", "'Despesas'!A1", "", "Abrir Despesas"
    ws.Range("F20:G22").Font.Bold = True
    ws.Range("F20:G22").Font.Size = 12

    ws.Range("J5").Value = "Status"
    ws.Range("J5").Font.Bold = True
    ws.Range("J6").Value = "JSON:"
    ws.Range("J7").Value = "comprovantes.json"
    ws.Range("J9").Value = "Pasta:"
    ws.Range("J10").Value = "comprovantes/"
    ws.Range("J12").Value = "Fluxo:"
    ws.Range("J13").Value = "1 Extrair"
    ws.Range("J14").Value = "2 Carregar fila"
    ws.Range("J15").Value = "3 Revisar"
    ws.Range("J16").Value = "4 Importar"
    ws.Range("J5:J16").Font.Color = RGB(55, 65, 81)

    ws.Range("B27:H27").Merge
    ws.Range("B27").Value = "Use a aba Fila para selecionar uma linha antes de usar Importar Linha Selecionada."
    ws.Range("B27").Font.Color = RGB(107, 114, 128)
    ws.Range("B27").Font.Italic = True
End Sub

Sub AddPanel(ws, addr, titulo, subtitulo)
    Dim rng
    Set rng = ws.Range(addr)
    rng.Merge
    rng.Interior.Color = RGB(255, 255, 255)
    rng.Borders.Color = RGB(209, 213, 219)
    rng.Borders.Weight = 2
    rng.VerticalAlignment = -4160
    rng.HorizontalAlignment = -4131

    rng.Cells(1, 1).Value = titulo & Chr(10) & subtitulo
    rng.Cells(1, 1).Font.Size = 11
    rng.Cells(1, 1).Font.Color = RGB(75, 85, 99)
    rng.Cells(1, 1).Characters(1, Len(titulo)).Font.Bold = True
    rng.Cells(1, 1).Characters(1, Len(titulo)).Font.Size = 14
    rng.WrapText = True
End Sub

Sub AddButton(ws, leftPos, topPos, widthVal, heightVal, caption, macroName)
    Dim btn
    Set btn = ws.Buttons.Add(leftPos, topPos, widthVal, heightVal)
    btn.Caption = caption
    btn.OnAction = macroName
    btn.Font.Bold = True
    btn.Font.Size = 10
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
