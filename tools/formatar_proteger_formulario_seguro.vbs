Option Explicit
Const msoAutomationSecurityLow = 1
Const msoShapeRoundedRectangle = 5
Const xlCenter = -4108
Const xlUnlockedCells = 1

Dim xl, wb, fso, base, arquivoXlsm, arquivoBas
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
arquivoXlsm = base & "\Planilha_Comprovantes_VBA.xlsm"
arquivoBas = base & "\vba\modulos_vba.bas"

On Error Resume Next
Set xl = CreateObject("Excel.Application")
Check "Criar Excel"
xl.Visible = False
xl.DisplayAlerts = False
xl.AutomationSecurity = msoAutomationSecurityLow
Set wb = xl.Workbooks.Open(arquivoXlsm, False, False)
Check "Abrir XLSM"

DesprotegerTudo wb
Check "Desproteger"
ImportarModuloSeguro wb, arquivoBas
Check "Importar ModuloComprovantes"
AtualizarEventoDespesas wb
Check "Atualizar evento Despesas"
AtualizarEventoControle wb
Check "Atualizar evento Controle"
MontarEditar wb
Check "Montar Editar"
FormatarTudo wb
Check "Formatar"
DesprotegerTudo wb
Check "Deixar destravado"

wb.Save
Check "Salvar"
FecharExcel
WScript.Echo "OK: formatacao, VBA e formulario aplicados; planilha destravada."
WScript.Quit 0

Sub FecharExcel()
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then
        xl.DisplayAlerts = False
        xl.Quit
    End If
    Set wb = Nothing
    Set xl = Nothing
    On Error GoTo 0
End Sub

Sub Check(etapa)
    If Err.Number <> 0 Then
        Dim erroMsg
        erroMsg = "ERRO em " & etapa & ": " & Err.Number & " - " & Err.Description
        Err.Clear
        FecharExcel
        WScript.Echo erroMsg
        WScript.Quit 1
    End If
End Sub

Sub ImportarModuloSeguro(wb, arquivoBas)
    Dim proj, comp
    Set proj = wb.VBProject
    On Error Resume Next
    Set comp = proj.VBComponents("ModuloComprovantes")
    If Not comp Is Nothing Then proj.VBComponents.Remove comp
    On Error GoTo 0
    proj.VBComponents.Import arquivoBas
End Sub

Sub AtualizarEventoDespesas(wb)
    Dim proj, comp, cm, code, startLine, lineCount
    Set proj = wb.VBProject
    Set comp = ComponenteDaAba(wb, proj, "Despesas")
    If comp Is Nothing Then Exit Sub
    Set cm = comp.CodeModule

    On Error Resume Next
    Err.Clear
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
    On Error GoTo 0

    code = "Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)" & vbCrLf & _
           "    If (Target.Range.Column = 21 Or Target.Range.Column = 22) And Target.Range.Row >= 5 And Target.Range.Row <= 120 Then" & vbCrLf & _
           "        ModuloComprovantes.AbrirRegistrosDaObservacao Target.Range" & vbCrLf & _
           "    End If" & vbCrLf & _
           "End Sub" & vbCrLf & _
           "Private Sub Worksheet_Activate()" & vbCrLf & _
           "    ModuloComprovantes.CongelarReferenciasDespesas" & vbCrLf & _
           "End Sub"
    cm.AddFromString code
End Sub

Sub AtualizarEventoControle(wb)
    Dim proj, comp, cm, code, startLine, lineCount
    Set proj = wb.VBProject
    Set comp = ComponenteDaAba(wb, proj, "Controle")
    If comp Is Nothing Then Exit Sub
    Set cm = comp.CodeModule

    On Error Resume Next
    Err.Clear
    startLine = cm.ProcStartLine("Worksheet_BeforeRightClick", 0)
    If Err.Number = 0 Then
        lineCount = cm.ProcCountLines("Worksheet_BeforeRightClick", 0)
        cm.DeleteLines startLine, lineCount
    End If
    Err.Clear
    On Error GoTo 0

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
End Sub

Function ComponenteDaAba(wb, proj, nomeAba)
    Dim sh
    For Each sh In wb.Worksheets
        If sh.Name = nomeAba Then
            Set ComponenteDaAba = proj.VBComponents(sh.CodeName)
            Exit Function
        End If
    Next
    Set ComponenteDaAba = Nothing
End Function

Sub DesprotegerTudo(wb)
    Dim ws
    For Each ws In wb.Worksheets
        DesprotegerAba ws
        On Error Resume Next
        ws.EnableSelection = 0
        On Error GoTo 0
    Next
End Sub

Sub DesprotegerAba(ws)
    On Error Resume Next
    ws.Unprotect "daisy2026"
    If Err.Number <> 0 Then
        Err.Clear
        ws.Unprotect ""
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Sub MontarEditar(wb)
    Dim ws, labels, i, macroPrefix
    macroPrefix = "'" & wb.Name & "'!"
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("Editar")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Editar"
    End If
    DesprotegerAba ws

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next
    ws.Cells.Clear
    ws.Cells.Locked = True
    ws.Range("A1:D23").Interior.Color = RGB(17, 24, 39)
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 26
    ws.Columns("E").ColumnWidth = 2
    ws.Rows("1:24").RowHeight = 20

    ws.Range("B2:D2").Merge
    ws.Range("B2").Value = "Editar registro"
    ws.Range("B2").Font.Name = "Segoe UI"
    ws.Range("B2").Font.Size = 18
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(243, 244, 246)

    ws.Range("B4").Value = "Origem"
    ws.Range("D4").Value = "Linha"
    ws.Range("B5").Value = "Fila"
    ws.Range("D5").Value = 2
    PrepararCampo ws.Range("B5")
    PrepararCampo ws.Range("D5")
    PrepararCampo ws.Range("D8:D18")

    labels = Array("Categoria", "Despesa / Recebedor", "Data", "Competencia", "Valor", "Pagador", "Recebedor", "Tipo / Observacao", "Periodicidade", "Arquivo", "Cota parte?")
    For i = 0 To UBound(labels)
        ws.Cells(8 + i, 2).Value = labels(i)
    Next
    ws.Range("B4:E18").Font.Name = "Segoe UI"
    ws.Range("B4:E18").Font.Size = 9
    ws.Range("B4:B18").Font.Color = RGB(156, 163, 175)
    ws.Range("B4:B18").Font.Bold = True
    ws.Range("D4").Font.Color = RGB(156, 163, 175)
    ws.Range("D4").Font.Bold = True
    ws.Range("D10").NumberFormat = "dd/mm/yyyy"
    ws.Range("D11").NumberFormat = "yyyy-mm"
    ws.Range("D12").NumberFormat = "R$ #,##0.00"

    AddButton ws, 58, 396, 122, 28, "Carregar", macroPrefix & "CarregarRegistroEdicao", RGB(14, 116, 144), RGB(255, 255, 255)
    AddButton ws, 192, 396, 122, 28, "Salvar", macroPrefix & "SalvarFormularioEdicao", RGB(22, 101, 52), RGB(255, 255, 255)
End Sub

Sub AddButton(ws, x, y, w, h, caption, macroName, fillColor, fontColor)
    Dim sh
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    sh.Fill.ForeColor.RGB = fillColor
    sh.Line.ForeColor.RGB = RGB(55, 65, 81)
    sh.TextFrame.Characters.Text = caption
    sh.TextFrame.HorizontalAlignment = xlCenter
    sh.TextFrame.VerticalAlignment = xlCenter
    sh.TextFrame.Characters.Font.Name = "Segoe UI"
    sh.TextFrame.Characters.Font.Size = 9
    sh.TextFrame.Characters.Font.Bold = True
    sh.TextFrame.Characters.Font.Color = fontColor
    sh.OnAction = macroName
End Sub

Sub PrepararCampo(rng)
    rng.Locked = False
    rng.Interior.Color = RGB(31, 41, 55)
    rng.Font.Color = RGB(243, 244, 246)
End Sub

Sub FormatarTudo(wb)
    Dim ws, colMedia
    Set ws = wb.Worksheets("Despesas")
    colMedia = HeaderCol(ws, 4, "Media mensal")
    If colMedia = 0 Then colMedia = HeaderCol(ws, 4, "Média mensal")
    ws.Range("C5:N122").NumberFormatLocal = "R$ #.##0,00"
    If colMedia > 0 Then ws.Range(ws.Cells(5, 3), ws.Cells(122, colMedia)).NumberFormatLocal = "R$ #.##0,00"
    ws.Range("O5:O120").NumberFormat = "General"

    Set ws = wb.Worksheets("Controle")
    ws.Columns(1).NumberFormat = "0"
    ws.Columns(4).NumberFormat = "yyyy-mm"
    ws.Columns(5).NumberFormatLocal = "R$ #.##0,00"

    Set ws = wb.Worksheets("Fila")
    ws.Columns(1).NumberFormat = "0"
    ws.Columns(5).NumberFormat = "dd/mm/yyyy"
    ws.Columns(6).NumberFormat = "yyyy-mm"
    ws.Columns(7).NumberFormatLocal = "R$ #.##0,00"
End Sub

Function HeaderCol(ws, rowNum, prefix)
    Dim c, txt, alvo
    alvo = LCase(RemoveAcento(prefix))
    For c = 1 To ws.Cells(rowNum, ws.Columns.Count).End(-4159).Column
        txt = LCase(RemoveAcento(CStr(ws.Cells(rowNum, c).Value)))
        If Left(txt, Len(alvo)) = alvo Then
            HeaderCol = c
            Exit Function
        End If
    Next
    HeaderCol = 0
End Function

Function RemoveAcento(s)
    s = LCase(s)
    s = Replace(s, "á", "a"): s = Replace(s, "à", "a"): s = Replace(s, "ã", "a"): s = Replace(s, "â", "a")
    s = Replace(s, "é", "e"): s = Replace(s, "ê", "e")
    s = Replace(s, "í", "i")
    s = Replace(s, "ó", "o"): s = Replace(s, "õ", "o"): s = Replace(s, "ô", "o")
    s = Replace(s, "ú", "u"): s = Replace(s, "ç", "c")
    RemoveAcento = s
End Function

Sub ProtegerTudo(wb)
    Dim ws
    For Each ws In wb.Worksheets
        DesprotegerAba ws
        ws.Cells.Locked = True
        If ws.Name = "Editar" Then
            ws.Range("B5").Locked = False
            ws.Range("D5").Locked = False
            ws.Range("D8:D18").Locked = False
        End If
        If ws.Name = "Despesas" Then ws.Range("R5:R120").Locked = False
        ws.Protect "daisy2026", True, True, True, True
        ws.EnableSelection = xlUnlockedCells
    Next
End Sub

