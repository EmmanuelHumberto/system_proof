Option Explicit

Const msoShapeRoundedRectangle = 5
Const msoShapeRectangle = 1
Const msoTextOrientationHorizontal = 1
Const xlCenter = -4108
Const xlLeft = -4131
Const xlUnlockedCells = 1
Const xlNoRestrictions = 0
Const xlNone = -4142
Const xlContinuous = 1
Const xlThin = 2
Const xlEdgeLeft = 7
Const xlEdgeTop = 8
Const xlEdgeBottom = 9
Const xlEdgeRight = 10
Const xlInsideVertical = 11
Const xlInsideHorizontal = 12

Dim fso, xl, wb, ws, base, arquivoXlsm, macroPrefix
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
arquivoXlsm = base & "\Planilha_Comprovantes_VBA.xlsm"

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False
Set wb = xl.Workbooks.Open(arquivoXlsm, False, False)
macroPrefix = "'" & wb.Name & "'!"

DesprotegerTudo
GarantirEstruturaDespesas
AplicarFundoGeral
FormatarDespesas
CongelarReferenciasDespesas
FormatarControle
FormatarFila
FormatarResumo
FormatarConfig
FormatarCadastros
RedesenharImport
OcultarAbaEditar
DeixarDestravado
NormalizarJanela "Import"

wb.Save
wb.Close False
xl.Quit
WScript.Echo "OK: visual pastel aplicado, resumo corrigido e planilha destravada."

Sub DesprotegerTudo()
    Dim sh
    For Each sh In wb.Worksheets
        DesprotegerAba sh
    Next
End Sub

Sub DesprotegerAba(sh)
    On Error Resume Next
    sh.Unprotect "daisy2026"
    sh.Unprotect ""
    sh.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub

Sub DeixarDestravado()
    Dim sh
    For Each sh In wb.Worksheets
        On Error Resume Next
        sh.Unprotect "daisy2026"
        sh.Unprotect ""
        sh.Cells.Locked = False
        sh.EnableSelection = xlNoRestrictions
        On Error GoTo 0
    Next
End Sub

Sub AplicarFundoGeral()
    Dim sh
    On Error Resume Next
    xl.ErrorCheckingOptions.BackgroundChecking = False
    On Error GoTo 0
    For Each sh In wb.Worksheets
        sh.Cells.EntireColumn.Hidden = False
        sh.Cells.EntireRow.Hidden = False
        sh.Cells.Interior.Pattern = xlNone
        sh.Cells.Borders.LineStyle = xlNone
        sh.Range("A1:AZ220").Interior.Color = RGB(226, 232, 240)
        sh.Cells.Font.Name = "Segoe UI"
        sh.Cells.Font.Size = 10
        sh.Cells.Font.Color = RGB(15, 23, 42)
        sh.Tab.Color = RGB(147, 197, 253)
    Next
End Sub

Sub GarantirEstruturaDespesas()
    Dim r
    Set ws = wb.Worksheets("Despesas")
    If LCase(Trim(CStr(ws.Cells(4, 19).Value))) <> "valor cota parte anual" Then
        ws.Columns(19).Insert
    End If
    ws.Cells(4, 18).Value = "Cota parte?"
    ws.Cells(4, 19).Value = "Valor cota parte anual"
    ws.Cells(4, 20).Value = "Valor cota parte mes"
    ws.Cells(4, 21).Value = "Comprovante"
    ws.Cells(4, 22).Value = "Observacoes"
    For r = 5 To 120
        ws.Cells(r, 16).Formula = "=IF(OR(A" & r & "<>"""",B" & r & "<>""""),(SUM(C" & r & ":N" & r & ")+O" & r & ")/12,"""")"
        ws.Cells(r, 17).Formula = "=IF(OR(A" & r & "<>"""",B" & r & "<>""""),SUM(C" & r & ":N" & r & ")+O" & r & ","""")"
        ws.Cells(r, 19).Formula = "=IF(OR(A" & r & "<>"""",B" & r & "<>""""),IF(R" & r & "=""Sim"",Q" & r & "/2,Q" & r & "),"""")"
        ws.Cells(r, 20).Formula = "=IF(OR(A" & r & "<>"""",B" & r & "<>""""),IF(R" & r & "=""Sim"",P" & r & "/2,P" & r & "),"""")"
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" Or Trim(CStr(ws.Cells(r, 2).Value)) <> "" Then
            If Trim(CStr(ws.Cells(r, 21).Value)) = "" Then ws.Cells(r, 21).Value = "Pendente"
        Else
            ws.Cells(r, 21).ClearContents
            ws.Cells(r, 22).ClearContents
        End If
    Next
    ws.Cells(122, 16).Formula = "=SUM(P5:P120)"
    ws.Cells(122, 17).Formula = "=SUM(Q5:Q120)"
    ws.Cells(122, 19).Formula = "=SUM(S5:S120)"
    ws.Cells(122, 20).Formula = "=SUM(T5:T120)"
End Sub

Sub FormatarDespesas()
    Dim r, fill, lastCol
    Set ws = wb.Worksheets("Despesas")
    lastCol = 23
    ws.Cells.EntireColumn.Hidden = False
    ws.Cells.EntireRow.Hidden = False
    ws.Range("A1:W221").Interior.Color = RGB(226, 232, 240)
    ws.Range("A1:W221").Font.Color = RGB(15, 23, 42)
    ws.Range("A1:W221").Borders.LineStyle = xlNone
    ws.Columns("X:XFD").Hidden = True
    ws.ScrollArea = ""
    Cabecalho ws.Range("A4:W4")
    For r = 5 To 120
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" Or Trim(CStr(ws.Cells(r, 2).Value)) <> "" Then
            fill = CorCategoria(ws.Cells(r, 1).Value)
        Else
            fill = RGB(226, 232, 240)
        End If
        ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = fill
        ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Font.Color = RGB(15, 23, 42)
        VincularStatusComprovante ws, r
        AplicarLink ws.Cells(r, 21), fill
        If Trim(CStr(ws.Cells(r, 22).Value)) <> "" Then AplicarLink ws.Cells(r, 22), fill
    Next
    AplicarStatusComprovante ws.Range("U5:U120")
    ws.Range("A4:W123").Borders.LineStyle = xlNone
    AplicarBordasDespesas ws, lastCol
    ws.Range("C5:Q123").NumberFormat = """R$"" #,##0.00"
    ws.Range("S5:T123").NumberFormat = """R$"" #,##0.00"
    ws.Range("R5:R120").HorizontalAlignment = xlCenter
    ws.Range("A122:W122").Interior.Color = RGB(148, 163, 184)
    ws.Range("A122:W122").Font.Color = RGB(15, 23, 42)
    ws.Range("A122:W122").Font.Bold = True
    ws.Columns("A:W").AutoFit
    ws.Columns("V").ColumnWidth = 58
    ws.Columns("W").ColumnWidth = 28
    ws.Range("V5:W123").WrapText = False
    ws.Range("V5:W123").ShrinkToFit = False
    ws.Range("V5:W123").HorizontalAlignment = xlLeft
    AtualizarVisibilidadeDespesas ws
    ws.Tab.Color = RGB(96, 165, 250)
End Sub


Sub AtualizarVisibilidadeDespesas(ws)
    Dim r, c, temCadastro, temValor, v
    On Error Resume Next
    For r = 5 To 120
        temCadastro = (Trim(CStr(ws.Cells(r, 1).Value)) <> "" Or Trim(CStr(ws.Cells(r, 2).Value)) <> "")
        temValor = False
        If temCadastro Then
            For c = 3 To 15
                If IsNumeric(ws.Cells(r, c).Value) Then
                    v = CDbl(ws.Cells(r, c).Value)
                    If Abs(v) > 0.000001 Then
                        temValor = True
                        Exit For
                    End If
                End If
            Next
            ws.Rows(r).Hidden = Not temValor
        Else
            ws.Rows(r).Hidden = True
        End If
    Next
    ws.Rows(4).Hidden = False
    ws.Rows(122).Hidden = False
    On Error GoTo 0
End Sub
Sub AplicarBordasDespesas(ws, lastCol)
    Dim rngDados, rngCab
    Set rngDados = ws.Range(ws.Cells(5, 1), ws.Cells(123, lastCol))
    rngDados.Borders.LineStyle = xlNone
    With rngDados.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(255, 255, 255)
    End With

    Set rngCab = ws.Range(ws.Cells(4, 1), ws.Cells(4, 23))
    With rngCab.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(255, 255, 255)
    End With
End Sub

Sub VincularStatusComprovante(ws, r)
    Dim status, obs
    status = Trim(CStr(ws.Cells(r, 21).Value))
    obs = Trim(CStr(ws.Cells(r, 22).Value))
    On Error Resume Next
    ws.Cells(r, 21).Hyperlinks.Delete
    On Error GoTo 0
    If status <> "" And obs <> "" Then
        ws.Hyperlinks.Add ws.Cells(r, 21), "", "'Controle'!A3", "", status
        ws.Hyperlinks.Add ws.Cells(r, 22), "", "'Controle'!A3", "", obs
    End If
End Sub

Sub CongelarReferenciasDespesas()
    wb.Worksheets("Despesas").Activate
    wb.Worksheets("Despesas").ScrollArea = ""
    wb.Worksheets("Despesas").Columns("V:W").Hidden = False
    wb.Worksheets("Despesas").Columns("X:XFD").Hidden = True
    wb.Worksheets("Despesas").Range("C5").Select
    With xl.ActiveWindow
        .FreezePanes = False
        .Split = False
        .SplitColumn = 2
        .SplitRow = 4
        .FreezePanes = True
        .ScrollColumn = 1
        .ScrollRow = 1
        .DisplayGridlines = False
        .DisplayHeadings = False
        .Zoom = 100
    End With
End Sub

Sub AplicarStatusComprovante(rng)
    Dim fc
    rng.FormatConditions.Delete
    rng.HorizontalAlignment = xlCenter
    rng.Font.Bold = True
    rng.Font.Underline = False

    Set fc = rng.FormatConditions.Add(1, 3, "=""Anexado""")
    fc.Interior.Color = RGB(22, 163, 74)
    fc.Font.Color = RGB(255, 255, 255)
    fc.Font.Bold = True

    AddStatusAusente rng, "Pendente"
    AddStatusAusente rng, "Ausente"
End Sub

Sub AddStatusAusente(rng, valor)
    Dim fc
    Set fc = rng.FormatConditions.Add(1, 3, "=""" & valor & """")
    fc.Interior.Color = RGB(203, 213, 225)
    fc.Font.Color = RGB(51, 65, 85)
    fc.Font.Bold = True
End Sub

Sub FormatarControle()
    Dim r, c, lastRow, lastCol, headerRow, catCol, linkCol, fill
    Set ws = wb.Worksheets("Controle")
    LimparShapes ws
    LimparTotalControleVisual ws
    headerRow = EncontrarLinhaCabecalho(ws, "categoria")
    If headerRow = 0 Then headerRow = 3
    lastRow = UltimaLinhaControleVisual(ws)
    If lastRow < headerRow Then lastRow = headerRow
    lastCol = UltimaColunaNaLinha(ws, headerRow)
    If lastCol < 9 Then lastCol = 9
    If lastCol > 27 Then lastCol = 27

    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow + 6, lastCol)).Interior.Color = RGB(226, 232, 240)
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow + 6, lastCol)).Font.Color = RGB(15, 23, 42)
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow + 6, lastCol)).Borders.LineStyle = xlNone

    Cabecalho ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
    ws.Columns(5).NumberFormat = """R$"" #,##0.00"
    Botao ws, "A1:C2", "Excluir selecionado", macroPrefix & "ExcluirRegistroControleSelecionado", RGB(254, 202, 202)
    Botao ws, "D1:E2", "Limpar filtro", macroPrefix & "LimparFiltroControle", RGB(219, 234, 254)

    catCol = EncontrarColunaCabecalho(ws, headerRow, "categoria")
    If catCol = 0 Then catCol = 2
    linkCol = EncontrarColunaCabecalho(ws, headerRow, "arquivo")
    If linkCol = 0 Then linkCol = EncontrarColunaCabecalho(ws, headerRow, "link")
    If linkCol = 0 Then linkCol = EncontrarColunaCabecalho(ws, headerRow, "comprovante")
    If linkCol = 0 Then linkCol = 7

    For r = headerRow + 1 To lastRow
        fill = CorCategoria(ws.Cells(r, catCol).Value)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = fill
        ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Font.Color = RGB(15, 23, 42)
        AplicarLink ws.Cells(r, linkCol), fill
    Next
    AtualizarTotalControleVisual ws, lastRow, lastCol
    ws.Range("J1:BA221").Interior.Color = RGB(226, 232, 240)
    ws.Range("J1:BA221").Borders.LineStyle = xlNone

    For c = 1 To lastCol
        ws.Columns(c).AutoFit
    Next
    ws.Tab.Color = RGB(125, 211, 252)
End Sub

Sub LimparTotalControleVisual(ws)
    Dim r
    For r = 4 To 2005
        If RemoveAcento(LCase(CStr(ws.Cells(r, 1).Value))) = "total filtrado" Then
            ws.Rows(r).Hidden = False
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).ClearContents
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Interior.Color = RGB(226, 232, 240)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Font.Bold = False
        End If
    Next
End Sub

Function UltimaLinhaControleVisual(ws)
    Dim r, lastA, lastHash, limite
    lastA = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    lastHash = ws.Cells(ws.Rows.Count, 9).End(-4162).Row
    limite = lastA
    If lastHash > limite Then limite = lastHash
    For r = limite To 4 Step -1
        If Trim(CStr(ws.Cells(r, 9).Value)) <> "" Then
            UltimaLinhaControleVisual = r
            Exit Function
        End If
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" And RemoveAcento(LCase(CStr(ws.Cells(r, 1).Value))) <> "total filtrado" Then
            UltimaLinhaControleVisual = r
            Exit Function
        End If
    Next
    UltimaLinhaControleVisual = 3
End Function

Sub AtualizarTotalControleVisual(ws, lastRow, lastCol)
    Dim linhaTotal
    If lastRow < 4 Then Exit Sub
    linhaTotal = lastRow + 1
    ws.Cells(linhaTotal, 1).Value = "Total filtrado"
    ws.Cells(linhaTotal, 4).Value = "Qtd."
    ws.Cells(linhaTotal, 5).Formula = "=SUBTOTAL(109,E4:E" & lastRow & ")"
    ws.Cells(linhaTotal, 6).Formula = "=SUBTOTAL(103,A4:A" & lastRow & ")"
    ws.Range(ws.Cells(linhaTotal, 1), ws.Cells(linhaTotal, lastCol)).Interior.Color = RGB(219, 234, 254)
    ws.Range(ws.Cells(linhaTotal, 1), ws.Cells(linhaTotal, lastCol)).Font.Color = RGB(15, 23, 42)
    ws.Range(ws.Cells(linhaTotal, 1), ws.Cells(linhaTotal, lastCol)).Font.Bold = True
    ws.Cells(linhaTotal, 5).NumberFormat = """R$"" #,##0.00"
End Sub

Sub FormatarFila()
    Dim r, lastRow, fill
    Set ws = wb.Worksheets("Fila")
    lastRow = UltimaLinha(ws, 1)
    If lastRow < 1 Then lastRow = 1
    Cabecalho ws.Range("A1:N1")
    If lastRow >= 2 Then
        For r = 2 To lastRow
            fill = CorCategoria(ws.Cells(r, 3).Value)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 14)).Interior.Color = fill
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 14)).Font.Color = RGB(15, 23, 42)
            AplicarLink ws.Cells(r, 11), fill
        Next
    End If
    ws.Range("A1:N" & lastRow).Borders.LineStyle = xlNone
    ws.Range("O1:AZ220").Interior.Color = RGB(226, 232, 240)
    ws.Range("O1:AZ220").Borders.LineStyle = xlNone
    ws.Columns("A:N").AutoFit
    ws.Tab.Color = RGB(45, 212, 191)
End Sub

Sub FormatarResumo()
    Dim r, fill
    Set ws = wb.Worksheets("Resumo e Rateio")
    ws.Cells.UnMerge
    ws.Cells.EntireColumn.Hidden = False
    ws.Cells.EntireRow.Hidden = False
    ws.Range("A1:H18").Interior.Color = RGB(226, 232, 240)
    ws.Range("A1:H18").Font.Color = RGB(15, 23, 42)
    ws.Range("A1:H18").Borders.LineStyle = xlNone
    ws.Range("A1:H18").VerticalAlignment = xlCenter
    ws.Range("I1:AZ220").Interior.Color = RGB(226, 232, 240)
    ws.Range("I1:AZ220").Borders.LineStyle = xlNone
    ws.Range("D1:D18").ClearContents
    ws.Range("D1:D18").Interior.Color = RGB(226, 232, 240)

    ws.Range("A1:G1").Merge
    ws.Range("A1").Value = "Resumo e Rateio"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(30, 58, 95)

    ws.Cells(3, 1).Value = "Categoria"
    ws.Cells(3, 2).Value = "Cota parte mes (R$)"
    ws.Cells(3, 3).Value = "Cota parte anual (R$)"
    Cabecalho ws.Range("A3:C3")
    For r = 4 To 13
        fill = CorCategoria(ws.Cells(r, 1).Value)
        ws.Range("A" & r & ":C" & r).Interior.Color = fill
        ws.Range("A" & r & ":C" & r).Font.Color = RGB(15, 23, 42)
        ws.Cells(r, 2).Formula = "=SUMIFS(Despesas!$T:$T,Despesas!$A:$A,A" & r & ")"
        ws.Cells(r, 3).Formula = "=SUMIFS(Despesas!$S:$S,Despesas!$A:$A,A" & r & ")"
    Next
    ws.Range("A14:C14").Interior.Color = RGB(219, 234, 254)
    ws.Range("A14:C14").Font.Color = RGB(15, 23, 42)
    ws.Range("A14:C14").Font.Bold = True
    ws.Cells(14, 1).Value = "TOTAL GERAL"
    ws.Cells(14, 2).Formula = "=SUM(B4:B13)"
    ws.Cells(14, 3).Formula = "=SUM(C4:C13)"
    ws.Range("B4:C14").NumberFormat = """R$"" #,##0.00"

    ws.Range("E3:G3").Merge
    ws.Range("E3").Value = "CAPACIDADE CONTRIBUTIVA E RATEIO"
    Cabecalho ws.Range("E3:G3")
    Cabecalho ws.Range("E4:G4")
    ws.Range("E4").Value = "Indicador"
    ws.Range("F4").Value = "Genitor 1"
    ws.Range("G4").Value = "Genitor 2"
    ws.Range("E5").Value = "Renda liquida mensal"
    ws.Range("E6").Value = "Outras obrigacoes essenciais"
    ws.Range("E7").Value = "Renda disponivel"
    ws.Range("E8").Value = "Percentual da renda disponivel"
    ws.Range("E9").Value = "Cota proporcional das despesas"
    ws.Range("F5").Formula = "=Config!B3"
    ws.Range("G5").Formula = "=Config!C3"
    ws.Range("F6").Formula = "=Config!B4"
    ws.Range("G6").Formula = "=Config!C4"
    ws.Range("F7").Formula = "=F5-F6"
    ws.Range("G7").Formula = "=G5-G6"
    ws.Range("F8").Formula = "=IF(F7+G7=0,0,F7/(F7+G7))"
    ws.Range("G8").Formula = "=IF(F7+G7=0,0,G7/(F7+G7))"
    ws.Range("F9").Formula = "=$B$14*F8"
    ws.Range("G9").Formula = "=$B$14*G8"
    ws.Range("E5:G9").Interior.Color = RGB(245, 243, 255)
    ws.Range("E5:G9").Font.Color = RGB(15, 23, 42)
    ws.Range("F5:G7").NumberFormat = """R$"" #,##0.00"
    ws.Range("F8:G8").NumberFormatLocal = "0,00%"
    ws.Range("F9:G9").NumberFormat = """R$"" #,##0.00"

    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B:C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 34
    ws.Columns("F:G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 2.5
    ws.Columns("I:XFD").Hidden = True
    ws.Rows("19:1048576").Hidden = True
    ws.ScrollArea = "A1:H18"
    ws.Tab.Color = RGB(167, 139, 250)
End Sub

Sub FormatarConfig()
    Dim renda1, renda2, obrig1, obrig2, moradores, caminho
    Set ws = wb.Worksheets("Config")
    DesprotegerAba ws

    renda1 = ws.Range("B3").Value
    renda2 = ws.Range("C3").Value
    obrig1 = ws.Range("B4").Value
    obrig2 = ws.Range("C4").Value
    moradores = ws.Range("B6").Value
    caminho = ws.Range("B7").Value

    ws.Cells.UnMerge
    ws.Cells.EntireColumn.Hidden = False
    ws.Cells.EntireRow.Hidden = False
    ws.Cells.Clear
    ws.Range("A1:F12").Interior.Color = RGB(219, 228, 238)
    ws.Range("A1:F12").Font.Name = "Segoe UI"
    ws.Range("A1:F12").Font.Size = 10
    ws.Range("A1:F12").Font.Color = RGB(15, 23, 42)
    ws.Range("A1:F12").Borders.LineStyle = xlNone

    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B:C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F").ColumnWidth = 22
    ws.Rows("1:12").RowHeight = 22
    ws.Rows("5:5").RowHeight = 8
    ws.Rows("8:8").RowHeight = 8

    ws.Range("A1:F1").Merge
    ws.Range("A1").Value = "Configura��es"
    ws.Range("A1").Font.Size = 17
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(30, 58, 95)

    ws.Range("A2:C2").Interior.Color = RGB(30, 58, 95)
    ws.Range("A2:C2").Font.Color = RGB(255, 255, 255)
    ws.Range("A2:C2").Font.Bold = True
    ws.Range("A2").Value = "Par�metro"
    ws.Range("B2").Value = "Genitor 1"
    ws.Range("C2").Value = "Genitor 2"
    ws.Range("A3").Value = "Renda l�quida mensal (R$)"
    ws.Range("A4").Value = "Outras obriga��es essenciais (R$)"
    ws.Range("B3").Value = renda1
    ws.Range("C3").Value = renda2
    ws.Range("B4").Value = obrig1
    ws.Range("C4").Value = obrig2
    ws.Range("A3:C4").Interior.Color = RGB(241, 245, 249)
    ws.Range("B3:C4").NumberFormat = """R$"" #,##0.00"

    ws.Range("A6:C6").Interior.Color = RGB(30, 58, 95)
    ws.Range("A6:C6").Font.Color = RGB(255, 255, 255)
    ws.Range("A6:C6").Font.Bold = True
    ws.Range("A6").Value = "Par�metro"
    ws.Range("B6").Value = moradores
    ws.Range("C6").Value = ""
    ws.Range("A7").Value = "Caminho da pasta de comprovantes"
    ws.Range("B7:C7").Merge
    ws.Range("B7").Value = caminho
    ws.Range("A7:C7").Interior.Color = RGB(241, 245, 249)
    ws.Range("A6").Value = "N� de moradores (rateio moradia)"
    ws.Range("B6").NumberFormat = "0"

    ws.Range("E2:F2").Merge
    ws.Range("E2").Value = "Status"
    ws.Range("E2:F2").Interior.Color = RGB(30, 58, 95)
    ws.Range("E2:F2").Font.Color = RGB(255, 255, 255)
    ws.Range("E2:F2").Font.Bold = True
    ws.Range("E3").Value = "Planilha"
    ws.Range("F3").Value = "Destravada"
    ws.Range("E4").Value = "Base"
    ws.Range("F4").Value = wb.Path
    ws.Range("E3:F4").Interior.Color = RGB(241, 245, 249)

    ws.Columns("G:XFD").Hidden = True
    ws.Rows("13:1048576").Hidden = True
    ws.ScrollArea = "A1:F12"
    ws.Tab.Color = RGB(96, 165, 250)
End Sub
Sub RedesenharImport()
    Set ws = wb.Worksheets("Import")
    LimparShapes ws
    ws.Cells.Clear
    ws.Cells.EntireColumn.Hidden = False
    ws.Cells.EntireRow.Hidden = False

    ws.Range("A1:K19").Interior.Color = RGB(214, 224, 234)
    ws.Range("A1:K19").Font.Color = RGB(15, 23, 42)
    ws.Range("A1:K19").Borders.LineStyle = xlNone
    ws.Range("A1:K19").HorizontalAlignment = xlLeft
    ws.Range("A1:K19").VerticalAlignment = xlCenter

    ws.Columns("A").ColumnWidth = 2.5
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 3
    ws.Columns("H").ColumnWidth = 18
    ws.Columns("I:K").ColumnWidth = 2.5
    ws.Rows("1:18").RowHeight = 22
    ws.Rows("3:3").RowHeight = 8
    ws.Rows("6:6").RowHeight = 8
    ws.Rows("10:10").RowHeight = 8
    ws.Rows("14:14").RowHeight = 8

    ws.Range("B2:F2").Merge
    ws.Range("B2").Value = "Comprovantes"
    ws.Range("B2").Font.Name = "Segoe UI"
    ws.Range("B2").Font.Size = 17
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(15, 23, 42)

    ws.Range("B4:H4").Interior.Color = RGB(198, 211, 224)
    ws.Range("B4:H4").Borders.LineStyle = xlNone
    SecaoImport ws, "B4", "Preparar"
    BotaoImport ws, "B5:C5", "Abrir ferramenta", macroPrefix & "AbrirFerramentaExtracao", RGB(51, 65, 85), RGB(71, 85, 105), RGB(248, 250, 252)
    BotaoImport ws, "E5:F5", "Extrair", macroPrefix & "ExtrairDados", RGB(37, 99, 235), RGB(29, 78, 216), RGB(255, 255, 255)
    BotaoImport ws, "H5:H5", "Fila JSON", macroPrefix & "CarregarFilaJson", RGB(15, 118, 110), RGB(13, 148, 136), RGB(255, 255, 255)

    ws.Range("B8:H8").Interior.Color = RGB(198, 211, 224)
    ws.Range("B8:H8").Borders.LineStyle = xlNone
    SecaoImport ws, "B8", "Importar"
    BotaoImport ws, "B9:C9", "Pendentes", macroPrefix & "ImportarTodosPendentes", RGB(22, 101, 52), RGB(21, 128, 61), RGB(255, 255, 255)
    BotaoImport ws, "E9:F9", "Linha selecionada", macroPrefix & "ImportarSelecionadoDaFila", RGB(6, 95, 70), RGB(5, 150, 105), RGB(255, 255, 255)

    ws.Range("B12:H12").Interior.Color = RGB(198, 211, 224)
    ws.Range("B12:H12").Borders.LineStyle = xlNone
    SecaoImport ws, "B12", "Manutencao"
    BotaoImport ws, "B13:C13", "Cadastros", macroPrefix & "AbrirCadastros", RGB(120, 53, 15), RGB(146, 64, 14), RGB(255, 255, 255)
    BotaoImport ws, "E13:F13", "Abrir arquivo", macroPrefix & "AbrirComprovante", RGB(124, 45, 18), RGB(154, 52, 18), RGB(255, 255, 255)
    BotaoImport ws, "H13:H13", "Limpar", macroPrefix & "LimparDados", RGB(127, 29, 29), RGB(153, 27, 27), RGB(255, 255, 255)

    ws.Range("B16:H16").Interior.Color = RGB(198, 211, 224)
    ws.Range("B16:H16").Borders.LineStyle = xlNone
    SecaoImport ws, "B16", "Edicao"
    BotaoImport ws, "B17:C17", "Destravar", macroPrefix & "DesbloquearEdicao", RGB(71, 85, 105), RGB(100, 116, 139), RGB(248, 250, 252)
    BotaoImport ws, "E17:F17", "Travar", macroPrefix & "BloquearEdicao", RGB(71, 85, 105), RGB(100, 116, 139), RGB(248, 250, 252)

    ws.Columns("L:XFD").Hidden = True
    ws.Rows("20:1048576").Hidden = True
    ws.ScrollArea = "A1:K19"
    ws.Tab.Color = RGB(71, 85, 105)
End Sub

Sub OcultarAbaEditar()
    On Error Resume Next
    wb.Worksheets("Editar").Visible = False
    On Error GoTo 0
End Sub

Sub FormatarCadastros()
    Dim ws, wsD, r, out, i
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("Cadastros")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Cadastros"
    End If
    Set wsD = wb.Worksheets("Despesas")
    DesprotegerAba ws
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next
    ws.Cells.EntireColumn.Hidden = False
    ws.Cells.EntireRow.Hidden = False
    ws.Range("A1:H70").Interior.Color = RGB(226, 232, 240)
    ws.Range("A1:H70").Borders.LineStyle = xlNone
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Cells.Font.Color = RGB(15, 23, 42)
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 42
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F:H").ColumnWidth = 16

    ws.Range("B2:G2").Merge
    ws.Range("B2").Value = "Cadastro de categorias e despesas"
    ws.Range("B2").Font.Size = 18
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(30, 58, 95)

    ws.Range("B4").Value = "Categoria"
    ws.Range("B6").Value = "Despesa"
    ws.Range("B8").Value = "Cota parte?"
    ws.Range("D8").Value = "Periodicidade"
    ws.Range("F8").Value = "Ativo"
    ws.Range("B4:F8").Font.Bold = True
    Campo ws.Range("C4:E4")
    Campo ws.Range("C6:G6")
    Campo ws.Range("C8")
    Campo ws.Range("E8")
    Campo ws.Range("G8")
    If Trim(CStr(ws.Range("C8").Value)) = "" Then ws.Range("C8").Value = "Nao"
    If Trim(CStr(ws.Range("E8").Value)) = "" Then ws.Range("E8").Value = "Mensal"
    If Trim(CStr(ws.Range("G8").Value)) = "" Then ws.Range("G8").Value = "Sim"

    Botao ws, "B10:C11", "Salvar cadastro", macroPrefix & "SalvarCadastroDespesa", RGB(187, 247, 208)
    Botao ws, "D10:E11", "Carregar selecao", macroPrefix & "CarregarCadastroSelecionado", RGB(191, 219, 254)
    Botao ws, "F10:G11", "Sincronizar", macroPrefix & "SincronizarCadastrosDespesas", RGB(253, 230, 138)

    ws.Cells(12, 1).Value = "Categoria"
    ws.Cells(12, 2).Value = "Despesa"
    ws.Cells(12, 3).Value = "Cota parte?"
    ws.Cells(12, 4).Value = "Periodicidade"
    ws.Cells(12, 5).Value = "Ativo"
    Cabecalho ws.Range("A12:E12")
    If Trim(CStr(ws.Cells(13, 1).Value)) = "" Then
        out = 13
        For r = 5 To 120
            If Trim(CStr(wsD.Cells(r, 1).Value)) <> "" And Trim(CStr(wsD.Cells(r, 2).Value)) <> "" Then
                ws.Cells(out, 1).Value = wsD.Cells(r, 1).Value
                ws.Cells(out, 2).Value = wsD.Cells(r, 2).Value
                ws.Cells(out, 3).Value = wsD.Cells(r, 18).Value
                ws.Cells(out, 4).Value = "Mensal"
                ws.Cells(out, 5).Value = "Sim"
                out = out + 1
            End If
        Next
    End If
    ws.Range("A13:E70").Borders.LineStyle = xlNone
    ws.Range("A13:E70").Interior.Color = RGB(241, 245, 249)
    ws.Range("A13:E70").Font.Color = RGB(15, 23, 42)
    ws.Columns("I:XFD").Hidden = True
    ws.Rows("71:1048576").Hidden = True
    ws.ScrollArea = "A1:H70"
    ws.Tab.Color = RGB(34, 197, 94)
End Sub

Sub RedesenharEditar()
    Set ws = wb.Worksheets("Editar")
    LimparShapes ws
    ws.Cells.Clear
    ws.Cells.EntireColumn.Hidden = False
    ws.Cells.EntireRow.Hidden = False
    ws.Range("A1:H24").Interior.Color = RGB(226, 232, 240)
    ws.Range("A1:H24").Borders.LineStyle = xlNone
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D:G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 3
    ws.Rows("1:24").RowHeight = 21
    ws.Range("B2:G2").Merge
    ws.Range("B2").Value = "Editar registro"
    ws.Range("B2").Font.Size = 18
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(30, 58, 95)
    ws.Range("B4").Value = "Origem"
    ws.Range("D4").Value = "Linha"
    ws.Range("B5").Value = "Fila"
    ws.Range("D5").Value = 2
    Campo ws.Range("B5")
    Campo ws.Range("D5:G5")
    Label ws, 8, "Categoria"
    Label ws, 9, "Despesa / Recebedor"
    Label ws, 10, "Data"
    Label ws, 11, "Competencia"
    Label ws, 12, "Valor"
    Label ws, 13, "Pagador"
    Label ws, 14, "Recebedor"
    Label ws, 15, "Tipo / Observacao"
    Label ws, 16, "Periodicidade"
    Label ws, 17, "Arquivo"
    Label ws, 18, "Cota parte?"
    Campo ws.Range("D8:G18")
    ws.Range("D10").NumberFormat = "dd/mm/yyyy"
    ws.Range("D11").NumberFormat = "yyyy-mm"
    ws.Range("D12").NumberFormat = """R$"" #,##0.00"
    Botao ws, "B21:C22", "Carregar", macroPrefix & "CarregarRegistroEdicao", RGB(191, 219, 254)
    Botao ws, "D21:E22", "Salvar", macroPrefix & "SalvarFormularioEdicao", RGB(187, 247, 208)
    ws.ScrollArea = "A1:H24"
    ws.Tab.Color = RGB(147, 197, 253)
End Sub

Sub Cabecalho(rng)
    rng.Interior.Color = RGB(30, 58, 95)
    rng.Font.Color = RGB(255, 255, 255)
    rng.Font.Bold = True
    rng.HorizontalAlignment = xlCenter
End Sub

Sub LinhaTitulo(ws, addr, txt)
    With ws.Range(addr)
        .Merge
        .Value = txt
        .Font.Bold = True
        .Font.Color = RGB(30, 58, 95)
        .Interior.Color = RGB(219, 234, 254)
    End With
End Sub

Sub SecaoImport(ws, addr, txt)
    With ws.Range(addr)
        .Value = UCase(txt)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = RGB(51, 65, 85)
        .Interior.Color = RGB(198, 211, 224)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlNone
    End With
End Sub

Sub Campo(rng)
    rng.Interior.Color = RGB(248, 250, 252)
    rng.Font.Color = RGB(30, 41, 59)
    rng.Borders.LineStyle = xlNone
End Sub

Sub AplicarLink(cel, fill)
    cel.Interior.Color = fill
    cel.Font.Color = RGB(37, 99, 235)
    cel.Font.Bold = True
    cel.Font.Underline = False
End Sub

Sub AplicarNeutro(cel)
    cel.Interior.Color = RGB(245, 243, 255)
    cel.Font.Color = RGB(15, 23, 42)
End Sub

Sub Label(ws, rowNum, txt)
    ws.Cells(rowNum, 2).Value = txt
    ws.Cells(rowNum, 2).Font.Bold = True
    ws.Cells(rowNum, 2).Font.Color = RGB(30, 58, 95)
End Sub

Sub Botao(ws, addr, caption, macroName, fill)
    Dim rng, sh
    Set rng = ws.Range(addr)
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left + 2, rng.Top + 2, rng.Width - 4, rng.Height - 4)
    sh.Fill.ForeColor.RGB = fill
    sh.Line.ForeColor.RGB = RGB(147, 197, 253)
    sh.TextFrame.Characters.Text = caption
    sh.TextFrame.HorizontalAlignment = xlCenter
    sh.TextFrame.VerticalAlignment = xlCenter
    sh.TextFrame.Characters.Font.Name = "Segoe UI"
    sh.TextFrame.Characters.Font.Size = 9
    sh.TextFrame.Characters.Font.Bold = True
    sh.TextFrame.Characters.Font.Color = RGB(30, 58, 95)
    sh.OnAction = macroName
    sh.Placement = 3
End Sub

Sub BotaoImport(ws, addr, caption, macroName, fill, borderColor, fontColor)
    Dim rng, sh
    Set rng = ws.Range(addr)
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left + 1, rng.Top + 1, rng.Width - 2, rng.Height - 2)
    sh.Fill.ForeColor.RGB = fill
    sh.Line.ForeColor.RGB = borderColor
    sh.Line.Weight = 1
    sh.TextFrame.Characters.Text = caption
    sh.TextFrame.HorizontalAlignment = xlCenter
    sh.TextFrame.VerticalAlignment = xlCenter
    sh.TextFrame.Characters.Font.Name = "Segoe UI"
    sh.TextFrame.Characters.Font.Size = 8
    sh.TextFrame.Characters.Font.Bold = True
    sh.TextFrame.Characters.Font.Color = fontColor
    On Error Resume Next
    sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fontColor
    sh.TextFrame2.VerticalAnchor = 3
    On Error GoTo 0
    sh.OnAction = macroName
    sh.Placement = 3
End Sub

Sub LimparShapes(ws)
    Dim i
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next
End Sub

Function UltimaLinha(ws, col)
    UltimaLinha = ws.Cells(ws.Rows.Count, col).End(-4162).Row
End Function

Function UltimaLinhaDados(ws)
    Dim c, v, maior
    maior = 1
    For c = 1 To 26
        v = ws.Cells(ws.Rows.Count, c).End(-4162).Row
        If v > maior Then maior = v
    Next
    UltimaLinhaDados = maior
End Function

Function UltimaColunaNaLinha(ws, rowNum)
    UltimaColunaNaLinha = ws.Cells(rowNum, ws.Columns.Count).End(-4159).Column
End Function

Function EncontrarLinhaCabecalho(ws, termo)
    Dim r, c, txt
    termo = RemoveAcento(LCase(CStr(termo)))
    For r = 1 To 20
        For c = 1 To 26
            txt = RemoveAcento(LCase(CStr(ws.Cells(r, c).Value)))
            If InStr(txt, termo) > 0 Then
                EncontrarLinhaCabecalho = r
                Exit Function
            End If
        Next
    Next
    EncontrarLinhaCabecalho = 0
End Function

Function EncontrarColunaCabecalho(ws, rowNum, termo)
    Dim c, txt
    termo = RemoveAcento(LCase(CStr(termo)))
    For c = 1 To 26
        txt = RemoveAcento(LCase(CStr(ws.Cells(rowNum, c).Value)))
        If InStr(txt, termo) > 0 Then
            EncontrarColunaCabecalho = c
            Exit Function
        End If
    Next
    EncontrarColunaCabecalho = 0
End Function

Function CorCategoria(cat)
    Dim t
    t = RemoveAcento(LCase(CStr(cat)))
    If InStr(t, "moradia") > 0 Then
        CorCategoria = RGB(191, 219, 254)
    ElseIf InStr(t, "alimenta") > 0 Then
        CorCategoria = RGB(187, 247, 208)
    ElseIf InStr(t, "educa") > 0 Then
        CorCategoria = RGB(253, 230, 138)
    ElseIf InStr(t, "saude") > 0 Then
        CorCategoria = RGB(251, 207, 232)
    ElseIf InStr(t, "transporte") > 0 Then
        CorCategoria = RGB(199, 210, 254)
    ElseIf InStr(t, "vestuario") > 0 Then
        CorCategoria = RGB(221, 214, 254)
    ElseIf InStr(t, "comunica") > 0 Then
        CorCategoria = RGB(153, 246, 228)
    ElseIf InStr(t, "lazer") > 0 Then
        CorCategoria = RGB(254, 215, 170)
    ElseIf InStr(t, "esporte") > 0 Then
        CorCategoria = RGB(186, 230, 253)
    Else
        CorCategoria = RGB(203, 213, 225)
    End If
End Function

Function RemoveAcento(s)
    s = Replace(s, ChrW(225), "a")
    s = Replace(s, ChrW(224), "a")
    s = Replace(s, ChrW(226), "a")
    s = Replace(s, ChrW(227), "a")
    s = Replace(s, ChrW(233), "e")
    s = Replace(s, ChrW(234), "e")
    s = Replace(s, ChrW(237), "i")
    s = Replace(s, ChrW(243), "o")
    s = Replace(s, ChrW(244), "o")
    s = Replace(s, ChrW(245), "o")
    s = Replace(s, ChrW(250), "u")
    s = Replace(s, ChrW(231), "c")
    RemoveAcento = s
End Function

Sub NormalizarJanela(nomeAba)
    wb.Worksheets(nomeAba).Activate
    wb.Worksheets(nomeAba).Range("A1").Select
    With xl.ActiveWindow
        .FreezePanes = False
        .Split = False
        .SplitColumn = 0
        .SplitRow = 0
        .ScrollColumn = 1
        .ScrollRow = 1
        .DisplayGridlines = False
        .DisplayHeadings = False
        .Zoom = 100
    End With
End Sub
