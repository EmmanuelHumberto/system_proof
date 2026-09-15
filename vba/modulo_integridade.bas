Attribute VB_Name = "ModuloIntegridade"
Option Explicit

Public Sub RegistrarAlias(categoria As String, despesa As String, id As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Aliases")
    For r = 2 To UltimaLinhaPreenchida(ws, 1)
        If LCase$(CStr(ws.Cells(r, 1).Value)) = LCase$(categoria) And LCase$(CStr(ws.Cells(r, 2).Value)) = LCase$(despesa) Then
            If CStr(ws.Cells(r, 3).Value) <> id Then Err.Raise vbObjectError + 220, , "Nome anterior pertence a outro cadastro."
            Exit Sub
        End If
    Next r
    r = Application.Max(2, UltimaLinhaPreenchida(ws, 1) + 1)
    ws.Cells(r, 1).Value = categoria: ws.Cells(r, 2).Value = despesa: ws.Cells(r, 3).Value = id
End Sub

Public Function ResolverAlias(categoria As String, despesa As String) As Long
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Aliases")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For r = 2 To UltimaLinhaPreenchida(ws, 1)
        If LCase$(CStr(ws.Cells(r, 1).Value)) = LCase$(categoria) And LCase$(CStr(ws.Cells(r, 2).Value)) = LCase$(despesa) Then
            ResolverAlias = LinhaPorId(ThisWorkbook.Worksheets("Despesas"), 24, ws.Cells(r, 3).Value, 5)
            Exit Function
        End If
    Next r
End Function

Public Function LinhaPorId(ws As Worksheet, coluna As Long, id As Variant, inicio As Long) As Long
    Dim r As Long
    If Trim$(CStr(id)) = "" Then Exit Function
    For r = inicio To UltimaLinhaPreenchida(ws, coluna)
        If CStr(ws.Cells(r, coluna).Value) = CStr(id) Then
            If LinhaPorId <> 0 Then Err.Raise vbObjectError + 201, , "Identificador repetido. Revise Conciliacao."
            LinhaPorId = r
        End If
    Next r
End Function

Public Function NovoControle() As Long
    Dim ws As Worksheet, r As Long, maior As Long
    Set ws = ThisWorkbook.Worksheets("Controle")
    maior = CLng(Val(CStr(ThisWorkbook.Worksheets("Config").Range("F10").Value)))
    For r = 4 To UltimaLinhaPreenchida(ws, 1)
        If IsNumeric(ws.Cells(r, 1).Value) Then maior = Application.Max(maior, Val(CStr(ws.Cells(r, 1).Value)))
    Next r
    NovoControle = maior + 1
    ThisWorkbook.Worksheets("Config").Range("F10").Value = NovoControle
End Function

Public Function NovaDespesa() As String
    SincronizarSequencias
    With ThisWorkbook.Worksheets("Config").Range("F11")
        .Value = CLng(Val(CStr(.Value))) + 1
        NovaDespesa = "D" & Format$(.Value, "000000")
    End With
End Function

Public Function Competencia(valor As Variant) As String
    Dim s As String, dt As Date, partes As Variant, y As Long, m As Long, dia As Long
    If VarType(valor) = vbDate Then
        Competencia = Format$(CDate(valor), "yyyy-mm")
        Exit Function
    End If
    s = Trim$(CStr(valor))
    If s Like "####-##" Then
        y = CLng(Left$(s, 4)): m = CLng(Right$(s, 2))
    ElseIf s Like "##/##/####" Then
        partes = Split(s, "/")
        dia = CLng(partes(0)): m = CLng(partes(1)): y = CLng(partes(2))
        dt = DateSerial(y, m, dia)
        If Day(dt) <> dia Or Month(dt) <> m Then Err.Raise vbObjectError + 202, , "Data invalida."
    Else
        Err.Raise vbObjectError + 203, , "Competencia invalida. Use AAAA-MM."
    End If
    If y < 1900 Or y > 9999 Or m < 1 Or m > 12 Then Err.Raise vbObjectError + 204, , "Competencia invalida."
    Competencia = Format$(DateSerial(y, m, 1), "yyyy-mm")
End Function

Public Function InicioPeriodo() As Date
    Dim ws As Worksheet, s As String, inicio As Date, c As Long
    Set ws = ThisWorkbook.Worksheets("Despesas")
    s = Competencia(ws.Cells(4, 3).Value)
    inicio = DateSerial(CLng(Left$(s, 4)), CLng(Right$(s, 2)), 1)
    For c = 3 To 14
        If Competencia(ws.Cells(4, c).Value) <> Format$(DateAdd("m", c - 3, inicio), "yyyy-mm") Then
            Err.Raise vbObjectError + 233, , "Cabecalhos devem conter 12 meses consecutivos. Nenhum valor alterado."
        End If
    Next c
    InicioPeriodo = inicio
End Function

Public Function ColunaPeriodo(mes As Variant, periodicidade As String, Optional validarAno As Boolean = True) As Long
    Dim s As String, inicio As Date, dataMes As Date, deslocamento As Long
    s = Competencia(mes)
    inicio = InicioPeriodo()
    dataMes = DateSerial(CLng(Left$(s, 4)), CLng(Right$(s, 2)), 1)
    deslocamento = DateDiff("m", inicio, dataMes)
    ' Keep the legacy argument, but never bypass the period when editing or deleting.
    If deslocamento < 0 Or deslocamento > 11 Then
        Err.Raise vbObjectError + 206, , "Competencia " & s & " fora do periodo " & Format$(inicio, "mm/yyyy") & " a " & Format$(DateAdd("m", 11, inicio), "mm/yyyy") & ". Nenhum valor alterado."
    End If
    If LCase$(periodicidade) = "anual" Or LCase$(periodicidade) = "eventual" Then
        ColunaPeriodo = 15
    ElseIf LCase$(periodicidade) = "mensal" Then
        ColunaPeriodo = deslocamento + 3
    Else
        Err.Raise vbObjectError + 205, , "Periodicidade deve ser Mensal, Anual ou Eventual."
    End If
End Function

Public Sub ValidarMovimento(r As Long, ByRef d As Long, ByRef col As Long)
    Dim wc As Worksheet, wd As Worksheet, mesma As Long
    Set wc = ThisWorkbook.Worksheets("Controle"): Set wd = ThisWorkbook.Worksheets("Despesas")
    If r < 4 Then Err.Raise vbObjectError + 207, , "Selecione um registro."
    mesma = LinhaPorId(wc, 1, wc.Cells(r, 1).Value, 4)
    If mesma = 0 Or CStr(wc.Cells(r, 12).Value) <> "OK" Then Err.Raise vbObjectError + 208, , "Registro pendente de conciliacao. Nenhum valor alterado."
    d = LinhaPorId(wd, 24, wc.Cells(r, 10).Value, 5)
    If d = 0 Then Err.Raise vbObjectError + 209, , "Vinculo de despesa ausente."
    col = ColunaPeriodo(wc.Cells(r, 4).Value, CStr(wc.Cells(r, 11).Value), False)
    If Not IsNumeric(wd.Cells(d, col).Value) Then Err.Raise vbObjectError + 210, , "Saldo da despesa ausente."
    If CDbl(wd.Cells(d, col).Value) + 0.001 < CDbl(wc.Cells(r, 5).Value) Then Err.Raise vbObjectError + 211, , "Saldo insuficiente. Confira Conciliacao."
End Sub

Public Sub AtualizarReferencias()
    Dim wc As Worksheet, wd As Worksheet, r As Long, d As Long, id As String, texto As String
    Dim referencias As Object, pendentes As Object
    Set referencias = CreateObject("Scripting.Dictionary")
    Set pendentes = CreateObject("Scripting.Dictionary")
    Set wc = ThisWorkbook.Worksheets("Controle"): Set wd = ThisWorkbook.Worksheets("Despesas")
    For r = 4 To UltimaLinhaPreenchida(wc, 1)
        id = CStr(wc.Cells(r, 10).Value)
        If id <> "" And IsNumeric(wc.Cells(r, 1).Value) Then
            If Not referencias.Exists(id) Then referencias.Add id, ""
            If referencias(id) <> "" Then referencias(id) = referencias(id) & ", "
            referencias(id) = referencias(id) & CStr(wc.Cells(r, 1).Value)
            If CStr(wc.Cells(r, 12).Value) <> "OK" Then pendentes(id) = True
        End If
    Next r
    For d = 5 To wd.Cells(wd.Rows.Count, 2).End(xlUp).Row
        id = CStr(wd.Cells(d, 24).Value)
        If id <> "" Then
            wd.Cells(d, 22).Hyperlinks.Delete
            wd.Cells(d, 21).Hyperlinks.Delete
            If referencias.Exists(id) Then
                texto = "Controle " & referencias(id)
                wd.Hyperlinks.Add Anchor:=wd.Cells(d, 22), Address:="", SubAddress:="'Controle'!A3", TextToDisplay:=texto
                If pendentes.Exists(id) Then wd.Cells(d, 21).Value = "Conferir" Else wd.Cells(d, 21).Value = "Anexado"
            Else
                wd.Cells(d, 22).ClearContents
                wd.Cells(d, 21).Value = "Conferir"
            End If
        End If
    Next d
End Sub

Public Sub CarregarControle(r As Long)
    Dim wc As Worksheet, wd As Worksheet, we As Worksheet, d As Long
    Set wc = ThisWorkbook.Worksheets("Controle"): Set wd = ThisWorkbook.Worksheets("Despesas")
    Set we = ThisWorkbook.Worksheets("Editar")
    d = LinhaPorId(wd, 24, wc.Cells(r, 10).Value, 5)
    If d = 0 Then Err.Raise vbObjectError + 212, , "Registro sem vinculo confirmado. Confira Conciliacao."
    we.Range("B7").Value = wc.Cells(r, 1).Value
    we.Range("D8").Value = wd.Cells(d, 1).Value
    we.Range("D9").Value = wd.Cells(d, 2).Value
    we.Range("D11").Value = wc.Cells(r, 4).Value
    we.Range("D12").Value = wc.Cells(r, 5).Value
    we.Range("D14").Value = wc.Cells(r, 3).Value
    we.Range("D15").Value = wc.Cells(r, 8).Value
    we.Range("D16").Value = wc.Cells(r, 11).Value
    we.Range("D17").Value = wc.Cells(r, 7).Value
    we.Range("D18").Value = wd.Cells(d, 18).Value
End Sub

Public Sub SalvarControle()
    Dim wc As Worksheet, wd As Worksheet, we As Worksheet, r As Long, d As Long, c As Long, novoC As Long
    Dim valor As Double, antigo As Double, oldControl As Variant, saldo As Variant, saldoNovo As Variant, gravando As Boolean
    Dim mes As String, per As String, msg As String
    On Error GoTo Falha
    Set wc = ThisWorkbook.Worksheets("Controle"): Set wd = ThisWorkbook.Worksheets("Despesas")
    Set we = ThisWorkbook.Worksheets("Editar")
    r = LinhaPorId(wc, 1, we.Range("B7").Value, 4)
    ValidarMovimento r, d, c
    If CStr(we.Range("D8").Value) <> CStr(wd.Cells(d, 1).Value) Or CStr(we.Range("D9").Value) <> CStr(wd.Cells(d, 2).Value) Then Err.Raise vbObjectError + 213, , "Para reclassificar um comprovante, confira o vinculo em Conciliacao."
    If CStr(we.Range("D17").Value) <> CStr(wc.Cells(r, 7).Value) Then Err.Raise vbObjectError + 214, , "Nao altere o arquivo sem recalcular seu hash."
    If Not IsNumeric(we.Range("D12").Value) Then Err.Raise vbObjectError + 215, , "Valor invalido."
    valor = CDbl(we.Range("D12").Value)
    If valor <= 0 Or Abs(valor - WorksheetFunction.Round(valor, 2)) > 0.000001 Then Err.Raise vbObjectError + 216, , "Informe um valor positivo com duas casas decimais."
    mes = Competencia(we.Range("D11").Value): per = CStr(we.Range("D16").Value)
    novoC = ColunaPeriodo(mes, per)
    antigo = CDbl(wc.Cells(r, 5).Value)
    oldControl = wc.Range(wc.Cells(r, 1), wc.Cells(r, 13)).Formula
    saldo = wd.Cells(d, c).Value: saldoNovo = wd.Cells(d, novoC).Value
    gravando = True
    wd.Cells(d, c).Value = WorksheetFunction.Round(CDbl(saldo) - antigo, 2)
    wd.Cells(d, novoC).Value = WorksheetFunction.Round(CDbl(wd.Cells(d, novoC).Value) + valor, 2)
    wc.Cells(r, 4).NumberFormat = "@": wc.Cells(r, 4).Value = mes
    wc.Cells(r, 5).Value = valor: wc.Cells(r, 11).Value = per
    wc.Cells(r, 3).Value = we.Range("D14").Value: wc.Cells(r, 8).Value = we.Range("D15").Value
    ModuloComprovantes.AtualizarCalculos
    Exit Sub
Falha:
    msg = Err.Description
    If gravando Then
        wd.Cells(d, c).Value = saldo: wd.Cells(d, novoC).Value = saldoNovo
        wc.Range(wc.Cells(r, 1), wc.Cells(r, 13)).Formula = oldControl
    End If
    Err.Raise vbObjectError + 217, , msg
End Sub

Public Sub ExcluirControle()
    Dim wc As Worksheet, wd As Worksheet, r As Long, d As Long, c As Long, valor As Double
    Dim oldControl As Variant, antes As Variant, gravando As Boolean, msg As String
    On Error GoTo Falha
    Set wc = ThisWorkbook.Worksheets("Controle"): Set wd = ThisWorkbook.Worksheets("Despesas")
    If Not ActiveSheet Is wc Then Err.Raise vbObjectError + 218, , "Selecione um registro em Controle."
    r = ActiveCell.Row
    ValidarMovimento r, d, c
    If Application.Visible Then
        If MsgBox("Excluir o controle " & wc.Cells(r, 1).Value & " e descontar seu valor?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    oldControl = wc.Range(wc.Cells(r, 1), wc.Cells(r, 13)).Formula
    antes = wd.Cells(d, c).Value: valor = CDbl(wc.Cells(r, 5).Value)
    gravando = True
    wd.Cells(d, c).Value = WorksheetFunction.Round(CDbl(antes) - valor, 2)
    ' Keep row positions and never reuse the monotonic identifier.
    wc.Range(wc.Cells(r, 1), wc.Cells(r, 13)).ClearContents
    ModuloComprovantes.AtualizarCalculos
    AtualizarReferencias
    Exit Sub
Falha:
    msg = Err.Description
    If gravando Then
        wd.Cells(d, c).Value = antes
        wc.Range(wc.Cells(r, 1), wc.Cells(r, 13)).Formula = oldControl
    End If
    Err.Raise vbObjectError + 219, , msg
End Sub

Public Function UltimaLinhaPreenchida(ws As Worksheet, coluna As Long) As Long
    Dim ultima As Range
    ' Unlike End(xlUp), Find also sees records hidden by an active filter.
    Set ultima = ws.Columns(coluna).Find(What:="*", After:=ws.Cells(1, coluna), LookIn:=xlFormulas, _
                                       LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                       MatchCase:=False, SearchFormat:=False)
    If ultima Is Nothing Then UltimaLinhaPreenchida = 1 Else UltimaLinhaPreenchida = ultima.Row
End Function

Public Sub SincronizarSequencias()
    Dim wc As Worksheet, wd As Worksheet, wr As Worksheet, cfg As Worksheet
    Dim r As Long, maiorControle As Long, maiorDespesa As Long, id As String
    Set wc = ThisWorkbook.Worksheets("Controle")
    Set wd = ThisWorkbook.Worksheets("Despesas")
    Set wr = ThisWorkbook.Worksheets("Cadastros")
    Set cfg = ThisWorkbook.Worksheets("Config")
    maiorControle = CLng(Val(CStr(cfg.Range("F10").Value)))
    maiorDespesa = CLng(Val(CStr(cfg.Range("F11").Value)))
    For r = 4 To UltimaLinhaPreenchida(wc, 1)
        If IsNumeric(wc.Cells(r, 1).Value) Then maiorControle = Application.Max(maiorControle, Val(CStr(wc.Cells(r, 1).Value)))
    Next r
    For r = 5 To UltimaLinhaPreenchida(wd, 24)
        id = CStr(wd.Cells(r, 24).Value)
        If id Like "D######" Then maiorDespesa = Application.Max(maiorDespesa, CLng(Mid$(id, 2)))
    Next r
    For r = 13 To UltimaLinhaPreenchida(wr, 6)
        id = CStr(wr.Cells(r, 6).Value)
        If id Like "D######" Then maiorDespesa = Application.Max(maiorDespesa, CLng(Mid$(id, 2)))
    Next r
    cfg.Range("F10").Value = maiorControle
    cfg.Range("F11").Value = maiorDespesa
End Sub
