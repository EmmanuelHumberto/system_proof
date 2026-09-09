' ============================================================================
' USERFORM DE IMPORTACAO DE COMPROVANTES (frmImportar)
' ============================================================================
' Controles esperados:
'   lstComprovantes, txtData, txtValor, txtPagador, txtRecebedor, txtTipo,
'   cboCategoria, cboDespesa, cboPeriodicidade, btnInserir, btnPular, lblStatus
' ============================================================================

Option Explicit

Private comprovantes As Collection
Private total As Long
Private idxAtual As Long

Private Sub UserForm_Initialize()
    If Not CarregarCSV Then
        MsgBox "CSV nao encontrado ou vazio em Config!B7.", vbExclamation
        Unload Me
        Exit Sub
    End If
    PreencherCategorias
    cboPeriodicidade.List = Array("Mensal", "Anual", "Eventual")
    cboPeriodicidade.Value = "Mensal"
    idxAtual = 1
    MostrarItem idxAtual
End Sub

Private Function CarregarCSV() As Boolean
    Dim caminho As String, linha As String, f As Integer, arr As Variant
    Set comprovantes = New Collection
    caminho = CaminhoBase() & Application.PathSeparator & "comprovantes.csv"
    If Dir(caminho) = "" Then Exit Function

    f = FreeFile
    Open caminho For Input As #f
    If Not EOF(f) Then Line Input #f, linha
    Do While Not EOF(f)
        Line Input #f, linha
        If Trim(linha) <> "" Then
            arr = ParseCsvSemicolon(linha)
            If UBound(arr) >= 11 Then comprovantes.Add arr
        End If
    Loop
    Close #f
    total = comprovantes.Count
    CarregarCSV = (total > 0)
End Function

Private Sub PreencherCategorias()
    Dim i As Long, col As Collection, v As Variant, linha As Variant
    Set col = New Collection
    On Error Resume Next
    For i = 1 To total
        linha = comprovantes(i)
        col.Add linha(1), linha(1)
    Next i
    On Error GoTo 0
    For Each v In col
        cboCategoria.AddItem v
    Next v

    lstComprovantes.ColumnCount = 4
    lstComprovantes.ColumnWidths = "40;90;80;180"
    For i = 1 To total
        linha = comprovantes(i)
        lstComprovantes.AddItem linha(0)
        lstComprovantes.List(i - 1, 1) = linha(3)
        lstComprovantes.List(i - 1, 2) = linha(5)
        lstComprovantes.List(i - 1, 3) = linha(7)
    Next i
End Sub

Private Sub MostrarItem(i As Long)
    Dim linha As Variant
    linha = comprovantes(i)
    txtData.Text = linha(3)
    txtValor.Text = linha(5)
    txtPagador.Text = linha(6)
    txtRecebedor.Text = linha(7)
    txtTipo.Text = linha(8)
    cboCategoria.Value = linha(1)
    cboDespesa.Value = linha(2)
    cboPeriodicidade.Value = linha(11)
    If cboPeriodicidade.Value = "" Then cboPeriodicidade.Value = "Mensal"
    lblStatus.Caption = i & " de " & total
    lstComprovantes.ListIndex = i - 1
End Sub

Private Sub cboCategoria_Change()
    cboDespesa.Clear
    Select Case cboCategoria.Value
        Case "Alimentação"
            cboDespesa.AddItem "Supermercado e alimentação domiciliar"
            cboDespesa.AddItem "Suplementos e alimentação especial"
            cboDespesa.AddItem "Alimentação escolar / cantina"
        Case "Moradia"
            cboDespesa.AddItem "Condomínio – quota-parte"
            cboDespesa.AddItem "IPTU – quota-parte"
            cboDespesa.AddItem "Energia elétrica – quota-parte"
            cboDespesa.AddItem "Água e esgoto – quota-parte"
            cboDespesa.AddItem "Gás – quota-parte"
            cboDespesa.AddItem "Internet residencial – quota-parte"
            cboDespesa.AddItem "Aluguel ou financiamento – quota-parte do adolescente"
            cboDespesa.AddItem "Outras despesas comprovadas"
        Case "Saúde"
            cboDespesa.AddItem "Medicamentos contínuos"
            cboDespesa.AddItem "Consultas médicas particulares"
            cboDespesa.AddItem "Tratamento odontológico"
            cboDespesa.AddItem "Exames médicos"
            cboDespesa.AddItem "Óculos, lentes e itens ortopédicos"
            cboDespesa.AddItem "Psicólogo / psiquiatra"
        Case "Transporte"
            cboDespesa.AddItem "Aplicativos de transporte"
            cboDespesa.AddItem "Transporte escolar"
            cboDespesa.AddItem "Passagens de transporte coletivo"
            cboDespesa.AddItem "Combustível para deslocamentos"
        Case "Educação"
            cboDespesa.AddItem "Mensalidade escolar"
            cboDespesa.AddItem "Reforço escolar"
            cboDespesa.AddItem "Curso preparatório / profissionalizante"
            cboDespesa.AddItem "Livros didáticos e paradidáticos"
            cboDespesa.AddItem "Cursos de música, arte ou cultura"
            cboDespesa.AddItem "Matrícula escolar"
        Case "Vestuário e higiene"
            cboDespesa.AddItem "Roupas"
            cboDespesa.AddItem "Calçados"
            cboDespesa.AddItem "Cabeleireiro / cuidados pessoais"
            cboDespesa.AddItem "Produtos de higiene pessoal"
        Case "Lazer e convivência"
            cboDespesa.AddItem "Assinaturas de streaming – quota-parte"
            cboDespesa.AddItem "Cinema, parques e passeios"
            cboDespesa.AddItem "Viagens e férias"
        Case "Comunicação e tecnologia"
            cboDespesa.AddItem "Plano de telefonia celular"
            cboDespesa.AddItem "Aparelho celular"
        Case Else
            cboDespesa.AddItem "Outras despesas comprovadas"
    End Select
End Sub

Private Sub btnInserir_Click()
    InserirNaPlanilha
    idxAtual = idxAtual + 1
    If idxAtual > total Then
        MsgBox "Todos os comprovantes foram processados.", vbInformation
        Unload Me
    Else
        MostrarItem idxAtual
    End If
End Sub

Private Sub btnPular_Click()
    idxAtual = idxAtual + 1
    If idxAtual > total Then Unload Me Else MostrarItem idxAtual
End Sub

Private Sub InserirNaPlanilha()
    Dim wsC As Worksheet, wsD As Worksheet, rC As Long, rD As Long, colMes As Long
    Dim mes As String, valor As Double, anual As Boolean, linha As Variant, hash As String

    Set wsC = ThisWorkbook.Worksheets("Controle")
    Set wsD = ThisWorkbook.Worksheets("Despesas")
    linha = comprovantes(idxAtual)
    hash = Trim(CStr(linha(10)))
    If hash <> "" And HashJaImportado(wsC, hash) Then
        MsgBox "Comprovante ja importado. Hash: " & hash, vbExclamation
        Exit Sub
    End If

    If Not TryParseDouble(txtValor.Text, valor) Then
        MsgBox "Valor invalido: " & txtValor.Text, vbExclamation
        Exit Sub
    End If
    mes = Left(txtData.Text, 7)
    anual = (cboPeriodicidade.Value <> "Mensal")

    rD = AcharLinhaDespesa(cboDespesa.Value)
    If rD = 0 Then
        MsgBox "Despesa nao localizada: " & cboDespesa.Value, vbExclamation
        Exit Sub
    End If

    wsC.Cells(3, 9).Value = "Hash"
    wsC.Columns(9).Hidden = True
    rC = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row + 1
    If rC < 4 Then rC = 4
    wsC.Cells(rC, 1).Value = rC - 3
    wsC.Cells(rC, 2).Value = cboCategoria.Value
    wsC.Cells(rC, 3).Value = txtRecebedor.Text
    wsC.Cells(rC, 4).Value = mes
    wsC.Cells(rC, 5).Value = valor
    wsC.Cells(rC, 6).Value = "Anexado"
    wsC.Cells(rC, 7).Value = linha(9)
    If linha(9) <> "" Then
        wsC.Hyperlinks.Add Anchor:=wsC.Cells(rC, 7), Address:=ResolverCaminhoArquivo(CStr(linha(9))), TextToDisplay:=CStr(linha(9))
    End If
    wsC.Cells(rC, 8).Value = txtTipo.Text & " - " & txtData.Text & " - pagador: " & txtPagador.Text
    wsC.Cells(rC, 9).Value = hash

    If anual Then
        wsD.Cells(rD, ColunaPorCabecalho(wsD, "Valor anual/eventual")).Value = NzD(wsD.Cells(rD, ColunaPorCabecalho(wsD, "Valor anual/eventual")).Value) + valor
    Else
        colMes = AcharOuCriarColunaMes(mes)
        wsD.Cells(rD, colMes).Value = NzD(wsD.Cells(rD, colMes).Value) + valor
    End If
    wsD.Cells(rD, ColunaPorCabecalho(wsD, "Comprovante")).Value = "Anexado"
    AtualizarFormulasDespesas wsD
End Sub

Private Function CaminhoBase() As String
    Dim p As String
    p = Trim(CStr(ThisWorkbook.Worksheets("Config").Range("B7").Value))
    If p = "" Then p = ThisWorkbook.Path
    CaminhoBase = p
End Function

Private Function ResolverCaminhoArquivo(relOuAbs As String) As String
    If relOuAbs Like "[A-Za-z]:\*" Or Left$(relOuAbs, 1) = "/" Or Left$(relOuAbs, 2) = "\\" Then
        ResolverCaminhoArquivo = relOuAbs
    Else
        ResolverCaminhoArquivo = CaminhoBase() & Application.PathSeparator & relOuAbs
    End If
End Function

Private Function HashJaImportado(ws As Worksheet, hash As String) As Boolean
    Dim ultima As Long, r As Long
    ultima = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    For r = 4 To ultima
        If Trim(CStr(ws.Cells(r, 9).Value)) = hash Then
            HashJaImportado = True
            Exit Function
        End If
    Next r
End Function

Private Function AcharLinhaDespesa(desp As String) As Long
    Dim r As Long
    With ThisWorkbook.Worksheets("Despesas")
        For r = 5 To 57
            If Trim(CStr(.Cells(r, 2).Value)) = Trim(desp) Then
                AcharLinhaDespesa = r
                Exit Function
            End If
        Next r
    End With
End Function

Private Function AcharOuCriarColunaMes(mes As String) As Long
    Dim ws As Worksheet, c As Long, alvo As String, colTotal As Long
    Set ws = ThisWorkbook.Worksheets("Despesas")
    colTotal = ColunaPorCabecalho(ws, "Total")
    alvo = "mes " & NomeMesPorExtenso(Right$(mes, 2)) & "/" & Left$(mes, 4)
    For c = 3 To colTotal - 1
        If NormalizarTexto(CStr(ws.Cells(4, c).Value)) = NormalizarTexto(alvo) Then
            AcharOuCriarColunaMes = c
            Exit Function
        End If
    Next c
    ws.Columns(colTotal).Insert Shift:=xlToRight
    ws.Cells(4, colTotal).Value = "mês " & NomeMesPorExtenso(Right$(mes, 2)) & "/" & Left$(mes, 4)
    AcharOuCriarColunaMes = colTotal
End Function

Private Sub AtualizarFormulasDespesas(ws As Worksheet)
    Dim r As Long, colTotal As Long, colAnual As Long, colMedia As Long
    colTotal = ColunaPorCabecalho(ws, "Total")
    colAnual = ColunaPorCabecalho(ws, "Valor anual/eventual")
    colMedia = ColunaPorCabecalho(ws, "Média mensal")
    For r = 5 To 57
        ws.Cells(r, colTotal).Formula = "=SUM(C" & r & ":" & ws.Cells(r, colTotal - 1).Address(False, False) & ")"
        ws.Cells(r, colMedia).Formula = "=MediaMeses(C" & r & ":" & ws.Cells(r, colTotal - 1).Address(False, False) & ")+" & ws.Cells(r, colAnual).Address(False, False) & "/12"
    Next r
    ws.Cells(59, colMedia).Formula = "=SUM(" & ws.Cells(5, colMedia).Address(False, False) & ":" & ws.Cells(57, colMedia).Address(False, False) & ")"
End Sub

Private Function ColunaPorCabecalho(ws As Worksheet, prefixo As String) As Long
    Dim c As Long, alvo As String, atual As String
    alvo = NormalizarTexto(prefixo)
    For c = 1 To ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
        atual = NormalizarTexto(CStr(ws.Cells(4, c).Value))
        If Left$(atual, Len(alvo)) = alvo Then
            ColunaPorCabecalho = c
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 1000, , "Cabecalho nao encontrado: " & prefixo
End Function

Private Function TryParseDouble(texto As String, ByRef valor As Double) As Boolean
    On Error GoTo Falha
    valor = CDbl(Replace(Trim(texto), ".", Application.DecimalSeparator))
    TryParseDouble = True
    Exit Function
Falha:
    TryParseDouble = False
End Function

Private Function NzD(v As Variant) As Double
    If IsNumeric(v) Then NzD = CDbl(v) Else NzD = 0
End Function

Private Function ParseCsvSemicolon(linha As String) As Variant
    Dim campos() As String, i As Long, ch As String, atual As String, emAspas As Boolean, n As Long
    ReDim campos(0 To 0)
    For i = 1 To Len(linha)
        ch = Mid$(linha, i, 1)
        If ch = Chr(34) Then
            If emAspas And i < Len(linha) And Mid$(linha, i + 1, 1) = Chr(34) Then
                atual = atual & Chr(34)
                i = i + 1
            Else
                emAspas = Not emAspas
            End If
        ElseIf ch = ";" And Not emAspas Then
            campos(n) = atual
            n = n + 1
            ReDim Preserve campos(0 To n)
            atual = ""
        Else
            atual = atual & ch
        End If
    Next i
    campos(n) = atual
    ParseCsvSemicolon = campos
End Function

Private Function NormalizarTexto(s As String) As String
    s = LCase$(Trim(s))
    s = Replace(s, "ê", "e")
    s = Replace(s, "é", "e")
    s = Replace(s, "è", "e")
    s = Replace(s, "á", "a")
    s = Replace(s, "à", "a")
    s = Replace(s, "ã", "a")
    s = Replace(s, "ç", "c")
    NormalizarTexto = s
End Function

Private Function NomeMesPorExtenso(n As String) As String
    Select Case n
        Case "01": NomeMesPorExtenso = "janeiro"
        Case "02": NomeMesPorExtenso = "fevereiro"
        Case "03": NomeMesPorExtenso = "março"
        Case "04": NomeMesPorExtenso = "abril"
        Case "05": NomeMesPorExtenso = "maio"
        Case "06": NomeMesPorExtenso = "junho"
        Case "07": NomeMesPorExtenso = "julho"
        Case "08": NomeMesPorExtenso = "agosto"
        Case "09": NomeMesPorExtenso = "setembro"
        Case "10": NomeMesPorExtenso = "outubro"
        Case "11": NomeMesPorExtenso = "novembro"
        Case "12": NomeMesPorExtenso = "dezembro"
    End Select
End Function
