Attribute VB_Name = "ModuloComprovantes"
' ============================================================================
' MODULO DE AUTOMACAO - PLANILHA DE COMPROVANTES (Pensao Alimenticia)
'
' Entrada oficial: comprovantes.json em Config!B7.
' Fluxo:
'   1. ExtrairDados gera/enfileira o JSON, sem preencher a planilha.
'   2. CarregarFilaJson carrega o JSON na aba Fila para revisao.
'   3. ImportarSelecionadoDaFila ou ImportarTodosPendentes grava na planilha.
' ============================================================================

Option Explicit

Private Const LINHA_CABECALHO As Long = 4
Private Const LINHA_INICIAL_DADOS As Long = 5
Private Const LINHA_INICIAL_CONTROLE As Long = 4
Private Const LINHA_FINAL_DADOS As Long = 120
Private Const LINHA_TOTAL As Long = 122
Private Const COL_COMPROVANTE As Long = 21
Private Const COL_OBSERVACOES As Long = 22
Private Const COL_HASH_CONTROLE As Long = 9
Private Const NOME_ABA_FORM As String = "Editar"
Private Const NOME_ABA_CADASTROS As String = "Cadastros"
Private Const SENHA_PROTECAO As String = "daisy2026"
Private Const COL_MES_INICIAL As Long = 3
Private Const COL_MES_FINAL As Long = 14
Private Const COL_EVENTUAL_ANUAL As Long = 15
Private Const COL_VALOR_MES As Long = 16
Private Const COL_VALOR_ANUAL As Long = 17
Private Const COL_COTA_PARTE As Long = 18
Private Const COL_VALOR_COTA_PARTE_ANUAL As Long = 19
Private Const COL_VALOR_COTA_PARTE As Long = 20

Public Function MediaMeses(rng As Range) As Double
    Dim c As Range, soma As Double, n As Long
    soma = 0
    n = 0
    For Each c In rng
        If IsNumeric(c.Value) Then
            If CDbl(c.Value) > 0 Then
                soma = soma + CDbl(c.Value)
                n = n + 1
            End If
        End If
    Next c
    If n > 0 Then MediaMeses = soma / n Else MediaMeses = 0
End Function

Public Sub AbrirComprovante()
    NormalizarJanela
    Dim ws As Worksheet, p As String, r As Long
    Set ws = GarantirAbaFila()

    If ActiveSheet.Name = "Fila" And ActiveCell.Row >= 2 Then
        r = ActiveCell.Row
        p = Trim(CStr(ws.Cells(r, 11).Value))
    ElseIf ActiveSheet.Name = "Controle" And ActiveCell.Row >= 4 Then
        r = ActiveCell.Row
        p = Trim(CStr(ActiveSheet.Cells(r, 7).Value))
    Else
        ws.Activate
        Aviso "Na aba Fila, clique diretamente no link da coluna arquivo ou selecione uma linha e use Abrir Comprovante.", vbInformation
        Exit Sub
    End If

    If p = "" Then
        Aviso "A linha selecionada nao tem caminho de comprovante.", vbExclamation
        Exit Sub
    End If
    p = ResolverCaminhoArquivo(p)
    If Dir(p) <> "" Then
        ThisWorkbook.FollowHyperlink p
    Else
        Aviso "Arquivo nao encontrado:" & vbLf & p, vbExclamation
    End If
End Sub

Public Sub ExtrairDados()
    NormalizarJanela
    Dim base As String, script As String, launcher As String, scriptShell As String, cmd As String
    base = CaminhoBase()
    script = base & Application.PathSeparator & "_scripts" & Application.PathSeparator & "integrar_extracao_planilha.py"
    launcher = base & Application.PathSeparator & "integrar_extracao.bat"
    scriptShell = CaminhoShell(script)
    On Error GoTo FalhaExtracao
    If EhWindows() Then
        cmd = "cmd /c " & AspasShell(launcher)
    Else
        cmd = "bash -lc " & AspasShell("exec python3 " & Chr(39) & scriptShell & Chr(39))
    End If
    Shell cmd, vbNormalFocus
    Aviso "Extracao iniciada. Aguarde o Python terminar e depois carregue a fila JSON.", vbInformation
    Exit Sub
FalhaExtracao:
    Aviso "Nao foi possivel iniciar o extrator:" & vbLf & scriptShell & vbLf & Err.Description, vbExclamation
End Sub

Public Sub AbrirFerramentaExtracao()
    NormalizarJanela
    Dim base As String, script As String, launcher As String, scriptShell As String, cmd As String
    base = CaminhoBase()
    script = base & Application.PathSeparator & "_scripts" & Application.PathSeparator & "extrair_gui.py"
    launcher = base & Application.PathSeparator & "extrair_gui.bat"
    scriptShell = CaminhoShell(script)

    If EhWindows() Then
        cmd = "cmd /c " & AspasShell(launcher)
    Else
        cmd = "bash -lc " & AspasShell("exec python3 " & Chr(39) & scriptShell & Chr(39))
    End If
    On Error GoTo FalhaFerramenta
    Shell cmd, vbNormalFocus
    Exit Sub
FalhaFerramenta:
    Aviso "Nao foi possivel iniciar a ferramenta de extracao:" & vbLf & scriptShell & vbLf & Err.Description, vbExclamation
End Sub

Private Function EhWindows() As Boolean
    EhWindows = InStr(1, Application.OperatingSystem, "Windows", vbTextCompare) > 0
End Function

Private Function AspasShell(valor As String) As String
    AspasShell = Chr(34) & valor & Chr(34)
End Function

Public Sub ImportarComprovantes()
    ImportarTodosPendentes
End Sub

Public Sub IrParaFila()
    ThisWorkbook.Worksheets("Fila").Activate
End Sub

Public Sub IrParaControle()
    ThisWorkbook.Worksheets("Controle").Activate
End Sub

Public Sub IrParaDespesas()
    CongelarReferenciasDespesas
End Sub

Public Sub CarregarFilaJson()
    NormalizarJanela
    Dim caminhoJson As String, texto As String, itens As Collection, item As Object
    Dim ws As Worksheet, wsC As Worksheet, r As Long
    Dim hashItem As String, carregados As Long, jaImportados As Long

    caminhoJson = CaminhoBase() & Application.PathSeparator & "comprovantes.json"
    If Dir(caminhoJson) = "" Then
        Aviso "JSON nao encontrado:" & vbLf & caminhoJson, vbExclamation
        Exit Sub
    End If

    texto = LerArquivoTexto(caminhoJson)
    Set itens = ParseComprovantesJson(texto)
    Set ws = GarantirAbaFila()
    Set wsC = ThisWorkbook.Worksheets("Controle")
    DesprotegerPlanilhas
    GarantirCabecalhoControle wsC
    LimparFila ws
    EscreverCabecalhoFila ws

    r = 2
    For Each item In itens
        hashItem = Trim(CStr(ValorDict(item, "hash")))
        If hashItem <> "" And HashJaImportado(wsC, hashItem) Then
            jaImportados = jaImportados + 1
        Else
            ws.Cells(r, 1).Value = ValorDict(item, "id")
            ws.Cells(r, 2).Value = "pendente"
            ws.Cells(r, 3).Value = ValorDict(item, "categoria")
            ws.Cells(r, 4).Value = ValorDict(item, "despesa")
            ws.Cells(r, 5).Value = ValorDict(item, "data")
            ws.Cells(r, 6).Value = ValorDict(item, "mes")
            ws.Cells(r, 7).Value = ValorDict(item, "valor")
            ws.Cells(r, 8).Value = ValorDict(item, "pagador")
            ws.Cells(r, 9).Value = ValorDict(item, "recebedor")
            ws.Cells(r, 10).Value = ValorDict(item, "tipo")
            ws.Cells(r, 11).Value = ValorDict(item, "arquivo")
            If Trim(CStr(ws.Cells(r, 11).Value)) <> "" Then ws.Hyperlinks.Add Anchor:=ws.Cells(r, 11), Address:=ResolverCaminhoArquivo(CStr(ws.Cells(r, 11).Value)), TextToDisplay:=CStr(ws.Cells(r, 11).Value)
            ws.Cells(r, 12).Value = hashItem
            ws.Cells(r, 13).Value = ValorDict(item, "periodicidade")
            If ws.Cells(r, 13).Value = "" Then ws.Cells(r, 13).Value = "Mensal"
            carregados = carregados + 1
            r = r + 1
        End If
    Next item

    ws.Columns("A:N").AutoFit
    AplicarFormatacaoPadrao
    ProtegerPlanilhas
    Aviso "Fila carregada." & vbLf & "Pendentes: " & carregados & vbLf & "Ja estavam no Controle e foram descartados: " & jaImportados, vbInformation
End Sub

Public Sub ImportarSelecionadoDaFila()
    NormalizarJanela
    Dim ws As Worksheet, r As Long, ultima As Long
    Dim pendentes As Long, primeiraPendente As Long
    Set ws = GarantirAbaFila()

    If ActiveSheet.Name = ws.Name And ActiveCell.Row >= 2 And Trim(CStr(ws.Cells(ActiveCell.Row, 1).Value)) <> "" Then
        r = ActiveCell.Row
    Else
        ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For r = 2 To ultima
            If LCase$(Trim(CStr(ws.Cells(r, 2).Value))) = "pendente" Then
                pendentes = pendentes + 1
                If primeiraPendente = 0 Then primeiraPendente = r
            End If
        Next r
        If pendentes = 1 Then
            r = primeiraPendente
            ws.Activate
            ws.Rows(r).Select
        Else
            ws.Activate
            Aviso "Selecione uma linha pendente na aba Fila. Pendentes encontrados: " & pendentes, vbExclamation
            Exit Sub
        End If
    End If

    If LCase$(Trim(CStr(ws.Cells(r, 2).Value))) <> "pendente" Then
        Aviso "A linha selecionada nao esta pendente. Status atual: " & ws.Cells(r, 2).Value, vbExclamation
        Exit Sub
    End If
    DesprotegerPlanilhas
    If Not ImportarLinhaFila(ws, r) Then
        ProtegerPlanilhas
        Aviso "Nao foi possivel importar a linha selecionada. Veja a coluna Resultado na aba Fila.", vbExclamation
        Exit Sub
    End If
    ws.Rows(r).Delete Shift:=xlUp
    AtualizarFormulasDespesas ThisWorkbook.Worksheets("Despesas")
    AtualizarResumoRateio
    AplicarFormatacaoPadrao
    ProtegerPlanilhas
    Aviso "Linha importada com sucesso e removida da Fila.", vbInformation
End Sub

Public Sub ImportarTodosPendentes()
    NormalizarJanela
    Dim ws As Worksheet, r As Long, ultima As Long
    Dim importados As Long, pulados As Long
    Set ws = GarantirAbaFila()
    If ws.Cells(1, 1).Value = "" Then CarregarFilaJson

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    DesprotegerPlanilhas
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    For r = ultima To 2 Step -1
        If LCase$(Trim(CStr(ws.Cells(r, 2).Value))) = "pendente" Then
            If ImportarLinhaFila(ws, r) Then
                importados = importados + 1
                ws.Rows(r).Delete Shift:=xlUp
            Else
                pulados = pulados + 1
            End If
        End If
    Next r
    AtualizarFormulasDespesas ThisWorkbook.Worksheets("Despesas")
    AtualizarResumoRateio
    AplicarFormatacaoPadrao
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ProtegerPlanilhas
    Aviso "Importacao da fila concluida." & vbLf & "Importados e removidos da Fila: " & importados & vbLf & "Mantidos para revisao: " & pulados, vbInformation
End Sub

Private Function ImportarLinhaFila(wsF As Worksheet, rF As Long) As Boolean
    Dim wsC As Worksheet, wsD As Worksheet
    Dim rC As Long, rD As Long, colMes As Long, valor As Double
    Dim cat As String, desp As String, dt As String, mes As String
    Dim pagador As String, recebedor As String, tipo As String, arquivo As String, hash As String, periodicidade As String

    On Error GoTo FalhaLinha
    Set wsC = ThisWorkbook.Worksheets("Controle")
    Set wsD = ThisWorkbook.Worksheets("Despesas")
    DesprotegerPlanilhas
    GarantirCabecalhoControle wsC

    cat = Trim(CStr(wsF.Cells(rF, 3).Value))
    desp = Trim(CStr(wsF.Cells(rF, 4).Value))
    dt = Trim(CStr(wsF.Cells(rF, 5).Value))
    mes = Trim(CStr(wsF.Cells(rF, 6).Value))
    pagador = Trim(CStr(wsF.Cells(rF, 8).Value))
    recebedor = Trim(CStr(wsF.Cells(rF, 9).Value))
    tipo = Trim(CStr(wsF.Cells(rF, 10).Value))
    arquivo = Trim(CStr(wsF.Cells(rF, 11).Value))
    hash = Trim(CStr(wsF.Cells(rF, 12).Value))
    periodicidade = Trim(CStr(wsF.Cells(rF, 13).Value))
    If periodicidade = "" Then periodicidade = "Mensal"

    If Not TryParseDouble(CStr(wsF.Cells(rF, 7).Value), valor) Then
        wsF.Cells(rF, 2).Value = "erro"
        wsF.Cells(rF, 14).Value = "Valor invalido"
        Exit Function
    End If
    If hash <> "" And HashJaImportado(wsC, hash) Then
        wsF.Cells(rF, 2).Value = "duplicado"
        wsF.Cells(rF, 14).Value = "Hash ja importado"
        Exit Function
    End If
    rD = AcharLinhaDespesa(desp, cat)
    If rD = 0 Then
        wsF.Cells(rF, 2).Value = "erro"
        wsF.Cells(rF, 14).Value = "Despesa nao localizada"
        Exit Function
    End If

    If UCase$(periodicidade) = "MENSAL" And mes = "" Then
        wsF.Cells(rF, 2).Value = "erro"
        wsF.Cells(rF, 14).Value = "Mes vazio"
        Exit Function
    End If

    rC = ProximaLinhaControle(wsC)
    wsC.Cells(rC, 1).Value = rC - 3
    wsC.Cells(rC, 2).Value = cat
    wsC.Cells(rC, 3).Value = recebedor
    wsC.Cells(rC, 4).Value = mes
    wsC.Cells(rC, 5).Value = valor
    wsC.Cells(rC, 6).Value = "Anexado"
    wsC.Cells(rC, 7).Value = arquivo
    If arquivo <> "" Then wsC.Hyperlinks.Add Anchor:=wsC.Cells(rC, 7), Address:=ResolverCaminhoArquivo(arquivo), TextToDisplay:=arquivo
    wsC.Cells(rC, 8).Value = tipo & " - " & dt & " - pagador: " & pagador
    wsC.Cells(rC, COL_HASH_CONTROLE).Value = hash

    If UCase$(periodicidade) = "MENSAL" Then
        colMes = AcharOuCriarColunaMes(mes)
        wsD.Cells(rD, colMes).Value = NzD(wsD.Cells(rD, colMes).Value) + valor
    Else
        wsD.Cells(rD, COL_EVENTUAL_ANUAL).Value = NzD(wsD.Cells(rD, COL_EVENTUAL_ANUAL).Value) + valor
    End If

    wsD.Cells(rD, COL_COMPROVANTE).Value = "Anexado"
    AtualizarObservacaoComLink wsD, rD, rC - 3, arquivo
    wsF.Cells(rF, 2).Value = "importado"
    wsF.Cells(rF, 14).Value = "Controle nº " & (rC - 3)
    ImportarLinhaFila = True
    Exit Function
FalhaLinha:
    wsF.Cells(rF, 2).Value = "erro"
    wsF.Cells(rF, 14).Value = Err.Description
    ImportarLinhaFila = False
End Function

Public Sub ImportarComprovantesCsv()
    Dim caminhoCsv As String, linha As String, campos As Variant
    Dim f As Integer, linhaCsv As Long
    Dim wsC As Worksheet, wsD As Worksheet
    Dim rC As Long, rD As Long, colMes As Long
    Dim cat As String, desp As String, dt As String, mes As String
    Dim valor As Double, pagador As String, recebedor As String, tipo As String
    Dim arquivo As String, hash As String, periodicidade As String
    Dim importados As Long, duplicados As Long, ignorados As Long, erros As String

    caminhoCsv = CaminhoBase() & Application.PathSeparator & "comprovantes.csv"
    If Dir(caminhoCsv) = "" Then
        Aviso "CSV nao encontrado:" & vbLf & caminhoCsv, vbExclamation
        Exit Sub
    End If

    Set wsC = ThisWorkbook.Worksheets("Controle")
    Set wsD = ThisWorkbook.Worksheets("Despesas")
    DesprotegerPlanilhas
    GarantirCabecalhoControle wsC
    AtualizarFormulasDespesas wsD
    AtualizarResumoRateio

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    f = FreeFile
    Open caminhoCsv For Input As #f
    If Not EOF(f) Then Line Input #f, linha
    linhaCsv = 1

    Do While Not EOF(f)
        Line Input #f, linha
        linhaCsv = linhaCsv + 1
        If Trim(linha) <> "" Then
            campos = ParseCsvSemicolon(linha)
            If UBound(campos) < 11 Then
                ignorados = ignorados + 1
                erros = erros & "Linha " & linhaCsv & ": colunas insuficientes." & vbLf
                GoTo ProximaLinha
            End If

            cat = Trim(CStr(campos(1)))
            desp = Trim(CStr(campos(2)))
            dt = Trim(CStr(campos(3)))
            mes = Trim(CStr(campos(4)))
            pagador = Trim(CStr(campos(6)))
            recebedor = Trim(CStr(campos(7)))
            tipo = Trim(CStr(campos(8)))
            arquivo = Trim(CStr(campos(9)))
            hash = Trim(CStr(campos(10)))
            periodicidade = Trim(CStr(campos(11)))
            If periodicidade = "" Then periodicidade = "Mensal"

            If Not TryParseDouble(CStr(campos(5)), valor) Then
                ignorados = ignorados + 1
                erros = erros & "Linha " & linhaCsv & ": valor invalido." & vbLf
                GoTo ProximaLinha
            End If

            If hash <> "" And HashJaImportado(wsC, hash) Then
                duplicados = duplicados + 1
                GoTo ProximaLinha
            End If

            rD = AcharLinhaDespesa(desp, cat)
            If rD = 0 Then
                ignorados = ignorados + 1
                erros = erros & "Linha " & linhaCsv & ": despesa nao localizada: " & desp & vbLf
                GoTo ProximaLinha
            End If

            rC = ProximaLinhaControle(wsC)
            wsC.Cells(rC, 1).Value = rC - 3
            wsC.Cells(rC, 2).Value = cat
            wsC.Cells(rC, 3).Value = recebedor
            wsC.Cells(rC, 4).Value = mes
            wsC.Cells(rC, 5).Value = valor
            wsC.Cells(rC, 6).Value = "Anexado"
            wsC.Cells(rC, 7).Value = arquivo
            If arquivo <> "" Then
                wsC.Hyperlinks.Add Anchor:=wsC.Cells(rC, 7), Address:=ResolverCaminhoArquivo(arquivo), TextToDisplay:=arquivo
            End If
            wsC.Cells(rC, 8).Value = tipo & " - " & dt & " - pagador: " & pagador
            wsC.Cells(rC, COL_HASH_CONTROLE).Value = hash

            If UCase$(periodicidade) = "MENSAL" Then
                colMes = AcharOuCriarColunaMes(mes)
                wsD.Cells(rD, colMes).Value = NzD(wsD.Cells(rD, colMes).Value) + valor
            Else
                wsD.Cells(rD, COL_EVENTUAL_ANUAL).Value = NzD(wsD.Cells(rD, COL_EVENTUAL_ANUAL).Value) + valor
            End If
            wsD.Cells(rD, COL_COMPROVANTE).Value = "Anexado"
            AtualizarObservacaoComLink wsD, rD, rC - 3, arquivo
            importados = importados + 1
        End If
ProximaLinha:
    Loop
    Close #f

    AtualizarFormulasDespesas wsD
    AtualizarResumoRateio
    AplicarFormatacaoPadrao
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ProtegerPlanilhas

    Aviso "Importacao concluida." & vbLf & _
           "Importados: " & importados & vbLf & _
           "Duplicados ignorados: " & duplicados & vbLf & _
           "Linhas ignoradas: " & ignorados & IIf(erros <> "", vbLf & vbLf & erros, ""), vbInformation
End Sub

Public Sub AbrirRegistrosDaObservacao(cel As Range)
    Dim wsC As Worksheet, numeros As Collection, criterios() As String
    Dim i As Long, ultima As Long, textoObs As String
    If cel.Column = COL_COMPROVANTE Then
        textoObs = CStr(cel.Worksheet.Cells(cel.Row, COL_OBSERVACOES).Value)
    Else
        textoObs = CStr(cel.Value)
    End If
    Set numeros = ExtrairNumerosControle(textoObs)
    Set wsC = ThisWorkbook.Worksheets("Controle")
    wsC.Activate
    ultima = UltimaLinhaControle(wsC)
    If wsC.AutoFilterMode Then wsC.AutoFilterMode = False
    If numeros.Count = 0 Or ultima < 4 Then Exit Sub
    ReDim criterios(0 To numeros.Count - 1)
    For i = 1 To numeros.Count
        criterios(i - 1) = CStr(numeros(i))
    Next i
    AtualizarTotalControle wsC
    wsC.Range("A3:I" & ultima).AutoFilter Field:=1, Criteria1:=criterios, Operator:=xlFilterValues
End Sub

Private Function ExtrairNumerosControle(texto As String) As Collection
    Dim re As Object, matches As Object, m As Object
    Dim col As New Collection
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "\d+"
    If re.Test(texto) Then
        Set matches = re.Execute(texto)
        For Each m In matches
            col.Add CLng(m.Value)
        Next m
    End If
    Set ExtrairNumerosControle = col
End Function

Public Sub LimparFiltroControle()
    With ThisWorkbook.Worksheets("Controle")
        If .AutoFilterMode Then .AutoFilterMode = False
        AtualizarTotalControle ThisWorkbook.Worksheets("Controle")
        .Activate
    End With
End Sub
Public Sub ExcluirRegistroControleSelecionado()
    Dim wsC As Worksheet, wsD As Worksheet, linhaControle As Long, numeroControle As Long
    Dim valor As Double, mes As String, linhaDespesa As Long, hash As String, arquivo As String
    Set wsC = ThisWorkbook.Worksheets("Controle")
    Set wsD = ThisWorkbook.Worksheets("Despesas")

    If ActiveSheet.Name <> "Controle" Or ActiveCell.Row < LINHA_INICIAL_CONTROLE Then
        wsC.Activate
        Aviso "Selecione na aba Controle a linha do comprovante que deseja excluir.", vbExclamation
        Exit Sub
    End If

    linhaControle = ActiveCell.Row
    If NormalizarTexto(CStr(wsC.Cells(linhaControle, 1).Value)) = "total filtrado" Or Trim(CStr(wsC.Cells(linhaControle, 1).Value)) = "" Then
        Aviso "Selecione uma linha de registro do Controle, nao a linha de total.", vbExclamation
        Exit Sub
    End If

    numeroControle = CLng(Val(wsC.Cells(linhaControle, 1).Value))
    valor = NzD(wsC.Cells(linhaControle, 5).Value)
    mes = Trim(CStr(wsC.Cells(linhaControle, 4).Value))
    arquivo = Trim(CStr(wsC.Cells(linhaControle, 7).Value))
    hash = Trim(CStr(wsC.Cells(linhaControle, COL_HASH_CONTROLE).Value))
    If numeroControle <= 0 Then
        Aviso "Numero de controle invalido para exclusao.", vbExclamation
        Exit Sub
    End If

    If Application.Visible Then
        If MsgBox("Excluir definitivamente o Controle nº " & numeroControle & " e remover seu valor/link da aba Despesas?", vbQuestion + vbYesNo, "Excluir registro") <> vbYes Then Exit Sub
    End If

    On Error GoTo Falha
    DesprotegerPlanilhas
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    If wsC.AutoFilterMode Then wsC.AutoFilterMode = False
    LimparTotalControle wsC

    linhaDespesa = AcharLinhaDespesaPorControle(numeroControle)
    If linhaDespesa > 0 Then
        RemoverValorDespesa wsD, linhaDespesa, mes, valor
        RemoverControleDaObservacao wsD, linhaDespesa, numeroControle
    End If

    wsC.Rows(linhaControle).Delete Shift:=xlUp
    RemoverDaFilaPorHashOuArquivo hash, arquivo
    GarantirCabecalhoControle wsC
    AtualizarTotalControle wsC
    AtualizarFormulasDespesas wsD
    AtualizarResumoRateio
    AplicarFormatacaoPadrao

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    DesprotegerPlanilhas
    Aviso "Registro excluido. Valor e link removidos da aba Despesas.", vbInformation
    Exit Sub

Falha:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Aviso "Erro ao excluir registro: " & Err.Description, vbCritical
End Sub

Private Sub RemoverDaFilaPorHashOuArquivo(hash As String, arquivo As String)
    Dim ws As Worksheet, r As Long, ultima As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Fila")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    hash = Trim(CStr(hash))
    arquivo = Trim(CStr(arquivo))
    If hash = "" And arquivo = "" Then Exit Sub

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = ultima To 2 Step -1
        If (hash <> "" And Trim(CStr(ws.Cells(r, 12).Value)) = hash) Or _
           (arquivo <> "" And Trim(CStr(ws.Cells(r, 11).Value)) = arquivo) Then
            ws.Rows(r).Delete Shift:=xlUp
        End If
    Next r
End Sub
Private Function AcharLinhaDespesaPorControle(numeroControle As Long) As Long
    Dim wsD As Worksheet, r As Long, i As Long, numeros As Collection
    Set wsD = ThisWorkbook.Worksheets("Despesas")
    For r = LINHA_INICIAL_DADOS To LINHA_FINAL_DADOS
        Set numeros = ExtrairNumerosControle(CStr(wsD.Cells(r, COL_OBSERVACOES).Value))
        For i = 1 To numeros.Count
            If CLng(numeros(i)) = numeroControle Then
                AcharLinhaDespesaPorControle = r
                Exit Function
            End If
        Next i
    Next r
End Function

Private Sub RemoverValorDespesa(wsD As Worksheet, linhaDespesa As Long, mes As String, valor As Double)
    Dim colMes As Long, novoValor As Double
    If valor <= 0 Then Exit Sub
    If Trim(mes) <> "" Then
        colMes = AcharOuCriarColunaMes(mes)
    Else
        colMes = COL_EVENTUAL_ANUAL
    End If
    novoValor = NzD(wsD.Cells(linhaDespesa, colMes).Value) - valor
    If novoValor <= 0.004 Then
        wsD.Cells(linhaDespesa, colMes).ClearContents
    Else
        wsD.Cells(linhaDespesa, colMes).Value = Round(novoValor, 2)
    End If
End Sub

Private Sub RemoverControleDaObservacao(ws As Worksheet, linhaDespesa As Long, numeroControle As Long)
    Dim cel As Range, celStatus As Range, texto As String
    Set cel = ws.Cells(linhaDespesa, COL_OBSERVACOES)
    Set celStatus = ws.Cells(linhaDespesa, COL_COMPROVANTE)
    texto = ObservacaoSemNumero(cel.Value, numeroControle)
    cel.Hyperlinks.Delete
    celStatus.Hyperlinks.Delete
    If texto = "" Then
        cel.ClearContents
        If Trim(CStr(ws.Cells(linhaDespesa, 1).Value)) <> "" Or Trim(CStr(ws.Cells(linhaDespesa, 2).Value)) <> "" Then
            celStatus.Value = "Pendente"
        Else
            celStatus.ClearContents
        End If
    Else
        cel.Value = texto
        celStatus.Value = "Anexado"
        ws.Hyperlinks.Add Anchor:=cel, Address:="", SubAddress:="'Controle'!A3", TextToDisplay:=texto
        ws.Hyperlinks.Add Anchor:=celStatus, Address:="", SubAddress:="'Controle'!A3", TextToDisplay:=CStr(celStatus.Value)
    End If
End Sub

Private Function ObservacaoSemNumero(valorAtual As Variant, numeroRemover As Long) As String
    Dim numeros As Collection, i As Long, s As String
    Set numeros = ExtrairNumerosControle(CStr(valorAtual))
    For i = 1 To numeros.Count
        If CLng(numeros(i)) <> numeroRemover Then
            If s = "" Then
                s = CStr(numeros(i))
            Else
                s = s & ", " & CStr(numeros(i))
            End If
        End If
    Next i
    If s <> "" Then ObservacaoSemNumero = "Controle nº " & s
End Function

Public Sub LimparDados()
    NormalizarJanela
    Dim wsC As Worksheet, wsD As Worksheet
    Dim sobra As Long, r As Long, c As Long, ultimaControle As Long
    Set wsC = ThisWorkbook.Worksheets("Controle")
    Set wsD = ThisWorkbook.Worksheets("Despesas")

    DesprotegerPlanilhas
    DesprotegerAbaSegura wsC
    DesprotegerAbaSegura wsD
    If wsC.ProtectContents Or wsD.ProtectContents Then
        Aviso "Nao foi possivel destravar Controle/Despesas para limpar. Use Destravar edicao e tente novamente.", vbCritical
        Exit Sub
    End If

    wsC.Hyperlinks.Delete
    ultimaControle = UltimaLinhaControle(wsC)
    If ultimaControle < 2000 Then ultimaControle = 2000
    For r = LINHA_INICIAL_CONTROLE To 2000
        For c = 1 To COL_HASH_CONTROLE
            wsC.Cells(r, c).ClearContents
        Next c
    Next r
    GarantirCabecalhoControle wsC
    ResetarStatusFila
    wsD.Range(wsD.Cells(LINHA_INICIAL_DADOS, COL_MES_INICIAL), wsD.Cells(LINHA_FINAL_DADOS, COL_MES_FINAL)).ClearContents
    wsD.Range(wsD.Cells(LINHA_INICIAL_DADOS, COL_EVENTUAL_ANUAL), wsD.Cells(LINHA_FINAL_DADOS, COL_EVENTUAL_ANUAL)).ClearContents
    wsD.Range(wsD.Cells(LINHA_INICIAL_DADOS, COL_COMPROVANTE), wsD.Cells(LINHA_FINAL_DADOS, COL_OBSERVACOES)).ClearContents
    AtualizarFormulasDespesas wsD
    AtualizarResumoRateio
    AplicarFormatacaoPadrao

    For r = LINHA_INICIAL_CONTROLE To LINHA_INICIAL_CONTROLE + 300
        If Trim(CStr(wsC.Cells(r, 1).Value)) <> "" Or Trim(CStr(wsC.Cells(r, COL_HASH_CONTROLE).Value)) <> "" Then sobra = sobra + 1
    Next r
    ProtegerPlanilhas
    NormalizarTelaImport
    If sobra > 0 Then
        Aviso "Limpeza incompleta no Controle: ainda restaram " & sobra & " linha(s).", vbCritical
    Else
        Aviso "Dados limpos.", vbInformation
    End If
End Sub
Public Sub AbrirFormularioEdicao()
    NormalizarJanela
    Dim ws As Worksheet
    GarantirFormularioEdicao
    Set ws = ActiveSheet
    If ws.Name = "Fila" And ActiveCell.Row >= 2 Then
        ThisWorkbook.Worksheets(NOME_ABA_FORM).Range("B5").Value = "Fila"
        ThisWorkbook.Worksheets(NOME_ABA_FORM).Range("B6").Value = ActiveCell.Row
        CarregarRegistroEdicao
    ElseIf ws.Name = "Controle" And ActiveCell.Row >= 4 Then
        ThisWorkbook.Worksheets(NOME_ABA_FORM).Range("B5").Value = "Controle"
        ThisWorkbook.Worksheets(NOME_ABA_FORM).Range("B6").Value = ActiveCell.Row
        CarregarRegistroEdicao
    Else
        ThisWorkbook.Worksheets(NOME_ABA_FORM).Activate
    End If
End Sub

Public Sub CarregarRegistroEdicao()
    Dim wsO As Worksheet, wsE As Worksheet, origem As String, linha As Long
    Set wsE = ThisWorkbook.Worksheets(NOME_ABA_FORM)
    origem = Trim(CStr(wsE.Range("B5").Value))
    linha = CLng(Val(wsE.Range("B6").Value))
    If linha <= 0 Then
        Aviso "Informe uma linha valida para carregar.", vbExclamation
        Exit Sub
    End If
    DesprotegerPlanilhas
    LimparFormularioEdicao wsE
    If origem = "Fila" Then
        Set wsO = ThisWorkbook.Worksheets("Fila")
        wsE.Range("D8").Value = wsO.Cells(linha, 3).Value
        wsE.Range("D9").Value = wsO.Cells(linha, 4).Value
        wsE.Range("D10").Value = wsO.Cells(linha, 5).Value
        wsE.Range("D11").Value = wsO.Cells(linha, 6).Value
        wsE.Range("D12").Value = wsO.Cells(linha, 7).Value
        wsE.Range("D13").Value = wsO.Cells(linha, 8).Value
        wsE.Range("D14").Value = wsO.Cells(linha, 9).Value
        wsE.Range("D15").Value = wsO.Cells(linha, 10).Value
        wsE.Range("D16").Value = wsO.Cells(linha, 13).Value
        wsE.Range("D17").Value = wsO.Cells(linha, 11).Value
        If AcharLinhaDespesa(CStr(wsO.Cells(linha, 4).Value)) > 0 Then wsE.Range("D18").Value = ThisWorkbook.Worksheets("Despesas").Cells(AcharLinhaDespesa(CStr(wsO.Cells(linha, 4).Value)), COL_COTA_PARTE).Value
    ElseIf origem = "Controle" Then
        Set wsO = ThisWorkbook.Worksheets("Controle")
        wsE.Range("D8").Value = wsO.Cells(linha, 2).Value
        wsE.Range("D9").Value = wsO.Cells(linha, 3).Value
        wsE.Range("D11").Value = wsO.Cells(linha, 4).Value
        wsE.Range("D12").Value = wsO.Cells(linha, 5).Value
        wsE.Range("D15").Value = wsO.Cells(linha, 8).Value
        wsE.Range("D17").Value = wsO.Cells(linha, 7).Value
    Else
        Aviso "Origem deve ser Fila ou Controle.", vbExclamation
    End If
    AplicarFormatacaoPadrao
    ProtegerPlanilhas
    wsE.Activate
End Sub

Public Sub SalvarFormularioEdicao()
    Dim wsO As Worksheet, wsE As Worksheet, origem As String, linha As Long, valor As Double
    Set wsE = ThisWorkbook.Worksheets(NOME_ABA_FORM)
    origem = Trim(CStr(wsE.Range("B5").Value))
    linha = CLng(Val(wsE.Range("B6").Value))
    If linha <= 0 Then
        Aviso "Informe uma linha valida para salvar.", vbExclamation
        Exit Sub
    End If
    If Not TryParseDouble(CStr(wsE.Range("D12").Value), valor) Then
        Aviso "Valor invalido. Use numero em formato de moeda.", vbExclamation
        Exit Sub
    End If
    DesprotegerPlanilhas
    If origem = "Fila" Then
        Set wsO = ThisWorkbook.Worksheets("Fila")
        wsO.Cells(linha, 3).Value = wsE.Range("D8").Value
        wsO.Cells(linha, 4).Value = wsE.Range("D9").Value
        wsO.Cells(linha, 5).Value = wsE.Range("D10").Value
        wsO.Cells(linha, 6).Value = wsE.Range("D11").Value
        wsO.Cells(linha, 7).Value = valor
        wsO.Cells(linha, 8).Value = wsE.Range("D13").Value
        wsO.Cells(linha, 9).Value = wsE.Range("D14").Value
        wsO.Cells(linha, 10).Value = wsE.Range("D15").Value
        wsO.Cells(linha, 13).Value = wsE.Range("D16").Value
        wsO.Cells(linha, 11).Value = wsE.Range("D17").Value
        AtualizarCotaParteDespesa CStr(wsE.Range("D9").Value), CStr(wsE.Range("D18").Value)
    ElseIf origem = "Controle" Then
        Set wsO = ThisWorkbook.Worksheets("Controle")
        wsO.Cells(linha, 2).Value = wsE.Range("D8").Value
        wsO.Cells(linha, 3).Value = wsE.Range("D9").Value
        wsO.Cells(linha, 4).Value = wsE.Range("D11").Value
        wsO.Cells(linha, 5).Value = valor
        wsO.Cells(linha, 7).Value = wsE.Range("D17").Value
        wsO.Cells(linha, 8).Value = wsE.Range("D15").Value
    Else
        Aviso "Origem deve ser Fila ou Controle.", vbExclamation
    End If
    AplicarFormatacaoPadrao
    ProtegerPlanilhas
    Aviso "Registro atualizado pelo formulario.", vbInformation
End Sub

Private Sub AbrirFormularioControle(numeroControle As Long)
    If numeroControle <= 0 Then Exit Sub
    GarantirFormularioEdicao
    With ThisWorkbook.Worksheets(NOME_ABA_FORM)
        .Range("B5").Value = "Controle"
        .Range("B6").Value = numeroControle + 3
    End With
    CarregarRegistroEdicao
End Sub

Private Function ProximaLinhaControle(ws As Worksheet) As Long
    Dim r As Long
    LimparTotalControle ws
    For r = LINHA_INICIAL_CONTROLE To 2000
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" Then
            ProximaLinhaControle = r
            Exit Function
        End If
    Next r
    ProximaLinhaControle = UltimaLinhaControle(ws) + 1
    If ProximaLinhaControle < LINHA_INICIAL_CONTROLE Then ProximaLinhaControle = LINHA_INICIAL_CONTROLE
End Function
Private Function UltimoNumeroControle() As Long
    Dim ws As Worksheet, ultima As Long
    Set ws = ThisWorkbook.Worksheets("Controle")
    ultima = UltimaLinhaControle(ws)
    If ultima >= 4 Then UltimoNumeroControle = CLng(Val(ws.Cells(ultima, 1).Value))
End Function

Private Sub AtualizarCotaParteDespesa(despesa As String, marcador As String)
    Dim r As Long, ws As Worksheet, v As String
    r = AcharLinhaDespesa(despesa)
    If r = 0 Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Despesas")
    v = LCase$(Trim$(marcador))
    If v = "sim" Or v = "s" Or v = "x" Or v = "1" Or v = ChrW(9679) Then
        ws.Cells(r, COL_COTA_PARTE).Value = "Sim"
    Else
        ws.Cells(r, COL_COTA_PARTE).Value = ""
    End If
End Sub
Private Sub LimparFormularioEdicao(ws As Worksheet)
    ws.Range("D8:D18").ClearContents
End Sub

Public Sub AplicarFormatacaoPadrao()
    Dim ws As Worksheet, colTotal As Long, colMedia As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Despesas")
    colTotal = ColunaPorCabecalho(ws, "Total")
    colMedia = ColunaPorCabecalho(ws, "Média mensal")
    If colMedia > 0 Then ws.Range(ws.Cells(LINHA_INICIAL_DADOS, COL_MES_INICIAL), ws.Cells(LINHA_TOTAL, COL_VALOR_COTA_PARTE)).NumberFormat = "R$ #,##0.00" Else ws.Range("C5:N59").NumberFormat = "R$ #,##0.00"
    ws.Range(ws.Cells(LINHA_INICIAL_DADOS, COL_COTA_PARTE), ws.Cells(LINHA_FINAL_DADOS, COL_COTA_PARTE)).NumberFormat = "General"
    Set ws = ThisWorkbook.Worksheets("Controle")
    AtualizarTotalControle ws
    ws.Columns(1).NumberFormat = "0"
    ws.Columns(4).NumberFormat = "yyyy-mm"
    ws.Columns(5).NumberFormat = "R$ #,##0.00"
    Set ws = ThisWorkbook.Worksheets("Fila")
    ws.Columns(1).NumberFormat = "0"
    ws.Columns(5).NumberFormat = "dd/mm/yyyy"
    ws.Columns(6).NumberFormat = "yyyy-mm"
    ws.Columns(7).NumberFormat = "R$ #,##0.00"
    Set ws = ThisWorkbook.Worksheets(NOME_ABA_FORM)
    ws.Range("D10").NumberFormat = "dd/mm/yyyy"
    ws.Range("D11").NumberFormat = "yyyy-mm"
    ws.Range("D12").NumberFormat = "R$ #,##0.00"
    On Error GoTo 0
End Sub

Public Sub ProtegerPlanilhas()
    ' As planilhas ficam destravadas por padrao; use BloquearEdicao para travar manualmente.
    DesprotegerPlanilhas
End Sub

Private Sub AplicarProtecaoPlanilhas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:=SENHA_PROTECAO, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
        ws.EnableSelection = xlUnlockedCells
    Next ws
End Sub
Private Sub DesprotegerAbaSegura(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=SENHA_PROTECAO
    If ws.ProtectContents Then ws.Unprotect Password:=""
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub
Public Sub DesprotegerPlanilhas()
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=SENHA_PROTECAO
        If ws.ProtectContents Then ws.Unprotect Password:=""
        ws.EnableSelection = xlNoRestrictions
    Next ws
    On Error GoTo 0
End Sub

Public Sub DesbloquearEdicao()
    NormalizarJanela
    DesprotegerPlanilhas
    Aviso "Edicao liberada. Senha: daisy2026", vbInformation
End Sub

Public Sub BloquearEdicao()
    NormalizarJanela
    AplicarFormatacaoPadrao
    AplicarProtecaoPlanilhas
    Aviso "Planilhas bloqueadas novamente.", vbInformation
End Sub
Private Sub GarantirFormularioEdicao()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOME_ABA_FORM)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = NOME_ABA_FORM
    End If
End Sub
Private Sub NormalizarJanela()
    On Error Resume Next
    If ActiveSheet.Name = "Import" Then FixarLayoutImport
    If ActiveSheet.Name = "Despesas" Then
        CongelarReferenciasDespesas
        Exit Sub
    End If
    Dim p As Pane
    With ActiveWindow
        .FreezePanes = False
        .Split = False
        .SplitColumn = 0
        .SplitRow = 0
        For Each p In .Panes
            p.ScrollRow = 1
            p.ScrollColumn = 1
        Next p
        .ScrollRow = 1
        .ScrollColumn = 1
        .DisplayGridlines = False
        .Zoom = 100
    End With
    On Error GoTo 0
End Sub

Public Sub CongelarReferenciasDespesas()
    On Error Resume Next
    With ThisWorkbook.Worksheets("Despesas")
        .Activate
        .ScrollArea = ""
        .Columns("V").Hidden = False
        .Columns("W:XFD").Hidden = True
        .Range("C5").Select
    End With
    With ActiveWindow
        .FreezePanes = False
        .Split = False
        .SplitColumn = 2
        .SplitRow = 4
        .FreezePanes = True
        .ScrollColumn = 1
        .ScrollRow = 1
        .DisplayGridlines = False
        .Zoom = 100
    End With
    On Error GoTo 0
End Sub

Private Sub RestaurarLayoutImport()
    NormalizarTelaImport
End Sub
Private Sub FixarLayoutImport()
    On Error Resume Next
    Dim ws As Worksheet, sh As Shape, estavaProtegida As Boolean
    Set ws = ThisWorkbook.Worksheets("Import")
    estavaProtegida = ws.ProtectContents
    If estavaProtegida Then DesprotegerAbaSegura ws
    ws.Cells.EntireColumn.Hidden = False
    ws.Cells.EntireRow.Hidden = False
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:G").ColumnWidth = 16
    ws.Columns("H:J").ColumnWidth = 3
    ws.Rows("1:28").RowHeight = 21
    For Each sh In ws.Shapes
        sh.Placement = 3
    Next sh
    ws.ScrollArea = "A1:J28"
    If estavaProtegida Then ws.Protect Password:=SENHA_PROTECAO, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    On Error GoTo 0
End Sub
Private Sub NormalizarTelaImport()
    On Error Resume Next
    Dim p As Pane
    FixarLayoutImport
    With ThisWorkbook.Worksheets("Import")
        .Activate
        .Range("A1").Select
        .ScrollArea = "A1:J28"
    End With
    With ActiveWindow
        .FreezePanes = False
        .Split = False
        .SplitColumn = 0
        .SplitRow = 0
        For Each p In .Panes
            p.ScrollRow = 1
            p.ScrollColumn = 1
        Next p
        .ScrollRow = 1
        .ScrollColumn = 1
        .DisplayGridlines = False
        .Zoom = 100
    End With
    On Error GoTo 0
End Sub
Private Sub Aviso(msg As String, estilo As VbMsgBoxStyle)
    If Application.Visible Then MsgBox msg, estilo
End Sub

Private Function GarantirAbaFila() As Worksheet
    On Error Resume Next
    Set GarantirAbaFila = ThisWorkbook.Worksheets("Fila")
    On Error GoTo 0
    If GarantirAbaFila Is Nothing Then
        Set GarantirAbaFila = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GarantirAbaFila.Name = "Fila"
    End If
End Function

Private Sub ResetarStatusFila()
    Dim ws As Worksheet, ultima As Long, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Fila")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultima < 2 Then Exit Sub
    For r = 2 To ultima
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" Then
            ws.Cells(r, 2).Value = "pendente"
            ws.Cells(r, 14).ClearContents
        End If
    Next r
End Sub
Private Sub LimparFila(ws As Worksheet)
    ws.Cells.Clear
End Sub

Private Sub EscreverCabecalhoFila(ws As Worksheet)
    Dim headers As Variant, i As Long
    headers = Array("id", "status", "categoria", "despesa", "data", "mes", "valor", "pagador", "recebedor", "tipo", "arquivo", "hash", "periodicidade", "resultado")
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
        ws.Cells(1, i + 1).Font.Bold = True
    Next i
End Sub

Private Function LerArquivoTexto(caminho As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile caminho
    LerArquivoTexto = stm.ReadText
    stm.Close
End Function

Private Function ParseComprovantesJson(texto As String) As Collection
    Dim col As New Collection
    Dim posLista As Long, pos As Long, ini As Long, fim As Long
    Dim nivel As Long, ch As String, emString As Boolean, escape As Boolean
    Dim objTexto As String

    posLista = InStr(1, texto, Chr(34) & "comprovantes" & Chr(34), vbTextCompare)
    If posLista = 0 Then
        Set ParseComprovantesJson = col
        Exit Function
    End If

    pos = InStr(posLista, texto, "[")
    Do While pos > 0 And pos <= Len(texto)
        ch = Mid$(texto, pos, 1)
        If emString Then
            If escape Then
                escape = False
            ElseIf ch = "\" Then
                escape = True
            ElseIf ch = Chr(34) Then
                emString = False
            End If
        Else
            If ch = Chr(34) Then
                emString = True
            ElseIf ch = "{" Then
                If nivel = 0 Then ini = pos
                nivel = nivel + 1
            ElseIf ch = "}" Then
                nivel = nivel - 1
                If nivel = 0 And ini > 0 Then
                    fim = pos
                    objTexto = Mid$(texto, ini, fim - ini + 1)
                    col.Add ParseJsonObjetoSimples(objTexto)
                    ini = 0
                End If
            ElseIf ch = "]" And nivel = 0 Then
                Exit Do
            End If
        End If
        pos = pos + 1
    Loop

    Set ParseComprovantesJson = col
End Function

Private Function ParseJsonObjetoSimples(objTexto As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("id") = JsonValor(objTexto, "id")
    d("status") = JsonValor(objTexto, "status")
    d("categoria") = JsonValor(objTexto, "categoria")
    d("despesa") = JsonValor(objTexto, "despesa")
    d("data") = JsonValor(objTexto, "data")
    d("mes") = JsonValor(objTexto, "mes")
    d("valor") = JsonValor(objTexto, "valor")
    d("pagador") = JsonValor(objTexto, "pagador")
    d("recebedor") = JsonValor(objTexto, "recebedor")
    d("tipo") = JsonValor(objTexto, "tipo")
    d("arquivo") = JsonValor(objTexto, "arquivo")
    d("hash") = JsonValor(objTexto, "hash")
    d("periodicidade") = JsonValor(objTexto, "periodicidade")
    Set ParseJsonObjetoSimples = d
End Function

Private Function JsonValor(objTexto As String, chave As String) As String
    Dim re As Object, m As Object, padrao As String
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.MultiLine = True
    re.IgnoreCase = False
    padrao = Chr(34) & chave & Chr(34) & "\s*:\s*(" & Chr(34) & "(([^" & Chr(34) & "\\]|\\.)*)" & Chr(34) & "|null|-?\d+(\.\d+)?)"
    re.Pattern = padrao
    If re.Test(objTexto) Then
        Set m = re.Execute(objTexto)(0)
        If LCase$(m.SubMatches(0)) = "null" Then
            JsonValor = ""
        ElseIf Left$(m.SubMatches(0), 1) = Chr(34) Then
            JsonValor = JsonUnescape(m.SubMatches(1))
        Else
            JsonValor = m.SubMatches(0)
        End If
    Else
        JsonValor = ""
    End If
End Function

Private Function JsonUnescape(s As String) As String
    s = Replace(s, "\" & Chr(34), Chr(34))
    s = Replace(s, "\\", "\")
    s = Replace(s, "\/", "/")
    s = Replace(s, "\n", vbLf)
    s = Replace(s, "\r", vbCr)
    s = Replace(s, "\t", vbTab)
    JsonUnescape = s
End Function

Private Function ValorDict(d As Object, chave As String) As String
    If d.Exists(chave) Then ValorDict = CStr(d(chave)) Else ValorDict = ""
End Function

Private Function CaminhoBase() As String
    Dim p As String
    If EhWindows() Then
        ' No Windows, o caminho Linux salvo em Config!B7 nao e valido.
        p = ThisWorkbook.Path
    Else
        p = Trim(CStr(ThisWorkbook.Worksheets("Config").Range("B7").Value))
        If p = "" Then p = ThisWorkbook.Path
    End If
    CaminhoBase = p
End Function

Private Function CaminhoShell(caminho As String) As String
    Dim p As String
    p = Replace(caminho, "\", "/")
    If UCase$(Left$(p, 3)) = "Z:/" Then p = Mid$(p, 3)
    CaminhoShell = p
End Function

Private Function ResolverCaminhoArquivo(relOuAbs As String) As String
    Dim p As String
    p = Trim(CStr(relOuAbs))
    If p = "" Then
        ResolverCaminhoArquivo = ""
        Exit Function
    End If
    If EhWindows() Then p = Replace(p, "/", Application.PathSeparator)
    If p Like "[A-Za-z]:\*" Or Left$(p, 2) = "\\" Then
        ResolverCaminhoArquivo = p
    ElseIf Left$(p, 1) = "/" Then
        ResolverCaminhoArquivo = p
    Else
        ResolverCaminhoArquivo = CaminhoBase() & Application.PathSeparator & p
    End If
End Function

Private Sub LimparTotalControle(ws As Worksheet)
    Dim r As Long
    For r = LINHA_INICIAL_CONTROLE To 2005
        If NormalizarTexto(CStr(ws.Cells(r, 1).Value)) = "total filtrado" Then
            ws.Rows(r).Hidden = False
            ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_HASH_CONTROLE)).ClearContents
            ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_HASH_CONTROLE)).Interior.ColorIndex = xlNone
            ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_HASH_CONTROLE)).Font.Bold = False
        End If
    Next r
End Sub

Private Sub AtualizarTotalControle(ws As Worksheet)
    Dim ultima As Long, linhaTotal As Long
    LimparTotalControle ws
    ultima = UltimaLinhaControle(ws)
    If ultima < LINHA_INICIAL_CONTROLE Then Exit Sub
    linhaTotal = ultima + 1
    ws.Cells(linhaTotal, 1).Value = "Total filtrado"
    ws.Cells(linhaTotal, 4).Value = "Qtd."
    ws.Cells(linhaTotal, 5).Formula = "=SUBTOTAL(109,E" & LINHA_INICIAL_CONTROLE & ":E" & ultima & ")"
    ws.Cells(linhaTotal, 6).Formula = "=SUBTOTAL(103,A" & LINHA_INICIAL_CONTROLE & ":A" & ultima & ")"
    ws.Range(ws.Cells(linhaTotal, 1), ws.Cells(linhaTotal, COL_HASH_CONTROLE)).Interior.Color = RGB(219, 234, 254)
    ws.Range(ws.Cells(linhaTotal, 1), ws.Cells(linhaTotal, COL_HASH_CONTROLE)).Font.Bold = True
    ws.Cells(linhaTotal, 5).NumberFormat = "R$ #,##0.00"
End Sub
Private Sub GarantirCabecalhoControle(ws As Worksheet)
    ws.Cells(3, COL_HASH_CONTROLE).Value = "Hash"
    ws.Columns(COL_HASH_CONTROLE).Hidden = True
End Sub

Private Function UltimaLinhaControle(ws As Worksheet) As Long
    Dim r As Long, lastA As Long, lastHash As Long, limite As Long
    lastA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastHash = ws.Cells(ws.Rows.Count, COL_HASH_CONTROLE).End(xlUp).Row
    limite = lastA
    If lastHash > limite Then limite = lastHash
    If limite < LINHA_INICIAL_CONTROLE Then
        UltimaLinhaControle = LINHA_INICIAL_CONTROLE - 1
        Exit Function
    End If
    For r = limite To LINHA_INICIAL_CONTROLE Step -1
        If Trim(CStr(ws.Cells(r, COL_HASH_CONTROLE).Value)) <> "" Then
            UltimaLinhaControle = r
            Exit Function
        End If
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" And NormalizarTexto(CStr(ws.Cells(r, 1).Value)) <> "total filtrado" Then
            UltimaLinhaControle = r
            Exit Function
        End If
    Next r
    UltimaLinhaControle = LINHA_INICIAL_CONTROLE - 1
End Function
Private Function HashJaImportado(ws As Worksheet, hash As String) As Boolean
    Dim r As Long
    For r = LINHA_INICIAL_CONTROLE To 2000
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" And Trim(CStr(ws.Cells(r, COL_HASH_CONTROLE).Value)) = hash Then
            HashJaImportado = True
            Exit Function
        End If
    Next r
End Function

Private Function AcharLinhaDespesa(desp As String, Optional categoria As String = "") As Long
    Dim r As Long, mesmaDespesa As Boolean, mesmaCategoria As Boolean
    With ThisWorkbook.Worksheets("Despesas")
        For r = LINHA_INICIAL_DADOS To LINHA_FINAL_DADOS
            mesmaDespesa = (NormalizarTexto(CStr(.Cells(r, 2).Value)) = NormalizarTexto(desp))
            mesmaCategoria = (Trim(categoria) = "" Or NormalizarTexto(CStr(.Cells(r, 1).Value)) = NormalizarTexto(categoria))
            If mesmaDespesa And mesmaCategoria Then
                AcharLinhaDespesa = r
                Exit Function
            End If
        Next r
    End With
End Function

Private Function AcharOuCriarColunaMes(mes As String) As Long
    Dim nMes As Long
    nMes = CLng(Val(Right$(Trim$(mes), 2)))
    If nMes < 1 Or nMes > 12 Then Err.Raise vbObjectError + 1001, , "Mes invalido: " & mes
    AcharOuCriarColunaMes = COL_MES_INICIAL + nMes - 1
End Function

Private Sub AtualizarFormulasDespesas(ws As Worksheet)
    Dim r As Long
    For r = LINHA_INICIAL_DADOS To LINHA_FINAL_DADOS
        ws.Cells(r, COL_VALOR_MES).Formula = "=IF(OR(A" & r & "<>" & Chr(34) & Chr(34) & ",B" & r & "<>" & Chr(34) & Chr(34) & "),(SUM(" & ws.Range(ws.Cells(r, COL_MES_INICIAL), ws.Cells(r, COL_MES_FINAL)).Address(False, False) & ")+" & ws.Cells(r, COL_EVENTUAL_ANUAL).Address(False, False) & ")/12," & Chr(34) & Chr(34) & ")"
        ws.Cells(r, COL_VALOR_ANUAL).Formula = "=IF(OR(A" & r & "<>" & Chr(34) & Chr(34) & ",B" & r & "<>" & Chr(34) & Chr(34) & "),SUM(" & ws.Range(ws.Cells(r, COL_MES_INICIAL), ws.Cells(r, COL_MES_FINAL)).Address(False, False) & ")+" & ws.Cells(r, COL_EVENTUAL_ANUAL).Address(False, False) & "," & Chr(34) & Chr(34) & ")"
        ws.Cells(r, COL_VALOR_COTA_PARTE_ANUAL).Formula = "=IF(OR(A" & r & "<>" & Chr(34) & Chr(34) & ",B" & r & "<>" & Chr(34) & Chr(34) & "),IF(" & ws.Cells(r, COL_COTA_PARTE).Address(False, False) & "=" & Chr(34) & "Sim" & Chr(34) & "," & ws.Cells(r, COL_VALOR_ANUAL).Address(False, False) & "/2," & ws.Cells(r, COL_VALOR_ANUAL).Address(False, False) & ")," & Chr(34) & Chr(34) & ")"
        ws.Cells(r, COL_VALOR_COTA_PARTE).Formula = "=IF(OR(A" & r & "<>" & Chr(34) & Chr(34) & ",B" & r & "<>" & Chr(34) & Chr(34) & "),IF(" & ws.Cells(r, COL_COTA_PARTE).Address(False, False) & "=" & Chr(34) & "Sim" & Chr(34) & "," & ws.Cells(r, COL_VALOR_MES).Address(False, False) & "/2," & ws.Cells(r, COL_VALOR_MES).Address(False, False) & ")," & Chr(34) & Chr(34) & ")"
        If Trim(CStr(ws.Cells(r, 1).Value)) <> "" Or Trim(CStr(ws.Cells(r, 2).Value)) <> "" Then
            If Trim(CStr(ws.Cells(r, COL_COMPROVANTE).Value)) = "" Then ws.Cells(r, COL_COMPROVANTE).Value = "Pendente"
        Else
            ws.Cells(r, COL_COMPROVANTE).ClearContents
            ws.Cells(r, COL_OBSERVACOES).ClearContents
        End If
    Next r
    ws.Cells(LINHA_TOTAL, COL_VALOR_MES).Formula = "=SUM(" & ws.Cells(LINHA_INICIAL_DADOS, COL_VALOR_MES).Address(False, False) & ":" & ws.Cells(LINHA_FINAL_DADOS, COL_VALOR_MES).Address(False, False) & ")"
    ws.Cells(LINHA_TOTAL, COL_VALOR_ANUAL).Formula = "=SUM(" & ws.Cells(LINHA_INICIAL_DADOS, COL_VALOR_ANUAL).Address(False, False) & ":" & ws.Cells(LINHA_FINAL_DADOS, COL_VALOR_ANUAL).Address(False, False) & ")"
    ws.Cells(LINHA_TOTAL, COL_VALOR_COTA_PARTE_ANUAL).Formula = "=SUM(" & ws.Cells(LINHA_INICIAL_DADOS, COL_VALOR_COTA_PARTE_ANUAL).Address(False, False) & ":" & ws.Cells(LINHA_FINAL_DADOS, COL_VALOR_COTA_PARTE_ANUAL).Address(False, False) & ")"
    ws.Cells(LINHA_TOTAL, COL_VALOR_COTA_PARTE).Formula = "=SUM(" & ws.Cells(LINHA_INICIAL_DADOS, COL_VALOR_COTA_PARTE).Address(False, False) & ":" & ws.Cells(LINHA_FINAL_DADOS, COL_VALOR_COTA_PARTE).Address(False, False) & ")"
End Sub

Private Sub AtualizarResumoRateio()
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Resumo e Rateio")
    If ws Is Nothing Then Exit Sub
    ws.Cells(3, 1).Value = "Categoria"
    ws.Cells(3, 2).Value = "Cota parte mes (R$)"
    ws.Cells(3, 3).Value = "Cota parte anual (R$)"
    For r = 4 To 13
        ws.Cells(r, 2).Formula = "=SUMIFS(Despesas!$T:$T,Despesas!$A:$A,A" & r & ")"
        ws.Cells(r, 3).Formula = "=SUMIFS(Despesas!$S:$S,Despesas!$A:$A,A" & r & ")"
    Next r
    ws.Cells(14, 2).Formula = "=SUM(B4:B13)"
    ws.Cells(14, 3).Formula = "=SUM(C4:C13)"
    ws.Range("B4:C14").NumberFormat = "R$ #,##0.00"
    On Error GoTo 0
End Sub
Private Function ColunaPorCabecalho(ws As Worksheet, prefixo As String) As Long
    Dim c As Long, alvo As String, atual As String
    alvo = NormalizarTexto(prefixo)
    For c = 1 To ws.Cells(LINHA_CABECALHO, ws.Columns.Count).End(xlToLeft).Column
        atual = NormalizarTexto(CStr(ws.Cells(LINHA_CABECALHO, c).Value))
        If Left$(atual, Len(alvo)) = alvo Then
            ColunaPorCabecalho = c
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 1000, , "Cabecalho nao encontrado: " & prefixo
End Function

Private Sub AtualizarObservacaoComLink(ws As Worksheet, linhaDespesa As Long, numeroControle As Long, arquivo As String)
    Dim cel As Range, celStatus As Range, texto As String
    Set cel = ws.Cells(linhaDespesa, COL_OBSERVACOES)
    Set celStatus = ws.Cells(linhaDespesa, COL_COMPROVANTE)
    texto = AtualizarObservacao(cel.Value, numeroControle)
    cel.Hyperlinks.Delete
    cel.Value = texto
    celStatus.Hyperlinks.Delete
    If Trim(CStr(celStatus.Value)) = "" Then celStatus.Value = "Anexado"
    ws.Hyperlinks.Add Anchor:=cel, Address:="", SubAddress:="'Controle'!A3", TextToDisplay:=texto
    ws.Hyperlinks.Add Anchor:=celStatus, Address:="", SubAddress:="'Controle'!A3", TextToDisplay:=CStr(celStatus.Value)
End Sub
Private Function AtualizarObservacao(valorAtual As Variant, numeroControle As Long) As String
    Dim s As String
    s = Trim(CStr(valorAtual))
    If s = "" Then
        AtualizarObservacao = "Controle nº " & numeroControle
    Else
        AtualizarObservacao = s & ", " & numeroControle
    End If
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

Public Sub AbrirCadastros()
    NormalizarJanela
    GarantirAbaCadastros
    ThisWorkbook.Worksheets(NOME_ABA_CADASTROS).Activate
End Sub

Private Sub GarantirAbaCadastros()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOME_ABA_CADASTROS)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = NOME_ABA_CADASTROS
    End If
End Sub

Public Sub CarregarCadastroSelecionado()
    Dim ws As Worksheet, r As Long
    GarantirAbaCadastros
    Set ws = ThisWorkbook.Worksheets(NOME_ABA_CADASTROS)
    If ActiveSheet.Name <> ws.Name Or ActiveCell.Row < 13 Then
        Aviso "Selecione uma linha da lista de cadastros.", vbExclamation
        ws.Activate
        Exit Sub
    End If
    r = ActiveCell.Row
    ws.Range("C4").Value = ws.Cells(r, 1).Value
    ws.Range("C6").Value = ws.Cells(r, 2).Value
    ws.Range("C8").Value = ws.Cells(r, 3).Value
    ws.Range("E8").Value = ws.Cells(r, 4).Value
    ws.Range("G8").Value = ws.Cells(r, 5).Value
End Sub

Public Sub SalvarCadastroDespesa()
    Dim ws As Worksheet, cat As String, desp As String, cota As String, periodicidade As String, ativo As String, r As Long
    GarantirAbaCadastros
    Set ws = ThisWorkbook.Worksheets(NOME_ABA_CADASTROS)
    cat = Trim(CStr(ws.Range("C4").Value))
    desp = Trim(CStr(ws.Range("C6").Value))
    cota = NormalizarSimNao(CStr(ws.Range("C8").Value))
    periodicidade = Trim(CStr(ws.Range("E8").Value))
    ativo = NormalizarSimNao(CStr(ws.Range("G8").Value))
    If ativo = "" Then ativo = "Sim"
    If periodicidade = "" Then periodicidade = "Mensal"
    If cat = "" Or desp = "" Then
        Aviso "Informe categoria e despesa antes de salvar.", vbExclamation
        Exit Sub
    End If
    DesprotegerPlanilhas
    r = LinhaCadastro(ws, cat, desp)
    If r = 0 Then r = ProximaLinhaCadastro(ws)
    ws.Cells(r, 1).Value = cat
    ws.Cells(r, 2).Value = desp
    ws.Cells(r, 3).Value = cota
    ws.Cells(r, 4).Value = periodicidade
    ws.Cells(r, 5).Value = ativo
    GarantirDespesaCadastro cat, desp, cota
    AplicarFormatacaoPadrao
    DesprotegerPlanilhas
    Aviso "Cadastro salvo e sincronizado com Despesas.", vbInformation
End Sub

Public Sub SincronizarCadastrosDespesas()
    Dim ws As Worksheet, r As Long, ultima As Long, cat As String, desp As String, cota As String, ativo As String, n As Long
    GarantirAbaCadastros
    Set ws = ThisWorkbook.Worksheets(NOME_ABA_CADASTROS)
    DesprotegerPlanilhas
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultima < 13 Then ultima = 13
    For r = 13 To ultima
        cat = Trim(CStr(ws.Cells(r, 1).Value))
        desp = Trim(CStr(ws.Cells(r, 2).Value))
        cota = NormalizarSimNao(CStr(ws.Cells(r, 3).Value))
        ativo = NormalizarSimNao(CStr(ws.Cells(r, 5).Value))
        If cat <> "" And desp <> "" And ativo <> "Nao" Then
            GarantirDespesaCadastro cat, desp, cota
            n = n + 1
        End If
    Next r
    AtualizarCadastroAPartirDespesas ws
    AtualizarFormulasDespesas ThisWorkbook.Worksheets("Despesas")
    AtualizarResumoRateio
    AplicarFormatacaoPadrao
    DesprotegerPlanilhas
    Aviso "Cadastros sincronizados: " & n & " despesa(s) ativa(s).", vbInformation
End Sub

Public Sub AtualizarCadastroAPartirDespesas(Optional wsC As Worksheet)
    Dim wsD As Worksheet, r As Long, out As Long, cat As String, desp As String
    If wsC Is Nothing Then
        GarantirAbaCadastros
        Set wsC = ThisWorkbook.Worksheets(NOME_ABA_CADASTROS)
    End If
    Set wsD = ThisWorkbook.Worksheets("Despesas")
    For r = 13 To 400
        wsC.Range(wsC.Cells(r, 1), wsC.Cells(r, 5)).ClearContents
    Next r
    out = 13
    For r = LINHA_INICIAL_DADOS To LINHA_FINAL_DADOS
        cat = Trim(CStr(wsD.Cells(r, 1).Value))
        desp = Trim(CStr(wsD.Cells(r, 2).Value))
        If cat <> "" And desp <> "" Then
            wsC.Cells(out, 1).Value = cat
            wsC.Cells(out, 2).Value = desp
            wsC.Cells(out, 3).Value = NormalizarSimNao(CStr(wsD.Cells(r, COL_COTA_PARTE).Value))
            wsC.Cells(out, 4).Value = "Mensal"
            wsC.Cells(out, 5).Value = "Sim"
            out = out + 1
        End If
    Next r
End Sub

Private Sub GarantirDespesaCadastro(categoria As String, despesa As String, cota As String)
    Dim wsD As Worksheet, r As Long
    Set wsD = ThisWorkbook.Worksheets("Despesas")
    r = AcharLinhaDespesa(despesa, categoria)
    If r = 0 Then r = ProximaLinhaDespesaCadastro(wsD)
    wsD.Cells(r, 1).Value = categoria
    wsD.Cells(r, 2).Value = despesa
    wsD.Cells(r, COL_COTA_PARTE).Value = cota
    If Trim(CStr(wsD.Cells(r, COL_COMPROVANTE).Value)) = "" Then wsD.Cells(r, COL_COMPROVANTE).Value = "Pendente"
    AtualizarFormulasDespesas wsD
    GarantirCategoriaResumo categoria
End Sub

Private Function ProximaLinhaDespesaCadastro(ws As Worksheet) As Long
    Dim r As Long
    For r = LINHA_INICIAL_DADOS To LINHA_FINAL_DADOS
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" And Trim(CStr(ws.Cells(r, 2).Value)) = "" Then
            ProximaLinhaDespesaCadastro = r
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 1100, , "Nao ha linhas livres em Despesas. Aumente a area de cadastro."
End Function

Private Sub GarantirCategoriaResumo(categoria As String)
    Dim ws As Worksheet, r As Long, livre As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Resumo e Rateio")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    For r = 4 To 13
        If NormalizarTexto(CStr(ws.Cells(r, 1).Value)) = NormalizarTexto(categoria) Then Exit Sub
        If livre = 0 And Trim(CStr(ws.Cells(r, 1).Value)) = "" Then livre = r
    Next r
    If livre > 0 Then ws.Cells(livre, 1).Value = categoria
End Sub

Private Function LinhaCadastro(ws As Worksheet, categoria As String, despesa As String) As Long
    Dim r As Long, ultima As Long
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 13 To ultima
        If NormalizarTexto(CStr(ws.Cells(r, 1).Value)) = NormalizarTexto(categoria) _
           And NormalizarTexto(CStr(ws.Cells(r, 2).Value)) = NormalizarTexto(despesa) Then
            LinhaCadastro = r
            Exit Function
        End If
    Next r
End Function

Private Function ProximaLinhaCadastro(ws As Worksheet) As Long
    Dim r As Long
    For r = 13 To 400
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" And Trim(CStr(ws.Cells(r, 2).Value)) = "" Then
            ProximaLinhaCadastro = r
            Exit Function
        End If
    Next r
    ProximaLinhaCadastro = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
End Function

Private Function NormalizarSimNao(valor As String) As String
    Dim v As String
    v = NormalizarTexto(valor)
    If v = "sim" Or v = "s" Or v = "x" Or v = "1" Or v = "true" Then
        NormalizarSimNao = "Sim"
    ElseIf v = "nao" Or v = "n" Or v = "0" Or v = "false" Then
        NormalizarSimNao = "Nao"
    Else
        NormalizarSimNao = Trim(valor)
    End If
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


