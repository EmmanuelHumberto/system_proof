Option Explicit
Const msoAutomationSecurityLow = 1
Const xlCenter = -4108

Dim xl, wb, ws, fso, base, arquivoXlsm
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
Set ws = wb.Worksheets("Despesas")
Check "Abrir Despesas"

BackupDespesas wb, ws
ReorganizarDespesas ws
wb.Save
Check "Salvar"
FecharExcel
WScript.Echo "OK: aba Despesas reorganizada em 12 meses + cota parte."
WScript.Quit 0

Sub BackupDespesas(wb, ws)
    Dim nome
    nome = "Backup_Despesas_" & Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2)
    ws.Copy , wb.Worksheets(wb.Worksheets.Count)
    Check "Criar backup Despesas"
    wb.Worksheets(wb.Worksheets.Count).Name = nome
    Check "Nomear backup Despesas"
End Sub

Sub ReorganizarDespesas(ws)
    Dim dados, linhas, r, c, mesIdx, oldCol, headers, oldMap
    Set oldMap = CreateObject("Scripting.Dictionary")
    ws.Unprotect "daisy2026"
    If Err.Number <> 0 Then Err.Clear: ws.Unprotect ""
    Check "Desproteger"

    For c = 3 To ws.Cells(4, ws.Columns.Count).End(-4159).Column
        mesIdx = MesDoCabecalho(CStr(ws.Cells(4, c).Value))
        If mesIdx > 0 Then oldMap(CStr(mesIdx)) = c
    Next

    ReDim dados(57, 12)
    ReDim linhas(57, 5)
    For r = 5 To 57
        linhas(r, 1) = ws.Cells(r, 1).Value
        linhas(r, 2) = ws.Cells(r, 2).Value
        linhas(r, 3) = ws.Cells(r, 15).Value
        linhas(r, 4) = ws.Cells(r, 20).Value
        linhas(r, 5) = ws.Cells(r, 21).Value
        For mesIdx = 1 To 12
            If oldMap.Exists(CStr(mesIdx)) Then
                oldCol = CLng(oldMap(CStr(mesIdx)))
                dados(r, mesIdx) = ws.Cells(r, oldCol).Value
            Else
                dados(r, mesIdx) = ""
            End If
        Next
    Next

    ws.Range("A4:V59").ClearContents
    headers = Array("Categoria", "Despesa", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez", "Despesa eventual/anual", "Valor mes", "Valor anual", "Cota parte?", "Valor cota parte anual", "Valor cota parte mes", "Comprovante", "Observacoes")
    For c = 0 To UBound(headers)
        ws.Cells(4, c + 1).Value = headers(c)
    Next

    For r = 5 To 57
        ws.Cells(r, 1).Value = linhas(r, 1)
        ws.Cells(r, 2).Value = linhas(r, 2)
        For mesIdx = 1 To 12
            ws.Cells(r, 2 + mesIdx).Value = dados(r, mesIdx)
        Next
        ws.Cells(r, 15).Value = linhas(r, 3)
        If InStr(1, LCase(RemoveAcento(CStr(ws.Cells(r, 2).Value))), "quota-parte", vbTextCompare) > 0 Then ws.Cells(r, 18).Value = "Sim"
        ws.Cells(r, 16).Formula = "=(SUM(C" & r & ":N" & r & ")+O" & r & ")/12"
        ws.Cells(r, 17).Formula = "=SUM(C" & r & ":N" & r & ")+O" & r
        ws.Cells(r, 19).Formula = "=IF(R" & r & "=""Sim"",Q" & r & "/2,Q" & r & ")"
        ws.Cells(r, 20).Formula = "=IF(R" & r & "=""Sim"",P" & r & "/2,P" & r & ")"
        ws.Cells(r, 21).Value = linhas(r, 4)
        ws.Cells(r, 22).Value = linhas(r, 5)
        If Trim(CStr(ws.Cells(r, 21).Value)) = "" Then ws.Cells(r, 21).Value = "Pendente"
    Next
    ws.Cells(59, 16).Formula = "=SUM(P5:P57)"
    ws.Cells(59, 17).Formula = "=SUM(Q5:Q57)"
    ws.Cells(59, 19).Formula = "=SUM(S5:S57)"
    ws.Cells(59, 20).Formula = "=SUM(T5:T57)"

    ws.Range("A4:V4").Font.Bold = True
    ws.Range("A4:V4").HorizontalAlignment = xlCenter
    ws.Range("C5:T59").NumberFormatLocal = "R$ #.##0,00"
    ws.Range("S5:T59").NumberFormatLocal = "R$ #.##0,00"
    ws.Range("R5:R57").HorizontalAlignment = xlCenter
    ws.Range("R5:R57").Font.Size = 12
    ws.Columns("A:U").AutoFit
    ws.Columns("T:U").ColumnWidth = 22
    ws.Cells.Locked = True
    ws.Range("R5:R57").Locked = False
    ws.Protect "daisy2026", True, True, True, True
    ws.EnableSelection = 1
End Sub

Function MesDoCabecalho(s)
    s = LCase(RemoveAcento(s))
    If InStr(s, "janeiro") > 0 Or s = "jan" Then MesDoCabecalho = 1: Exit Function
    If InStr(s, "fevereiro") > 0 Or s = "fev" Then MesDoCabecalho = 2: Exit Function
    If InStr(s, "marco") > 0 Or s = "mar" Then MesDoCabecalho = 3: Exit Function
    If InStr(s, "abril") > 0 Or s = "abr" Then MesDoCabecalho = 4: Exit Function
    If InStr(s, "maio") > 0 Or s = "mai" Then MesDoCabecalho = 5: Exit Function
    If InStr(s, "junho") > 0 Or s = "jun" Then MesDoCabecalho = 6: Exit Function
    If InStr(s, "julho") > 0 Or s = "jul" Then MesDoCabecalho = 7: Exit Function
    If InStr(s, "agosto") > 0 Or s = "ago" Then MesDoCabecalho = 8: Exit Function
    If InStr(s, "setembro") > 0 Or s = "set" Then MesDoCabecalho = 9: Exit Function
    If InStr(s, "outubro") > 0 Or s = "out" Then MesDoCabecalho = 10: Exit Function
    If InStr(s, "novembro") > 0 Or s = "nov" Then MesDoCabecalho = 11: Exit Function
    If InStr(s, "dezembro") > 0 Or s = "dez" Then MesDoCabecalho = 12: Exit Function
    MesDoCabecalho = 0
End Function

Function RemoveAcento(s)
    s = Replace(s, "á", "a"): s = Replace(s, "à", "a"): s = Replace(s, "ã", "a"): s = Replace(s, "â", "a")
    s = Replace(s, "é", "e"): s = Replace(s, "ê", "e")
    s = Replace(s, "í", "i")
    s = Replace(s, "ó", "o"): s = Replace(s, "õ", "o"): s = Replace(s, "ô", "o")
    s = Replace(s, "ú", "u"): s = Replace(s, "ç", "c")
    RemoveAcento = s
End Function

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
