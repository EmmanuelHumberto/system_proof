Option Explicit

Const xlToLeft = -4159
Const xlCenter = -4108
Const xlUnlockedCells = 1

Dim fso, xl, wb, wsD, wsR, ws, r
Set fso = CreateObject("Scripting.FileSystemObject")
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False
Set wb = xl.Workbooks.Open(fso.GetAbsolutePathName("Planilha_Comprovantes_VBA.xlsm"))

For Each ws In wb.Worksheets
    On Error Resume Next
    ws.Unprotect "daisy2026"
    ws.Unprotect ""
    On Error GoTo 0
Next

Set wsD = wb.Worksheets("Despesas")

' Estrutura final:
' R = Cota parte?
' S = Valor cota parte anual
' T = Valor cota parte mes
' U = Comprovante
' V = Observacoes
If LCase(Trim(CStr(wsD.Cells(4, 19).Value))) <> "valor cota parte anual" Then
    wsD.Columns(19).Insert
End If

wsD.Cells(4, 18).Value = "Cota parte?"
wsD.Cells(4, 19).Value = "Valor cota parte anual"
wsD.Cells(4, 20).Value = "Valor cota parte mes"
wsD.Cells(4, 21).Value = "Comprovante"
wsD.Cells(4, 22).Value = "Observacoes"

For r = 5 To 57
    wsD.Cells(r, 16).Formula = "=(SUM(C" & r & ":N" & r & ")+O" & r & ")/12"
    wsD.Cells(r, 17).Formula = "=SUM(C" & r & ":N" & r & ")+O" & r
    wsD.Cells(r, 19).Formula = "=IF(R" & r & "=""Sim"",Q" & r & "/2,Q" & r & ")"
    wsD.Cells(r, 20).Formula = "=IF(R" & r & "=""Sim"",P" & r & "/2,P" & r & ")"
    If Trim(CStr(wsD.Cells(r, 21).Value)) = "" Then wsD.Cells(r, 21).Value = "Pendente"
Next
wsD.Cells(59, 16).Formula = "=SUM(P5:P57)"
wsD.Cells(59, 17).Formula = "=SUM(Q5:Q57)"
wsD.Cells(59, 19).Formula = "=SUM(S5:S57)"
wsD.Cells(59, 20).Formula = "=SUM(T5:T57)"

wsD.Range("A4:V4").Font.Bold = True
wsD.Range("A4:V4").HorizontalAlignment = xlCenter
wsD.Range("C5:Q59").NumberFormat = """R$"" #,##0.00"
wsD.Range("S5:T59").NumberFormat = """R$"" #,##0.00"
wsD.Range("R5:R57").NumberFormat = "General"
wsD.Columns("A:V").AutoFit

Set wsR = wb.Worksheets("Resumo e Rateio")
For r = 4 To 13
    wsR.Cells(r, 2).Formula = "=SUMIFS(Despesas!$T:$T,Despesas!$A:$A,A" & r & ")"
Next
wsR.Cells(14, 2).Formula = "=SUM(B4:B13)"
wsR.Range("B4:B14").NumberFormat = """R$"" #,##0.00"

For Each ws In wb.Worksheets
    On Error Resume Next
    ws.Cells.Locked = True
    If ws.Name = "Editar" Then
        ws.Range("B5").Locked = False
        ws.Range("D5").Locked = False
        ws.Range("D8:D18").Locked = False
    End If
    If ws.Name = "Despesas" Then ws.Range("R5:R57").Locked = False
    ws.Protect "daisy2026", True, True, True, True
    ws.EnableSelection = xlUnlockedCells
    On Error GoTo 0
Next

wb.Worksheets("Import").Activate
wb.Worksheets("Import").Range("A1").Select
With xl.ActiveWindow
    .FreezePanes = False
    .Split = False
    .SplitColumn = 0
    .SplitRow = 0
    .ScrollColumn = 1
    .ScrollRow = 1
    .DisplayGridlines = False
    .Zoom = 100
End With

wb.Save
wb.Close False
xl.Quit
WScript.Echo "OK: cota anual e resumo aplicados."
