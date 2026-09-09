Option Explicit

Const xlUnlockedCells = 1
Const xlNoRestrictions = 0

Dim fso, xl, wb, ws, i
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
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
Next

' Remove abas temporarias/de backup criadas durante manutencoes anteriores.
For i = wb.Worksheets.Count To 1 Step -1
    Set ws = wb.Worksheets(i)
    If LCase(Left(ws.Name, 7)) = "backup_" Then
        ws.Delete
    End If
Next

' Limpa dados importados, mantendo cabecalhos e estrutura.
Set ws = wb.Worksheets("Controle")
ws.Hyperlinks.Delete
For r = 4 To 2000
    For i = 1 To 9
        ws.Cells(r, i).ClearContents
    Next
Next
ws.Range("I:I").ClearContents
ws.Range("I3").Value = "Hash"
ws.Columns("I").Hidden = True

Set ws = wb.Worksheets("Fila")
ws.Hyperlinks.Delete
ws.Cells.ClearContents

Set ws = wb.Worksheets("Despesas")
ws.Range("C5:N57").ClearContents
ws.Range("O5:O57").ClearContents
ws.Range("U5:V57").ClearContents

' Restaura formulas das colunas calculadas de Despesas.
Dim r
For r = 5 To 57
    ws.Cells(r, 16).Formula = "=(SUM(C" & r & ":N" & r & ")+O" & r & ")/12"
    ws.Cells(r, 17).Formula = "=SUM(C" & r & ":N" & r & ")+O" & r
    ws.Cells(r, 19).Formula = "=IF(R" & r & "=""Sim"",Q" & r & "/2,Q" & r & ")"
    ws.Cells(r, 20).Formula = "=IF(R" & r & "=""Sim"",P" & r & "/2,P" & r & ")"
Next
ws.Cells(59, 16).Formula = "=SUM(P5:P57)"
ws.Cells(59, 17).Formula = "=SUM(Q5:Q57)"
ws.Cells(59, 19).Formula = "=SUM(S5:S57)"
ws.Cells(59, 20).Formula = "=SUM(T5:T57)"

' Normaliza visual das abas principais.
For Each ws In wb.Worksheets
    ws.Activate
    xl.ActiveWindow.FreezePanes = False
    xl.ActiveWindow.Split = False
    xl.ActiveWindow.SplitColumn = 0
    xl.ActiveWindow.SplitRow = 0
    xl.ActiveWindow.ScrollColumn = 1
    xl.ActiveWindow.ScrollRow = 1
    xl.ActiveWindow.DisplayGridlines = False
    xl.ActiveWindow.Zoom = 100
Next

wb.Worksheets("Import").Activate
wb.Worksheets("Import").Range("A1").Select

For Each ws In wb.Worksheets
    On Error Resume Next
    ws.Protect "daisy2026", True, True, True, True
    ws.EnableSelection = xlUnlockedCells
    On Error GoTo 0
Next

wb.Save
wb.Close False
xl.Quit
WScript.Echo "OK: planilha principal limpa."
