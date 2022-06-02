Public Sub DeleteBlankRows()
    Dim SourceRange As Range
    Dim EntireRow As Range
 
    Set SourceRange = Application.Selection
 
    If Not (SourceRange Is Nothing) Then
        Application.ScreenUpdating = False
 
        For I = SourceRange.Rows.Count To 1 Step -1
            Set EntireRow = SourceRange.Cells(I, 1).EntireRow
            If Application.WorksheetFunction.CountA(EntireRow) = 0 Then
                EntireRow.Delete
            End If
        Next
 
        Application.ScreenUpdating = True
    End If
End Sub
Sub AprovaçãoN2()
'
' AprovaçãoN2 Macro
'
    Application.CutCopyMode = False
    Sheets("SPOT_2022").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=8, Criteria1:="N2"
    Range("E3:E400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("A4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("U3:U400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("F3:F400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("C4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("P3:P400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("D4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub AprovaçãoNBSSâmmya()
'
' AprovaçãoNBS Macro
'
    Application.CutCopyMode = False
    Sheets("SPOT_2022").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=5, Criteria1:="<>DEPÓSITO"
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=8, Criteria1:="NBS"
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=9, Criteria1:="<>"
    Range("E3:E400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("A36").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("I3:I400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("B36").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("F3:F400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("C36").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("P3:P400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("D36").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub AprovaçãoNBSWagner()
'
' AprovaçãoNBS Macro
'
    Application.CutCopyMode = False
    Sheets("SPOT_2022").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=5, Criteria1:="DEPÓSITO"
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=8, Criteria1:="NBS"
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=9, Criteria1:="<>"
    Range("E3:E400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("A67").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("I3:I400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("B67").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("F3:F400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("C67").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("P3:P400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("D67").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub Aprovação()
'
' Aprovação Macro
'
' Atalho do teclado: Ctrl+q
'
    Application.CutCopyMode = False
    Sheets("Ajudador1").Cells.Copy
    Sheets("APROVAÇÃO").Select
    Cells.Select
    Range("A1").Activate
    ActiveSheet.Paste

    Range("A97:D97").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Dim myAnswer As Variant

    myAnswer = MsgBox("Checar N2?", vbYesNo, "Macro de Aprovações")
    If myAnswer = vbYes Then Call AprovaçãoN2
    myAnswer = MsgBox("Checar NBS da Sâmmya?", vbYesNo, "Macro de Aprovações")
    If myAnswer = vbYes Then Call AprovaçãoNBSSâmmya
    myAnswer = MsgBox("Checar NBS do Wagner?", vbYesNo, "Macro de Aprovações")
    If myAnswer = vbYes Then Call AprovaçãoNBSWagner
End Sub

