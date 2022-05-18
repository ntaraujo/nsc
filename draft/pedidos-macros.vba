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
    Sheets("SPOT_2022").Select
    Application.CutCopyMode = False
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=8, Criteria1:="N2"
    Range("E3:E400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("A4").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    Sheets("SPOT_2022").Select
    Range("U3:U400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("B4").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    Sheets("SPOT_2022").Select
    Range("F3:F400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("C4").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    Sheets("SPOT_2022").Select
    Range("P3:P400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("D4").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
End Sub
Sub AprovaçãoNBSSâmmya()
'
' AprovaçãoNBS Macro
'
    Sheets("SPOT_2022").Select
    Application.CutCopyMode = False
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=5, Criteria1:="<>DEPÓSITO"
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=8, Criteria1:="NBS"
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=9, Criteria1:="<>"
    Range("E3:E400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("A36").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    Sheets("SPOT_2022").Select
    Range("I3:I400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("B36").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    Sheets("SPOT_2022").Select
    Range("F3:F400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("C36").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    Sheets("SPOT_2022").Select
    Range("P3:P400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("D36").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
End Sub
Sub AprovaçãoNBSWagner()
'
' AprovaçãoNBS Macro
'
    Sheets("SPOT_2022").Select
    Application.CutCopyMode = False
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=5, Criteria1:="DEPÓSITO"
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=8, Criteria1:="NBS"
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=9, Criteria1:="<>"
    Range("E3:E400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("A67").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    Sheets("SPOT_2022").Select
    Range("I3:I400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("B67").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    Sheets("SPOT_2022").Select
    Range("F3:F400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("C67").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    Sheets("SPOT_2022").Select
    Range("P3:P400").Select
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Sheets("APROVAÇÃO").Select
        Range("D67").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
End Sub
Sub Aprovação()
'
' Aprovação Macro
'
' Atalho do teclado: Ctrl+q
'
    Sheets("Ajudador1").Cells.Copy
    Sheets("APROVAÇÃO").Select
    Cells.Select
    Range("A1").Activate
    ActiveSheet.Paste

    Range("A97:D97").Select
    Application.CutCopyMode = False
    Selection.Copy
    If WorksheetFunction.CountA(Selection) <> 0 Then
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End IF
    Call AprovaçãoN2
    Call AprovaçãoNBSSâmmya
    Call AprovaçãoNBSWagner
End Sub
