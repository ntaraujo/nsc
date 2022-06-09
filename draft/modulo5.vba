Sub AprovaçãoN1()
'
' AprovaçãoN1 Macro
'
    Application.CutCopyMode = False
    Sheets("SPOT_2022").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    Cells.EntireColumn.Hidden = False

    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=8, Criteria1:="N1"
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=7, Operator:=xlFilterNoFill
    Range("AI3:AI400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("U3:U400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("F3:F400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("P3:P400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("D3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("T3:T400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("E3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub AprovaçãoN2()
'
' AprovaçãoN2 Macro
'
    Application.CutCopyMode = False
    Sheets("SPOT_2022").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    Cells.EntireColumn.Hidden = False

    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=8, Criteria1:="N2"
    Range("AI3:AI400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("A35").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("U3:U400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("B35").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("F3:F400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("C35").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("P3:P400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("D35").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("SPOT_2022").Select
    Range("T3:T400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("E35").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub AprovaçãoNBS()
'
' AprovaçãoNBS Macro
'
    Application.CutCopyMode = False
    Sheets("SPOT_2022").Select
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilter.ShowAllData
    Cells.EntireColumn.Hidden = False

    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=8, Criteria1:="NBS"
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=9, Criteria1:="<>"
    ActiveSheet.Range("$A$2:$XFC$400").AutoFilter Field:=7, Operator:=xlFilterNoFill
    Range("AI3:AI400").Select
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
    Sheets("SPOT_2022").Select
    Range("AH3:AH400").Select
    Selection.Copy
    Sheets("APROVAÇÃO").Select
    Range("E67").Select
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
    Sheets("Ajudador4").Cells.Copy
    Sheets("APROVAÇÃO").Select
    Cells.Select
    Range("A1").Activate
    ActiveSheet.Paste

    Range("A66:E66").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Dim myAnswer As Variant

    myAnswer = MsgBox("Checar N1?", vbYesNo, "Macro de Aprovações")
    If myAnswer = vbYes Then Call AprovaçãoN1
    myAnswer = MsgBox("Checar N2?", vbYesNo, "Macro de Aprovações")
    If myAnswer = vbYes Then Call AprovaçãoN2
    myAnswer = MsgBox("Checar NBS?", vbYesNo, "Macro de Aprovações")
    If myAnswer = vbYes Then Call AprovaçãoNBS

    Cells.Select
    Cells.EntireColumn.AutoFit
End Sub
