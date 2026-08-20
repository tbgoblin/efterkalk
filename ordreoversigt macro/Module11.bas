Attribute VB_Name = "Module11"
Sub Makro3()
Attribute Makro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro3 Makro
'

'
    Range("Tabel_Forespørgsel_fra_Visma47128[Kolonne1]").Select
    Selection.Copy
    Application.CutCopyMode = False
End Sub
Sub Makro4()
Attribute Makro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro4 Makro
'

'
    Range("Tabel_Forespørgsel_fra_Visma47128[Kolonne1]").Select
    Selection.Copy
    Range("Tabel_Forespørgsel_fra_Visma47128[Kolonne2]").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Tabel_Forespørgsel_fra_Visma47128[Kolonne2]").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B18").Select
    Application.CutCopyMode = False
    Application.Goto Reference:="Makro4"
    Range("Tabel_Forespørgsel_fra_Visma47128[Kolonne2]").Select
    Selection.ClearContents
    Range("G13").Select
End Sub
