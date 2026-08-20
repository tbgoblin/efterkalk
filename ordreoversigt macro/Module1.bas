Attribute VB_Name = "Module1"
Sub Knap1_Klik()
Attribute Knap1_Klik.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Knap1_Klik Makro
'

'
'  sj
    
' Test om ordrenummer er en salgsordre

    Sheets("SO-hoved").Select
    Range("Tabel_Forespørgsel_fra_Visma17[CustNo]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("Forside").Select


If Range("Q4") = 0 Then
    Svar = MsgBox(" Indtast ordrenummer igen", vbCritical, "Ordrenr. er ikke en salgsordre")
    End
Else

 End If
        
    
    Sheets("LaserListe").Select
    Range("Tabel_Forespørgsel_fra_Visma4711614[[#Headers],[ProdNr.]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("Forside").Select
   
    
    
    Range("C12").Select
    ActiveWorkbook.RefreshAll
    Application.Run "'Ordreoversigt.xlsm'!SortRute"
    
    Sheets("VareLinier").Select
    Rows("25:100").EntireRow.AutoFit
   
    Sheets("Forside").Select
    Rows("25:100").EntireRow.AutoFit

' tvangsstyre formel i kolonne "O"

    Sheets("Rute2").Select
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-11]=R[1]C[-11],RC[-1]&R[1]C[-1],RC[-1])"
    Sheets("Forside").Select
    
    Rows("25:100").EntireRow.AutoFit
    Range("C7").Select
    
    Range("Tabel_Forespørgsel_fra_Visma_1[[#Headers],[Stk.]]").Select
    ActiveCell.FormulaR1C1 = "Antal"
'    Range("Tabel_Forespørgsel_fra_Visma_1[[#Headers],[ ]]").Select
'    ActiveCell.FormulaR1C1 = "Enh."
    
    Range("C7").Select

' Tjek for Projekt(2), Multi(3) eller Reklamation(5) og indsæt tekst herfor
If Range("AB4") = 2 Then
    ActiveSheet.Shapes.Range(Array("Projekt")).Select
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 7).Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0.75
        .Solid
    End With
    Application.CommandBars("Format Object").Visible = False
Else
   ActiveSheet.Shapes.Range(Array("Projekt")).Select
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 7).Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 1
        .Solid
    End With
    Application.CommandBars("Format Object").Visible = False

    End If

    If Range("AB4") = 3 Then
    ActiveSheet.Shapes.Range(Array("Multi")).Select
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 5).Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0.75
        .Solid
    End With
    Application.CommandBars("Format Object").Visible = False
Else
   ActiveSheet.Shapes.Range(Array("Multi")).Select
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 5).Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 1
        .Solid
    End With
    Application.CommandBars("Format Object").Visible = False

    End If

    If Range("AB4") = 5 Then
    ActiveSheet.Shapes.Range(Array("Reklamation")).Select
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 11).Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0.75
        .Solid
    End With
    Application.CommandBars("Format Object").Visible = False
Else
   ActiveSheet.Shapes.Range(Array("Reklamation")).Select
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 11).Font.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 1
        .Solid
    End With
    Application.CommandBars("Format Object").Visible = False

    End If
    Range("W12:X23").Select
    Selection.Copy
    Range("E12:F12").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'    Juster saveliste
    Sheets("SaveListe").Select
    Rows("4:250").EntireRow.AutoFit
    ActiveSheet.ListObjects("Tabel_Forespørgsel_fra_Visma4711").Range.AutoFilter _
        Field:=5, Criteria1:="=Save*", Operator:=xlOr, Criteria2:="=Pladesaks*"
    Sheets("Forside").Select


'  Leverandør til indkøbsliste
    Sheets("LevAlt").Select
    Cells.Select
    Range("A535").Activate
    Selection.Copy
    Sheets("LevAltSrt").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'  Sortere LevAltSrt
    Sheets("LevAltSrt").Select
    ActiveWorkbook.Worksheets("LevAltSrt").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("LevAltSrt").AutoFilter.Sort.SortFields.Add Key:= _
        Range("A2:A200000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveWorkbook.Worksheets("LevAltSrt").AutoFilter.Sort.SortFields.Add Key:= _
        Range("H2:H200000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("LevAltSrt").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With




'  InkdLst justere liniehøjde
   Sheets("IndkLst").Select
    Rows("6:50").Select
    Rows("6:50").EntireRow.AutoFit
    Range("A1").Select

' Sortere "Laser inf fra rute" (tekster fra rute (prod.Inf2) til brug i Laserliste)
'
    Sheets("Laser inf fra rute SORT").Select
    Columns("A:G").Select
    Selection.ClearContents
    Sheets("Laser inf fra rute").Select
    Columns("A:G").Select
    Selection.Copy
    Sheets("Laser inf fra rute SORT").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("Laser inf fra rute SORT").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Laser inf fra rute SORT").Sort.SortFields.Add Key _
        :=Range("B1:B2000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Laser inf fra rute SORT").Sort.SortFields.Add Key _
        :=Range("F1:F2000"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Laser inf fra rute SORT").Sort
        .SetRange Range("A1:G2000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With



    Sheets("Forside").Select
    
    Rows("25:100").EntireRow.AutoFit
   
    
    Range("C12").Select

    Sheets("LaserListe").Select
    Rows("6:500").EntireRow.AutoFit

    Sheets("Forside").Select

End Sub





Sub Knap2_Klik()
Attribute Knap2_Klik.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Knap2_Klik Makro
'
'  opdatere alt

'
    
' Test om ordrenummer er en salgsordre

    Sheets("SO-hoved").Select
    Range("Tabel_Forespørgsel_fra_Visma17[CustNo]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("Forside").Select


If Range("Q4") = 0 Then
    Svar = MsgBox(" Indtast ordrenummer igen", vbCritical, "Ordrenr. er ikke en salgsordre")
    End
Else

 End If
        
    
    Sheets("LaserListe").Select
    Range("Tabel_Forespørgsel_fra_Visma4711614[[#Headers],[ProdNr.]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("Forside").Select
    
    
    ActiveWorkbook.RefreshAll
    Application.Run "'Ordreoversigt.xlsm'!SortRute"


    Sheets("VareLinier").Select
    Rows("8:83").Select
    Rows("8:83").EntireRow.AutoFit

   Sheets("Forside").Select
   
If Range("AB12") = 0 Then
   Range("F5").Select
    With Selection.Font
        .Color = 15849925
    End With
End If
   
   
   ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    
    
    Sheets("VareLinier").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    
    
'  Tjek for laserskære
   
   Sheets("Forside").Select

If Range("AB8") >= 1 Then

    Sheets("LaserListe").Select
    Range("D2").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Sheets("PladeLager").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
   
End If
   
   
   
   Sheets("Forside").Select

    
    
' Tjek for save/klippe og evt. udskrift

If Range("AB6") >= 1 Then
    
    Range("C4:D4").Select
    ActiveCell.FormulaR1C1 = "Forside save/klippe-liste"
   ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Range("C4:D4").Select
    Selection.ClearContents
   
    Sheets("VareLinier").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
   
    Sheets("SaveListe").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
   
     Sheets("Forside").Select
End If


' tjek for indkøb og evt. udskrift

If Range("AB7") >= 1 Then

   Sheets("Indklst").Select
   ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Range("C15").Select
'
   Sheets("Forside").Select
    Range("C15").Select


End If
    
' Tjek for L-beskrivelse liste

If Range("AB10") >= 1 Then

   Sheets("L- beksriv").Select
   ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Range("C15").Select
'

End If

    Sheets("LaserListe").Select

Range("D2").Value = "(" & Environ("USERNAME") & ")"

    Application.Run "Ordreoversigt.xlsm!PDF_kopi"

   Sheets("Forside").Select
    
If Range("AB12") = 0 Then
   Range("F5").Select
    With Selection.Font
        .Color = 255
    End With
End If

End Sub
Sub Knap4_Klik()
Attribute Knap4_Klik.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Knap4_Klik Makro
'
'  opdatere alt
'
    ActiveWorkbook.RefreshAll
    Application.Run "'Prodordre forside.xlsm'!SortRute"
'
'
   Sheets("Forside").Select
   ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Sheets("VareLinier").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Sheets("LaserListe").Select
    Range("D2").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Sheets("PladeLager").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    
    Sheets("Forside").Select
    Range("C15").Select


' tjek for indkøb og evt. udskrift

If Range("AB7") >= 1 Then

   Sheets("Indklst").Select
   ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Range("C15").Select
'
   Sheets("Forside").Select
    Range("E12").Select


End If


End Sub
