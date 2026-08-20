# Ordreoversigt.xlsm workbook audit

## 1) VBA modules
### Module1.bas (`VBA/Module1`)
```vb
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
```

### Denne_projektmappe.cls (`VBA/Denne_projektmappe`)
```vb
Attribute VB_Name = "Denne_projektmappe"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark1.cls (`VBA/Ark1`)
```vb
Attribute VB_Name = "Ark1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark21.cls (`VBA/Ark21`)
```vb
Attribute VB_Name = "Ark21"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark13.cls (`VBA/Ark13`)
```vb
Attribute VB_Name = "Ark13"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark4.cls (`VBA/Ark4`)
```vb
Attribute VB_Name = "Ark4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark5.cls (`VBA/Ark5`)
```vb
Attribute VB_Name = "Ark5"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark6.cls (`VBA/Ark6`)
```vb
Attribute VB_Name = "Ark6"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark16.cls (`VBA/Ark16`)
```vb
Attribute VB_Name = "Ark16"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module7.bas (`VBA/Module7`)
```vb
Attribute VB_Name = "Module7"
```

### Ark9.cls (`VBA/Ark9`)
```vb
Attribute VB_Name = "Ark9"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module4.bas (`VBA/Module4`)
```vb
Attribute VB_Name = "Module4"
```

### Ark7.cls (`VBA/Ark7`)
```vb
Attribute VB_Name = "Ark7"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark12.cls (`VBA/Ark12`)
```vb
Attribute VB_Name = "Ark12"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module2.bas (`VBA/Module2`)
```vb
Attribute VB_Name = "Module2"
```

### Ark3.cls (`VBA/Ark3`)
```vb
Attribute VB_Name = "Ark3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark14.cls (`VBA/Ark14`)
```vb
Attribute VB_Name = "Ark14"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module3.bas (`VBA/Module3`)
```vb
Attribute VB_Name = "Module3"
Sub sja()
Attribute sja.VB_ProcData.VB_Invoke_Func = " \n14"
'
' sja Makro
'

    ActiveWorkbook.RefreshAll
    Application.Run "'Prodordre forside.xlsm'!SortRute"

'
    
    Sheets("VareLinier").Select
     Range("L2").Select
    ActiveCell.FormulaR1C1 = "Save"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Range("L2").Select
    Selection.ClearContents
  
     Sheets("Forside").Select
    Range("C15").Select
End Sub
Sub SortRute()
Attribute SortRute.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SortRute Makro
'

'
    Sheets("Rute sort").Select
    Cells.Select
    Selection.ClearContents
    Range("F22").Select
    Sheets("Rute").Select
    Columns("A:W").Select
    Selection.Copy
    Sheets("Rute sort").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("J2").Select
    Application.CutCopyMode = False
'    ActiveWorkbook.Worksheets("Rute sort").Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("Rute sort").Sort.SortFields.Add Key:=Range( _
'        "B2:B1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortNormal
'    ActiveWorkbook.Worksheets("Rute sort").Sort.SortFields.Add Key:=Range( _
'        "P2:P1000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
'     xlSortNormal
    With ActiveWorkbook.Worksheets("Rute sort").Sort
        .SetRange Range("A1:W1000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    

    
    Sheets("VareLinier").Select
    Columns("O:O").Select
    With Selection
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    Rows("8:555").EntireRow.AutoFit


End Sub
Sub close_Awb()
  Application.DisplayAlerts = False
  With ActiveWorkbook
    .Saved = False
    .Close
  Application.Quit

  End With
End Sub
```

### Module10.bas (`VBA/Module10`)
```vb
Attribute VB_Name = "Module10"
```

### Ark10.cls (`VBA/Ark10`)
```vb
Attribute VB_Name = "Ark10"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module5.bas (`VBA/Module5`)
```vb
Attribute VB_Name = "Module5"
Sub Knap6_Klik()
Attribute Knap6_Klik.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Knap6_Klik Makro
'
'  opdatere alt
'
    ActiveWorkbook.RefreshAll
    Application.Run "'Prodordre forside.xlsm'!SortRute"
'
'
   Sheets("Indklst").Select
   ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
    Range("C15").Select
'
   Sheets("Forside").Select
    Range("E12").Select

' lukke

'  Application.DisplayAlerts = False
'  With ActiveWorkbook
'    .Saved = False
'    .Close
' End With

End Sub
```

### Ark17.cls (`VBA/Ark17`)
```vb
Attribute VB_Name = "Ark17"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark20.cls (`VBA/Ark20`)
```vb
Attribute VB_Name = "Ark20"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module6.bas (`VBA/Module6`)
```vb
Attribute VB_Name = "Module6"
```

### Ark18.cls (`VBA/Ark18`)
```vb
Attribute VB_Name = "Ark18"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark19.cls (`VBA/Ark19`)
```vb
Attribute VB_Name = "Ark19"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark22.cls (`VBA/Ark22`)
```vb
Attribute VB_Name = "Ark22"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark15.cls (`VBA/Ark15`)
```vb
Attribute VB_Name = "Ark15"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module8.bas (`VBA/Module8`)
```vb
Attribute VB_Name = "Module8"
```

### Module9.bas (`VBA/Module9`)
```vb
Attribute VB_Name = "Module9"
```

### Ark2.cls (`VBA/Ark2`)
```vb
Attribute VB_Name = "Ark2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark23.cls (`VBA/Ark23`)
```vb
Attribute VB_Name = "Ark23"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module11.bas (`VBA/Module11`)
```vb
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
```

### Module12.bas (`VBA/Module12`)
```vb
Attribute VB_Name = "Module12"
Sub Makro5()
Attribute Makro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro5 Makro
'

'
    ActiveSheet.Paste
End Sub
```

### Module13.bas (`VBA/Module13`)
```vb
Attribute VB_Name = "Module13"
```

### Module14.bas (`VBA/Module14`)
```vb
Attribute VB_Name = "Module14"
```

### Module15.bas (`VBA/Module15`)
```vb
Attribute VB_Name = "Module15"
```

### Ark8.cls (`VBA/Ark8`)
```vb
Attribute VB_Name = "Ark8"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module16.bas (`VBA/Module16`)
```vb
Attribute VB_Name = "Module16"
Public Function UserName()
    UserName = Environ$("UserName")
End Function
```

### Module17.bas (`VBA/Module17`)
```vb
Attribute VB_Name = "Module17"
```

### Ark11.cls (`VBA/Ark11`)
```vb
Attribute VB_Name = "Ark11"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark24.cls (`VBA/Ark24`)
```vb
Attribute VB_Name = "Ark24"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark25.cls (`VBA/Ark25`)
```vb
Attribute VB_Name = "Ark25"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module18.bas (`VBA/Module18`)
```vb
Attribute VB_Name = "Module18"
```

### Module19.bas (`VBA/Module19`)
```vb
Attribute VB_Name = "Module19"
```

### Ark26.cls (`VBA/Ark26`)
```vb
Attribute VB_Name = "Ark26"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark27.cls (`VBA/Ark27`)
```vb
Attribute VB_Name = "Ark27"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module20.bas (`VBA/Module20`)
```vb
Attribute VB_Name = "Module20"
```

### Ark28.cls (`VBA/Ark28`)
```vb
Attribute VB_Name = "Ark28"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module21.bas (`VBA/Module21`)
```vb
Attribute VB_Name = "Module21"
Sub PDF_kopi()
Attribute PDF_kopi.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PDF_kopi Makro
'
Calculate



Dim DataSti, Filnavn As String

DataSti = "P:\Visma\Laserliste\Arkiv\" 'Der hvor filen skal gemmes, husk at afslutte med \


Filnavn = Range("D1").Text

'Tjekker om mappen 'DataSti' eksisterer, hvis ikke oprettes den
If Dir(DataSti, vbDirectory) = "" Then
    MkDir DataSti
End If

'Gemmer den aktive workbook som .pdf
ActiveSheet.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=DataSti & Filnavn, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    From:=1, To:=Sheets.Count, _
    OpenAfterPublish:=False

' MsgBox "Filen er gemt som " & DataSti & Filnavn & ".pdf", vbInformation


'
End Sub
```

### Module22.bas (`VBA/Module22`)
```vb
Attribute VB_Name = "Module22"
```

### Ark29.cls (`VBA/Ark29`)
```vb
Attribute VB_Name = "Ark29"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Ark30.cls (`VBA/Ark30`)
```vb
Attribute VB_Name = "Ark30"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module23.bas (`VBA/Module23`)
```vb
Attribute VB_Name = "Module23"
```

### Module24.bas (`VBA/Module24`)
```vb
Attribute VB_Name = "Module24"
```

### Ark31.cls (`VBA/Ark31`)
```vb
Attribute VB_Name = "Ark31"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

### Module25.bas (`VBA/Module25`)
```vb
Attribute VB_Name = "Module25"
Sub sj()
Attribute sj.VB_ProcData.VB_Invoke_Func = " \n14"
'
' sj Makro
'

'
    Sheets("LaserListe").Select
    Range("Tabel_Forespørgsel_fra_Visma4711614[[#Headers],[ProdNr.]]").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("Forside").Select
End Sub
```

### Ark32.cls (`VBA/Ark32`)
```vb
Attribute VB_Name = "Ark32"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
```

## 2) Power Query queries
No Power Query queries found in this workbook. The file contains classic query tables and connections, not a Power Query / DataMashup package.

## 3) Forside formulas
- AB4: `+Tabel_Forespørgsel_fra_Visma17[Gr4]` (value: `1`)
- AB6: `+Rute!V1` (value: `0`)
- AB7: `+Tabel_Forespørgsel_fra_Visma4711918[ProdNr.]` (value: `1073400460-1`)
- AB8: `+LaserListe!Z1` (value: `10`)
- AB10: `+'L- beksriv'!G3` (value: `0`)
- AB12: `+IndkLst!K1` (value: `1`)

## 4) How the key sheets are built
### VareLinier
- Connection 14 — Forespørgsel fra Visma21
  - Table: Tabel_Forespørgsel_fra_Visma47
  - Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
  - SQL:
```sql
SELECT Ord_1.OrdBasNo AS 'SalgsOrdre', Ord.MainOrd AS 'HovOrd', Ord.OrdNo AS 'ProdOrd', OrdLn.NoInvoAb+OrdLn.NoFin-OrdLn.NoInvo AS 'Ant.', OrdLn.ProdNo AS 'ProdNr.', Ord.OrdBasNo AS 'OrdGrNo', OrdLn.Descr AS 'Beskrivelse', Ord_1.MainOrd AS 'main', Prod.Inf7 AS 'TegnNr', OrdLn.Un, OrdLn.TrInf1
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod
WHERE Ord.OrdNo = OrdLn.OrdNo AND Prod.ProdNo = OrdLn.ProdNo AND Ord.MainOrd = Ord_1.OrdNo AND ((OrdLn.TrTp=7) AND (OrdLn.ProdNo Not Like 'V%') AND (Ord_1.OrdBasNo=?) AND (OrdLn.NoInvoAb+OrdLn.NoFin-OrdLn.NoInvo<>.000000))
```

### SaveListe
- Connection 19 — Forespørgsel fra Visma214
  - Table: Tabel_Forespørgsel_fra_Visma4711
  - Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
  - SQL:
```sql
SELECT Ord_1.OrdBasNo AS 'SalgsOrdre', Ord.MainOrd AS 'HovOrd', Ord.OrdNo AS 'ProdOrd', OrdLn.NoInvoAb AS 'Ant.', OrdLn.ProdNo AS 'ProdNr.', Ord.OrdBasNo AS 'OrdGrNo', OrdLn.Descr AS 'Beskrivelse', Ord_1.MainOrd AS 'main', Prod.Inf7
 AS 'TegnNr', OrdLn.TrInf3 AS 'Savelængde', Ord.Gr5
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod
WHERE Ord.OrdNo = OrdLn.OrdNo AND Prod.ProdNo = OrdLn.ProdNo AND Ord.MainOrd = Ord_1.OrdNo AND ((OrdLn.TrTp=7) AND (OrdLn.ProdNo Not Like 'V%') AND (Ord_1.OrdBasNo=?))
```

### LaserListe
- Connection 21 — Forespørgsel fra Visma21411
  - Table: Tabel_Forespørgsel_fra_Visma4711614
  - Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
  - SQL:
```sql
SELECT DISTINCT 
    Ord.OrdNo AS 'ProdOrd', 
    OrdLn.ProdNo AS 'ProdNr.', 
    Prod.Descr AS 'Beskrivelse', 
    Prod.Inf7 AS 'TegnNr', 
    OrdLn.TrInf3 AS 'SavLgd', 
    SUBSTRING(Struct.SubProd, 1, 13) AS 'Råvarenr', 
    SUBSTRING(Prod_1.Descr, 1, 50) AS 'Råvarebetegn.', 
    OrdLn.NoInvoAb AS 'Ant', 
    Prod_2.PictNo AS 'Pict', 
    OrdLn_1.TrInf1 AS 'OplNest', 
    Prod.Inf4 AS 'K.Varenr.', 
    Struct.Descr, 
    Struct_1.NoPerStr AS 'Deling',
    ISNULL(ProdCat.Descr, '-') AS 'retn' -- Sostituire 'N/A' con un valore predefinito se necessario
FROM 
    F0001.dbo.Ord Ord
    LEFT JOIN F0001.dbo.OrdLn OrdLn ON Ord.OrdNo = OrdLn.OrdNo
    LEFT JOIN F0001.dbo.Prod Prod ON Prod.ProdNo = OrdLn.ProdNo
    LEFT JOIN F0001.dbo.Ord Ord_1 ON Ord.MainOrd = Ord_1.OrdNo
    LEFT JOIN F0001.dbo.Struct Struct ON Struct.ProdNo = OrdLn.ProdNo
    LEFT JOIN F0001.dbo.Prod Prod_1 ON Prod_1.ProdNo = Struct.SubProd
    LEFT JOIN F0001.dbo.Struct Struct_1 ON Struct_1.SubProd = OrdLn.ProdNo
    LEFT JOIN F0001.dbo.Prod Prod_2 ON Struct_1.ProdNo = Prod_2.ProdNo
    LEFT JOIN F0001.dbo.OrdLn OrdLn_1 ON OrdLn_1.PurcNo = Ord_1.MainOrd
    LEFT JOIN F0001.dbo.ProdCat ProdCat ON Prod.PrCatNo = ProdCat.PrCatNo
WHERE 
    OrdLn.TrTp = 5 
    AND OrdLn.ProdNo LIKE '%L%' 
    AND Ord_1.OrdBasNo = ? -- Sostituire con il valore effettivo o il parametro
    AND Struct.SubProd LIKE '3%' 
    AND OrdLn.NoInvoAb <> 0.000000 -- Sostituire con il valore corretto se necessario
    AND Prod.Inf5 <> 'komb'
ORDER BY 
    Ord.OrdNo, 
    OrdLn.ProdNo;
```

### IndkLst
- Connection 26 — Forespørgsel fra Visma21412
  - Table: Tabel_Forespørgsel_fra_Visma4711918
  - Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
  - SQL:
```sql
SELECT DISTINCT OrdLn.ProdNo AS 'ProdNr.', OrdLn.Descr AS 'Beskrivelse', Txt.Txt AS 'Enh.', Prod.Gr6, StcBal.PoPhStB AS 'Beholdning', StcBal.InProdO AS 'Reserveret', Prod.Gr5
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod, F0001.dbo.StcBal StcBal, F0001.dbo.Txt Txt
WHERE Ord.OrdNo = OrdLn.OrdNo AND Prod.ProdNo = OrdLn.ProdNo AND Ord.MainOrd = Ord_1.OrdNo AND Prod.StSaleUn = Txt.TxtNo AND StcBal.ProdNo = OrdLn.ProdNo AND ((OrdLn.TrTp=5) AND (Ord_1.OrdBasNo=?) AND (Prod.Gr5 In (2,3,11)) AND (OrdLn.ProdNo Not Like '%L%') AND (Txt.TxtTp=16) AND Txt.Lang=45 AND (Prod.Gr6<>1))

UNION

SELECT OrdLn.ProdNo, OrdLn.Descr, Txt.Txt, Prod.Gr6, StcBal.PoPhStB, StcBal.InProdO, Prod.Gr5
FROM F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod, F0001.dbo.StcBal StcBal, F0001.dbo.Txt Txt
WHERE Prod.ProdNo = OrdLn.ProdNo AND Prod.StSaleUn = Txt.TxtNo AND StcBal.ProdNo = Prod.ProdNo AND ((Txt.TxtTp=16) AND Txt.Lang=45 AND (OrdLn.OrdNo=?) AND (Prod.Gr5=3) AND (Prod.Gr6<>1))
```

### PladeLager
- Connection 22 — Forespørgsel fra Visma214111
  - Table: Tabel_Forespørgsel_fra_Visma47116142
  - Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
  - SQL:
```sql
SELECT DISTINCT Prod_1.R3 AS 'Bærer', Prod.ProdNo AS 'ProdNr', Prod.Descr AS 'Betegnelse', StcBal.PoPhStB AS 'Beholdning', Prod.NWgtU AS 'Pladevægt', Prod.HgtU AS 'Tykkelse ', StcBal.PoPhStB/Prod.NWgtU AS 'Plader  ', (StcBal.ShpRsv+StcBal.ShpRsvIn)/Prod.NWgtU AS 'Resveret ', StcBal.PhCstPr AS 'FIFO pris  '
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod, F0001.dbo.Prod Prod_1, F0001.dbo.StcBal StcBal, F0001.dbo.Struct Struct
WHERE Ord.OrdNo = OrdLn.OrdNo AND Ord.MainOrd = Ord_1.OrdNo AND Struct.ProdNo = OrdLn.ProdNo AND Prod_1.ProdNo = Struct.SubProd AND Prod_1.R3 = Prod.R3 AND StcBal.ProdNo = Prod.ProdNo AND ((OrdLn.TrTp=5) AND (OrdLn.ProdNo Like '%L%') AND (Ord_1.OrdBasNo=?) AND (Struct.SubProd Like '3%') AND (StcBal.StcNo=1) AND (Prod.ProdNo<>'301001' And Prod.ProdNo Like '3%') AND (OrdLn.NoInvoAb<>$0))
ORDER BY Prod.ProdNo
```

### Rute
- Connection 16 — Forespørgsel fra Visma213
  - Table: Tabel_Forespørgsel_fra_Visma4712
  - Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
  - SQL:
```sql
SELECT Ord.OrdNo AS 'OrdNr', OrdLn.ProdNo AS 'ProdNr', OrdLn.NoInvoAb AS 'Ant', Ord.MainOrd AS 'HovedOrdre', Ord.OrdBasNo AS 'OrdGrNo', OrdLn.Descr, Ord_1.MainOrd, Ord_1.OrdBasNo, SUBSTRING(R7.Nm,1,3), Struct_1.TrInf4, Struct_1.Descr, OrdLn.CfDelDt, R7.RNo, R7.Gr11, Struct_1.TrInf3 AS 'PrgNr', R7.Gr3
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.R7 R7, F0001.dbo.Struct Struct, F0001.dbo.Struct Struct_1
WHERE Ord.OrdNo = OrdLn.OrdNo AND Ord.MainOrd = Ord_1.OrdNo AND OrdLn.ProdNo = Struct.ProdNo AND Struct_1.ProdNo = Struct.SubProd AND Struct_1.R7 = R7.RNo AND ((OrdLn.TrTp=7) AND (OrdLn.ProdNo Not Like 'V%') AND (Ord_1.OrdBasNo=?) AND (Struct_1.ProdTp4=1) OR (OrdLn.TrTp=7) AND (OrdLn.ProdNo Not Like 'V%') AND (Ord_1.OrdBasNo=?) AND (Struct_1.ProdTp4=7))
ORDER BY OrdLn.ProdNo, Struct_1.TrInf4
```

### Laser inf fra rute
- Connection 25 — Forespørgsel fra Visma214114
  - Table: Tabel_Forespørgsel_fra_Visma471161433
  - Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
  - SQL:
```sql
SELECT DISTINCT Ord.OrdNo AS 'ProdOrd', OrdLn.ProdNo AS 'ProdNr.', Struct_2.SubProd, Prod.Inf2, R7.Gr3
FROM F0001.dbo.BgtLn BgtLn, F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod, F0001.dbo.R7 R7, F0001.dbo.Struct Struct, F0001.dbo.Struct Struct_1, F0001.dbo.Struct Struct_2
WHERE Ord.OrdNo = OrdLn.OrdNo AND Ord.MainOrd = Ord_1.OrdNo AND OrdLn.ProdNo = Struct.SubProd AND Struct_1.ProdNo = Struct.ProdNo AND Struct_1.SubProd = Struct_2.ProdNo AND Prod.ProdNo = Struct_2.SubProd AND Struct_2.SubProd = BgtLn.ProdNo AND R7.RNo = BgtLn.R7 AND ((OrdLn.TrTp=5) AND (OrdLn.ProdNo Like '%L%') AND (Ord_1.OrdBasNo=?) AND (OrdLn.NoInvoAb<>$.000000) AND (Struct_1.SubProd Like 'V%') AND (Prod.Inf2<>''))
ORDER BY Ord.OrdNo, OrdLn.ProdNo
```

## 5) Workbook connections and data sources
### [Ark1] 34 — Forespørgsel fra Visma91
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=SJ;DATABASE=F0001`
- Table: Tabel_EksterneData_1
- SQL:
```sql
SELECT Prod.ProdNo, Prod.Free1
FROM F0001.dbo.Prod
WHERE Prod.Free1=1
AND Prod.ProdGr <> '99999'
```

### [Cert] 33 — Forespørgsel fra Visma9
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma27
- SQL:
```sql
SELECT Sum(OrdLn.NoInvoAb+OrdLn.NoFin) AS 'Stk.'
FROM {oj F0001.dbo.OrdLn OrdLn LEFT OUTER JOIN F0001.dbo.Ord Ord ON OrdLn.PurcNo = Ord.OrdNo}
WHERE (OrdLn.OrdNo=?) AND (OrdLn.ProdNo='556')
```

### [Dok] 18 — Forespørgsel fra Visma2132
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma47128
- SQL:
```sql
SELECT DISTINCT OrdLn.ProdNo AS 'ProdNr', OrdLn.Descr, Ord_1.OrdBasNo, Doc.DocGr, Doc.FileNm
FROM F0001.dbo.Doc Doc, F0001.dbo.DocLink DocLink, F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn
WHERE Ord.OrdNo = OrdLn.OrdNo AND Ord.MainOrd = Ord_1.OrdNo AND DocLink.ProdNo = OrdLn.ProdNo AND Doc.DocNo = DocLink.DocNo AND Doc.VerNo = DocLink.VerNo AND ((OrdLn.TrTp=7) AND (Ord_1.OrdBasNo=?) OR (OrdLn.TrTp=7) AND (Ord_1.OrdBasNo=?))
ORDER BY OrdLn.ProdNo
```

### [Forside] 2 — Forespørgsel fra Visma1
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma_1
- SQL:
```sql
SELECT OrdLn.LnNo AS 'Line', OrdLn.ProdNo AS 'ProdNr', OrdLn.Descr AS 'Beskrivelse ', Ord.TrTp AS 'Indkøb?', OrdLn.DurDt AS 'U-levDt', OrdLn.DelDt AS 'LevDt', OrdLn.NoInvoAb+OrdLn.NoFin-OrdLn.NoInvo AS 'Stk.', ' ' AS 'Kvitering', OrdLn.Un, OrdLn.PurcNo
FROM {oj F0001.dbo.OrdLn OrdLn LEFT OUTER JOIN F0001.dbo.Ord Ord ON OrdLn.PurcNo = Ord.OrdNo}
WHERE (OrdLn.OrdNo=?) AND (OrdLn.NoInvoAb+OrdLn.NoFin-OrdLn.NoInvo<>.000000) OR (OrdLn.OrdNo=?) AND (OrdLn.ProdNo='')
ORDER BY OrdLn.Srt
```

### [IndkLst] 26 — Forespørgsel fra Visma21412
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma4711918
- SQL:
```sql
SELECT DISTINCT OrdLn.ProdNo AS 'ProdNr.', OrdLn.Descr AS 'Beskrivelse', Txt.Txt AS 'Enh.', Prod.Gr6, StcBal.PoPhStB AS 'Beholdning', StcBal.InProdO AS 'Reserveret', Prod.Gr5
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod, F0001.dbo.StcBal StcBal, F0001.dbo.Txt Txt
WHERE Ord.OrdNo = OrdLn.OrdNo AND Prod.ProdNo = OrdLn.ProdNo AND Ord.MainOrd = Ord_1.OrdNo AND Prod.StSaleUn = Txt.TxtNo AND StcBal.ProdNo = OrdLn.ProdNo AND ((OrdLn.TrTp=5) AND (Ord_1.OrdBasNo=?) AND (Prod.Gr5 In (2,3,11)) AND (OrdLn.ProdNo Not Like '%L%') AND (Txt.TxtTp=16) AND Txt.Lang=45 AND (Prod.Gr6<>1))

UNION

SELECT OrdLn.ProdNo, OrdLn.Descr, Txt.Txt, Prod.Gr6, StcBal.PoPhStB, StcBal.InProdO, Prod.Gr5
FROM F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod, F0001.dbo.StcBal StcBal, F0001.dbo.Txt Txt
WHERE Prod.ProdNo = OrdLn.ProdNo AND Prod.StSaleUn = Txt.TxtNo AND StcBal.ProdNo = Prod.ProdNo AND ((Txt.TxtTp=16) AND Txt.Lang=45 AND (OrdLn.OrdNo=?) AND (Prod.Gr5=3) AND (Prod.Gr6<>1))
```

### [IndkLst detaljeret] 20 — Forespørgsel fra Visma2141
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma47119
- SQL:
```sql
SELECT Ord_1.OrdBasNo AS 'SalgsOrdre', Ord.MainOrd AS 'HovOrd', Ord.OrdNo AS 'ProdOrd', OrdLn.ProdNo AS 'ProdNr.', Ord.OrdBasNo AS 'OrdGrNo', OrdLn.Descr AS 'Beskrivelse', Ord_1.MainOrd AS 'main', Prod.Inf2 AS 'TegnNr', OrdLn.TrInf3 AS 'Savelængde', Prod.Gr5 AS 'type', OrdLn.NoInvoAb AS 'Antal', Prod.StSaleUn AS 'Enh', Txt.Txt
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod, F0001.dbo.Txt Txt
WHERE Ord.OrdNo = OrdLn.OrdNo AND Prod.ProdNo = OrdLn.ProdNo AND Ord.MainOrd = Ord_1.OrdNo AND Prod.StSaleUn = Txt.TxtNo AND ((OrdLn.TrTp=5) AND (Ord_1.OrdBasNo=?) AND (Prod.Gr5 In (2,3,11)) AND (OrdLn.ProdNo Not Like '%L%') AND (Txt.TxtTp=16) AND Txt.Lang=45)
Union

SELECT Ord.OrdNo, Ord.OrdNo, Ord.OrdNo, OrdLn.ProdNo, Ord.OrdNo, OrdLn.Descr, Ord.OrdNo, Prod.Inf2, OrdLn.TrInf3, Prod.Gr5, OrdLn.NoInvoAb, Prod.StSaleUn, Txt.Txt
FROM F0001.dbo.Ord Ord, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod, F0001.dbo.Txt Txt
WHERE Ord.OrdNo = OrdLn.OrdNo AND OrdLn.ProdNo = Prod.ProdNo AND Prod.StSaleUn = Txt.TxtNo AND ((Txt.TxtTp=16) AND Txt.Lang=45 AND (Ord.OrdNo=?) AND (Prod.Gr5=3))
```

### [KontaktPers] 32 — Forespørgsel fra Visma8
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma16
- SQL:
```sql
SELECT Actor.ActNo, Actor.Nm, Actor.Phone, Actor.MobPh, Actor.LiaAct
FROM F0001.dbo.Actor Actor
WHERE (Actor.LiaAct<>0)
```

### [L- beksriv] 24 — Forespørgsel fra Visma214113
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma471161431
- SQL:
```sql
SELECT Ord.OrdNo AS 'ProdOrd', OrdLn.ProdNo AS 'ProdNr.', Prod.Descr AS 'Beskrivelse', ProdDesc.Descr AS 'Ekstra inf.', ProdDesc.LnNo
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod, F0001.dbo.ProdDesc ProdDesc
WHERE Ord.OrdNo = OrdLn.OrdNo AND Prod.ProdNo = OrdLn.ProdNo AND Ord.MainOrd = Ord_1.OrdNo AND ProdDesc.ProdNo = Prod.ProdNo AND ((OrdLn.TrTp=5) AND (OrdLn.ProdNo Like '%L%') AND (Ord_1.OrdBasNo=?) AND (OrdLn.NoInvoAb<>$.000000))
ORDER BY Ord.OrdNo, ProdDesc.LnNo
```

### [LasList KuVaNr] 23 — Forespørgsel fra Visma214112
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma471161430
- SQL:
```sql
SELECT Ord.OrdNo AS 'ProdOrd', OrdLn.ProdNo AS 'ProdNr.', Prod.Descr AS 'Beskrivelse', Prod.Inf7 AS 'TegnNr', OrdLn.TrInf3 AS 'SavLgd', SUBSTRING(Struct.SubProd,1,13) AS 'Råvarenr', SUBSTRING(Prod_1.Descr,1,50) AS 'Råvarebetegn.', OrdLn.NoInvoAb AS 'Ant', Prod_2.PictNo AS 'Pict', OrdLn_1.TrInf1 AS 'OplNest', Prod.Inf4 AS 'K.Varenr.', Struct.Descr, Struct_1.NoPerStr AS 'Deling', Prod.Inf4
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.OrdLn OrdLn_1, F0001.dbo.Prod Prod, F0001.dbo.Prod Prod_1, F0001.dbo.Prod Prod_2, F0001.dbo.Struct Struct, F0001.dbo.Struct Struct_1
WHERE Ord.OrdNo = OrdLn.OrdNo AND Prod.ProdNo = OrdLn.ProdNo AND Ord.MainOrd = Ord_1.OrdNo AND Struct.ProdNo = OrdLn.ProdNo AND Prod_1.ProdNo = Struct.SubProd AND Struct_1.ProdNo = Prod_2.ProdNo AND Struct_1.SubProd = OrdLn.ProdNo AND OrdLn_1.PurcNo = Ord_1.MainOrd AND ((OrdLn.TrTp=5) AND (OrdLn.ProdNo Like '%L%') AND (Ord_1.OrdBasNo=?) AND (Struct.SubProd Like '3%') AND (OrdLn.NoInvoAb<>$.000000))
ORDER BY Ord.OrdNo, OrdLn.ProdNo
```

### [Laser inf fra rute] 25 — Forespørgsel fra Visma214114
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma471161433
- SQL:
```sql
SELECT DISTINCT Ord.OrdNo AS 'ProdOrd', OrdLn.ProdNo AS 'ProdNr.', Struct_2.SubProd, Prod.Inf2, R7.Gr3
FROM F0001.dbo.BgtLn BgtLn, F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod, F0001.dbo.R7 R7, F0001.dbo.Struct Struct, F0001.dbo.Struct Struct_1, F0001.dbo.Struct Struct_2
WHERE Ord.OrdNo = OrdLn.OrdNo AND Ord.MainOrd = Ord_1.OrdNo AND OrdLn.ProdNo = Struct.SubProd AND Struct_1.ProdNo = Struct.ProdNo AND Struct_1.SubProd = Struct_2.ProdNo AND Prod.ProdNo = Struct_2.SubProd AND Struct_2.SubProd = BgtLn.ProdNo AND R7.RNo = BgtLn.R7 AND ((OrdLn.TrTp=5) AND (OrdLn.ProdNo Like '%L%') AND (Ord_1.OrdBasNo=?) AND (OrdLn.NoInvoAb<>$.000000) AND (Struct_1.SubProd Like 'V%') AND (Prod.Inf2<>''))
ORDER BY Ord.OrdNo, OrdLn.ProdNo
```

### [LaserListe] 21 — Forespørgsel fra Visma21411
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma4711614
- SQL:
```sql
SELECT DISTINCT 
    Ord.OrdNo AS 'ProdOrd', 
    OrdLn.ProdNo AS 'ProdNr.', 
    Prod.Descr AS 'Beskrivelse', 
    Prod.Inf7 AS 'TegnNr', 
    OrdLn.TrInf3 AS 'SavLgd', 
    SUBSTRING(Struct.SubProd, 1, 13) AS 'Råvarenr', 
    SUBSTRING(Prod_1.Descr, 1, 50) AS 'Råvarebetegn.', 
    OrdLn.NoInvoAb AS 'Ant', 
    Prod_2.PictNo AS 'Pict', 
    OrdLn_1.TrInf1 AS 'OplNest', 
    Prod.Inf4 AS 'K.Varenr.', 
    Struct.Descr, 
    Struct_1.NoPerStr AS 'Deling',
    ISNULL(ProdCat.Descr, '-') AS 'retn' -- Sostituire 'N/A' con un valore predefinito se necessario
FROM 
    F0001.dbo.Ord Ord
    LEFT JOIN F0001.dbo.OrdLn OrdLn ON Ord.OrdNo = OrdLn.OrdNo
    LEFT JOIN F0001.dbo.Prod Prod ON Prod.ProdNo = OrdLn.ProdNo
    LEFT JOIN F0001.dbo.Ord Ord_1 ON Ord.MainOrd = Ord_1.OrdNo
    LEFT JOIN F0001.dbo.Struct Struct ON Struct.ProdNo = OrdLn.ProdNo
    LEFT JOIN F0001.dbo.Prod Prod_1 ON Prod_1.ProdNo = Struct.SubProd
    LEFT JOIN F0001.dbo.Struct Struct_1 ON Struct_1.SubProd = OrdLn.ProdNo
    LEFT JOIN F0001.dbo.Prod Prod_2 ON Struct_1.ProdNo = Prod_2.ProdNo
    LEFT JOIN F0001.dbo.OrdLn OrdLn_1 ON OrdLn_1.PurcNo = Ord_1.MainOrd
    LEFT JOIN F0001.dbo.ProdCat ProdCat ON Prod.PrCatNo = ProdCat.PrCatNo
WHERE 
    OrdLn.TrTp = 5 
    AND OrdLn.ProdNo LIKE '%L%' 
    AND Ord_1.OrdBasNo = ? -- Sostituire con il valore effettivo o il parametro
    AND Struct.SubProd LIKE '3%' 
    AND OrdLn.NoInvoAb <> 0.000000 -- Sostituire con il valore corretto se necessario
    AND Prod.Inf5 <> 'komb'
ORDER BY 
    Ord.OrdNo, 
    OrdLn.ProdNo;
```

### [LevAlt] 9 — Forespørgsel fra Visma16
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma28
- SQL:
```sql
SELECT DelAlt.ProdNo, DelAlt.SupNo, DelAlt.SupProd, DelAlt.SupDescr, Actor.Shrt, DelAlt.Srt
FROM F0001.dbo.Actor Actor, F0001.dbo.DelAlt DelAlt
WHERE Actor.SupNo = DelAlt.SupNo AND ((DelAlt.SupNo<>0))
```

### [OpfyldVarer] 10 — Forespørgsel fra Visma17
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma29
- SQL:
```sql
SELECT OrdLn.OrdNo, OrdLn.ProdNo, StcBal.PoPhStB
FROM F0001.dbo.OrdLn OrdLn, F0001.dbo.StcBal StcBal
WHERE StcBal.ProdNo = OrdLn.ProdNo AND ((OrdLn.OrdNo=50) AND (OrdLn.ProdNo Like '%L%'))
```

### [OppTil sum] 13 — Forespørgsel fra Visma2
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma4
- SQL:
```sql
SELECT OrdLn.ProdNo, OrdLn.NoInvoAb, Ord.MainOrd, Ord.OrdBasNo, Ord.OrdNo, OrdLn.CfDelDt, OrdLn.Descr, OrdLn.LnNo, OrdLn.R7, R7.Nm, R7.Gr3, FreeInf1.Dt1, FreeInf1.Val1
FROM F0001.dbo.FreeInf1 FreeInf1, F0001.dbo.Ord Ord, F0001.dbo.OrdLn OrdLn, F0001.dbo.R7 R7
WHERE Ord.OrdNo = OrdLn.OrdNo AND R7.RNo = OrdLn.R7 AND FreeInf1.OrdNo = OrdLn.OrdNo AND OrdLn.LnNo = FreeInf1.OrdLnNo AND ((OrdLn.ProdNo Like 'R%') AND (Ord.OrdBasNo=?) AND (OrdLn.NoInvoAb<>$0) AND (OrdLn.ProdNo<>'R8200'))
ORDER BY FreeInf1.Dt1
```

### [PladeLager] 22 — Forespørgsel fra Visma214111
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma47116142
- SQL:
```sql
SELECT DISTINCT Prod_1.R3 AS 'Bærer', Prod.ProdNo AS 'ProdNr', Prod.Descr AS 'Betegnelse', StcBal.PoPhStB AS 'Beholdning', Prod.NWgtU AS 'Pladevægt', Prod.HgtU AS 'Tykkelse ', StcBal.PoPhStB/Prod.NWgtU AS 'Plader  ', (StcBal.ShpRsv+StcBal.ShpRsvIn)/Prod.NWgtU AS 'Resveret ', StcBal.PhCstPr AS 'FIFO pris  '
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod, F0001.dbo.Prod Prod_1, F0001.dbo.StcBal StcBal, F0001.dbo.Struct Struct
WHERE Ord.OrdNo = OrdLn.OrdNo AND Ord.MainOrd = Ord_1.OrdNo AND Struct.ProdNo = OrdLn.ProdNo AND Prod_1.ProdNo = Struct.SubProd AND Prod_1.R3 = Prod.R3 AND StcBal.ProdNo = Prod.ProdNo AND ((OrdLn.TrTp=5) AND (OrdLn.ProdNo Like '%L%') AND (Ord_1.OrdBasNo=?) AND (Struct.SubProd Like '3%') AND (StcBal.StcNo=1) AND (Prod.ProdNo<>'301001' And Prod.ProdNo Like '3%') AND (OrdLn.NoInvoAb<>$0))
ORDER BY Prod.ProdNo
```

### [Planlægn] 1 — Forespørgsel fra Visma
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma
- SQL:
```sql
SELECT FreeInf1.R7 AS 'RessNo ', FreeInf1.Dt1 AS 'Dato ', FreeInf1.OrdNo AS 'OdrNo ', FreeInf1.OrdLnNo AS 'Linje ', Ord.Nm AS 'Kunde ', Ord.DelDt AS 'LevDato', FreeInf1.Val1*-1 AS 'Tid ', OrdLn.TransGr3 AS 'Status', Ord.OrdNo, OrdLn.R7, R7.Gr3, FreeInf1.Dt1, OrdLn.Descr
FROM F0001.dbo.FreeInf1 FreeInf1, F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.Ord Ord_2, F0001.dbo.OrdLn OrdLn, F0001.dbo.R7 R7
WHERE OrdLn.OrdNo = FreeInf1.OrdNo AND OrdLn.LnNo = FreeInf1.OrdLnNo AND OrdLn.OrdNo = Ord_1.OrdNo AND Ord_2.OrdNo = Ord_1.MainOrd AND Ord.OrdNo = Ord_2.OrdBasNo AND OrdLn.R7 = R7.RNo AND ((Ord_1.TransGr2<=30) AND (OrdLn.TransGr3<80) AND (Ord.OrdNo=?))
ORDER BY FreeInf1.Dt1
```

### [Ress..] 4 — Forespørgsel fra Visma11
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001;AutoTranslate=No;QuotedId=No;AnsiNPW=No`
- Table: Tabel_Forespørgsel_fra_Visma_121
- SQL:
```sql
SELECT Txt.Lang, Txt.TxtTp, Txt.TxtNo, Txt.Txt
FROM F0001.dbo.Txt Txt
WHERE (Txt.Lang=45) AND (Txt.TxtTp=16)
```

### [Ress..] 29 — Forespørgsel fra Visma5
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001;AutoTranslate=No;QuotedId=No;AnsiNPW=No`
- Table: Tabel_Forespørgsel_fra_Visma13
- SQL:
```sql
SELECT Prod.ProdNo, Prod.Descr
FROM F0001.dbo.Prod Prod
WHERE (Prod.ProdNo Like 'R%')
```

### [Rute] 16 — Forespørgsel fra Visma213
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma4712
- SQL:
```sql
SELECT Ord.OrdNo AS 'OrdNr', OrdLn.ProdNo AS 'ProdNr', OrdLn.NoInvoAb AS 'Ant', Ord.MainOrd AS 'HovedOrdre', Ord.OrdBasNo AS 'OrdGrNo', OrdLn.Descr, Ord_1.MainOrd, Ord_1.OrdBasNo, SUBSTRING(R7.Nm,1,3), Struct_1.TrInf4, Struct_1.Descr, OrdLn.CfDelDt, R7.RNo, R7.Gr11, Struct_1.TrInf3 AS 'PrgNr', R7.Gr3
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.R7 R7, F0001.dbo.Struct Struct, F0001.dbo.Struct Struct_1
WHERE Ord.OrdNo = OrdLn.OrdNo AND Ord.MainOrd = Ord_1.OrdNo AND OrdLn.ProdNo = Struct.ProdNo AND Struct_1.ProdNo = Struct.SubProd AND Struct_1.R7 = R7.RNo AND ((OrdLn.TrTp=7) AND (OrdLn.ProdNo Not Like 'V%') AND (Ord_1.OrdBasNo=?) AND (Struct_1.ProdTp4=1) OR (OrdLn.TrTp=7) AND (OrdLn.ProdNo Not Like 'V%') AND (Ord_1.OrdBasNo=?) AND (Struct_1.ProdTp4=7))
ORDER BY OrdLn.ProdNo, Struct_1.TrInf4
```

### [Rute forside] 12 — Forespørgsel fra Visma19
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2016;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma34
- SQL:
```sql
SELECT DISTINCT Ord.OrdBasNo, Ord.OrdNo, OrdLn.TrInf4, OrdLn.R7, OrdLn.ProdNo
FROM F0001.dbo.Ord Ord, F0001.dbo.OrdLn OrdLn
WHERE Ord.OrdNo = OrdLn.OrdNo AND ((Ord.Gr3=0) AND (Ord.OrdBasNo=?) AND (OrdLn.R7<>'' And OrdLn.R7<>'82' And OrdLn.R7<>'90') AND (OrdLn.ProdNo Like 'R%') OR (Ord.Gr3=0) AND (Ord.OrdBasNo=147999) AND (OrdLn.R7<>'' And OrdLn.R7<>'82' And OrdLn.R7<>'90') AND (OrdLn.ProdNo Like 'R%'))
ORDER BY Ord.OrdNo, OrdLn.TrInf4
```

### [Rute2] 17 — Forespørgsel fra Visma2131
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma47125
- SQL:
```sql
SELECT Ord.OrdNo AS 'OrdNr', OrdLn.ProdNo AS 'ProdNr', OrdLn.NoInvoAb AS 'Ant', Ord.MainOrd AS 'HovedOrdre', Ord.OrdBasNo AS 'OrdGrNo', OrdLn.Descr, Ord_1.MainOrd, Ord_1.OrdBasNo, SUBSTRING(R7.Nm,1,3), Struct_1.TrInf4, Struct_1.Descr, OrdLn.CfDelDt, R7.RNo
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.R7 R7, F0001.dbo.Struct Struct, F0001.dbo.Struct Struct_1
WHERE Ord.OrdNo = OrdLn.OrdNo AND Ord.MainOrd = Ord_1.OrdNo AND OrdLn.ProdNo = Struct.ProdNo AND Struct_1.ProdNo = Struct.SubProd AND Struct_1.R7 = R7.RNo AND ((OrdLn.TrTp=7) AND (OrdLn.ProdNo Not Like 'V%') AND (Ord_1.OrdBasNo=?) AND (Struct_1.ProdTp4=1) AND (R7.RNo='90') OR (OrdLn.TrTp=7) AND (OrdLn.ProdNo Not Like 'V%') AND (Ord_1.OrdBasNo=?) AND (Struct_1.ProdTp4=7) AND (R7.RNo='90'))
ORDER BY OrdLn.ProdNo, Struct_1.TrInf4
```

### [Råvare] 15 — Forespørgsel fra Visma212
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma4710
- SQL:
```sql
SELECT Ord.OrdNo AS 'OrdNr', OrdLn.ProdNo AS 'ProdNr', OrdLn.NoInvoAb AS 'Ant', Ord.MainOrd AS 'HovedOrdre', Ord.OrdBasNo AS 'OrdGrNo', OrdLn.Descr, Ord_1.MainOrd, Ord_1.OrdBasNo, OrdLn.TrInf3
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn
WHERE Ord.OrdNo = OrdLn.OrdNo AND Ord.MainOrd = Ord_1.OrdNo AND ((OrdLn.TrTp=5) AND (OrdLn.ProdNo Like '3%') AND (Ord_1.OrdBasNo=?) OR (OrdLn.TrTp=5) AND (OrdLn.ProdNo Like '6%') AND (Ord_1.OrdBasNo=?))
```

### [SO-hoved] 8 — Forespørgsel fra Visma15
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001;AutoTranslate=No;QuotedId=No;AnsiNPW=No`
- Table: Tabel_Forespørgsel_fra_Visma_125
- SQL:
```sql
SELECT Txt.Lang, Txt.TxtTp, Txt.TxtNo, Txt.LnNo, Txt.Txt
FROM F0001.dbo.Txt Txt
WHERE (Txt.Lang=45) AND (Txt.TxtTp=5)
```

### [SO-hoved] 27 — Forespørgsel fra Visma3
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma_26
- SQL:
```sql
SELECT Ord.OrdNo, Ord.CustNo, Ord.Nm, Actor_1.Nm, Actor_1.Ad1, Actor_1.Ad2, Actor_1.PNo, Actor_1.PArea
FROM F0001.dbo.Actor Actor_1, F0001.dbo.Ord Ord
WHERE Ord.ShpActNo = Actor_1.ActNo AND ((Ord.OrdNo=?))
```

### [SO-hoved] 30 — Forespørgsel fra Visma6
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2013;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma_326
- SQL:
```sql
SELECT Ord.OrdNo, ActInf.Qty1, ActInf.LnNo, ActInf.Txt1, 'Yes'
FROM F0001.dbo.ActInf ActInf, F0001.dbo.Actor Actor, F0001.dbo.Ord Ord
WHERE ActInf.ActNo = Actor.ActNo AND Ord.CustNo = Actor.CustNo AND ((Ord.OrdNo=?) AND (ActInf.InfTp=10))

UNION

SELECT DISTINCT 10, ActInf.Qty1, ActInf.LnNo, ActInf.Txt1, 'Noo'
FROM F0001.dbo.ActInf ActInf, F0001.dbo.Actor Actor, F0001.dbo.Ord Ord
WHERE ActInf.ActNo = Actor.ActNo AND ((Actor.CustNo=75757522))
```

### [SO-hoved] 31 — Forespørgsel fra Visma7
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma17
- SQL:
```sql
SELECT DISTINCT Ord.OrdNo, Ord.CustNo, Ord.Nm, Ord.ReqNo, Ord.LiaActNo, Ord.SelBuy, Actor_1.Usr, Ord.DelDt, Ord.Ad1, Ord.Ad2, Ord.PNo, Ord.PArea, Ord.Rsp, Ord.CreUsr, Ord.DelMt, Ord.DelNm, Ord.DelAd1, Ord.DelAd2, Ord.DelPNo, Ord.DelPArea, Ord.Gr4, Ord.Inf2, Ord.ShpActNo,Ord.DelTrm
FROM F0001.dbo.Actor Actor_1, F0001.dbo.Ord Ord
WHERE Actor_1.EmpNo = Ord.SelBuy AND ((Ord.OrdNo=?))
```

### [SaveListe] 19 — Forespørgsel fra Visma214
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma4711
- SQL:
```sql
SELECT Ord_1.OrdBasNo AS 'SalgsOrdre', Ord.MainOrd AS 'HovOrd', Ord.OrdNo AS 'ProdOrd', OrdLn.NoInvoAb AS 'Ant.', OrdLn.ProdNo AS 'ProdNr.', Ord.OrdBasNo AS 'OrdGrNo', OrdLn.Descr AS 'Beskrivelse', Ord_1.MainOrd AS 'main', Prod.Inf7
 AS 'TegnNr', OrdLn.TrInf3 AS 'Savelængde', Ord.Gr5
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod
WHERE Ord.OrdNo = OrdLn.OrdNo AND Prod.ProdNo = OrdLn.ProdNo AND Ord.MainOrd = Ord_1.OrdNo AND ((OrdLn.TrTp=7) AND (OrdLn.ProdNo Not Like 'V%') AND (Ord_1.OrdBasNo=?))
```

### [Txt] 3 — Forespørgsel fra Visma10
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma_120
- SQL:
```sql
SELECT ProdDesc.Srt, ProdDesc.ProdNo, ProdDesc.Descr
FROM F0001.dbo.ProdDesc ProdDesc
WHERE (ProdDesc.Srt=2) AND (ProdDesc.LangNo=45)
```

### [Txt] 5 — Forespørgsel fra Visma12
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma_2
- SQL:
```sql
SELECT ProdDesc.Srt, ProdDesc.ProdNo, ProdDesc.Descr
FROM F0001.dbo.ProdDesc ProdDesc
WHERE (ProdDesc.Srt=3) AND (ProdDesc.LangNo=45)
```

### [Txt] 6 — Forespørgsel fra Visma13
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma_3
- SQL:
```sql
SELECT ProdDesc.Srt, ProdDesc.ProdNo, ProdDesc.Descr
FROM F0001.dbo.ProdDesc ProdDesc
WHERE (ProdDesc.Srt=4) AND (ProdDesc.LangNo=45)
```

### [Txt] 7 — Forespørgsel fra Visma14
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma_4
- SQL:
```sql
SELECT ProdDesc.Srt, ProdDesc.ProdNo, ProdDesc.Descr
FROM F0001.dbo.ProdDesc ProdDesc
WHERE (ProdDesc.Srt=5) AND (ProdDesc.LangNo=45)
```

### [Txt] 28 — Forespørgsel fra Visma4
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma19
- SQL:
```sql
SELECT ProdDesc.Srt, ProdDesc.ProdNo, ProdDesc.Descr
FROM F0001.dbo.ProdDesc ProdDesc
WHERE (ProdDesc.Srt=1) AND (ProdDesc.LangNo=45)
```

### [Ulev ordre] 11 — Forespørgsel fra Visma18
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2016;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma32
- SQL:
```sql
SELECT Ord.TrTp, Ord.OrdNo, Ord_2.OrdBasNo
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.Ord Ord_2
WHERE Ord.OrdBasNo = Ord_1.OrdNo AND Ord_1.MainOrd = Ord_2.OrdNo AND ((Ord_2.OrdBasNo>?) AND (Ord.TrTp=6))
```

### [VareLinier] 14 — Forespørgsel fra Visma21
- Data source: `DSN=Visma;UID=sj;Trusted_Connection=Yes;APP=Microsoft Office 2010;WSID=SJ;DATABASE=F0001`
- Table: Tabel_Forespørgsel_fra_Visma47
- SQL:
```sql
SELECT Ord_1.OrdBasNo AS 'SalgsOrdre', Ord.MainOrd AS 'HovOrd', Ord.OrdNo AS 'ProdOrd', OrdLn.NoInvoAb+OrdLn.NoFin-OrdLn.NoInvo AS 'Ant.', OrdLn.ProdNo AS 'ProdNr.', Ord.OrdBasNo AS 'OrdGrNo', OrdLn.Descr AS 'Beskrivelse', Ord_1.MainOrd AS 'main', Prod.Inf7 AS 'TegnNr', OrdLn.Un, OrdLn.TrInf1
FROM F0001.dbo.Ord Ord, F0001.dbo.Ord Ord_1, F0001.dbo.OrdLn OrdLn, F0001.dbo.Prod Prod
WHERE Ord.OrdNo = OrdLn.OrdNo AND Prod.ProdNo = OrdLn.ProdNo AND Ord.MainOrd = Ord_1.OrdNo AND ((OrdLn.TrTp=7) AND (OrdLn.ProdNo Not Like 'V%') AND (Ord_1.OrdBasNo=?) AND (OrdLn.NoInvoAb+OrdLn.NoFin-OrdLn.NoInvo<>.000000))
```

