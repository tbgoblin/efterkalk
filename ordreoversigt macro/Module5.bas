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
