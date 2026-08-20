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
