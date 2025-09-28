Attribute VB_Name = "Module1"
Function customFormat(inputvalue As Double) As String
customFormat = Format(inputvalue, "#,###.##")
If (Right(customFormat, 1) = ".") Then
    customFormat = Left(customFormat, Len(customFormat) - 1)
    If customFormat = "" Then
        customFormat = 0
    End If
End If
    

End Function
Sub districtFactsheetGenerator()
'=============== Create Objects, Define variables and input data location
Dim ed As Excel.Application
Dim ws As Excel.Worksheet
Dim wb As Excel.Workbook
Dim i, j, k As Integer

Set ed = CreateObject("Excel.Application")
'Set wb = ed.Workbooks.Open("C:\Users\Y.farzi\Desktop\Implementation\results\totRes-color4.xlsx")
Set wb = ed.Workbooks.Open("C:\Users\2740554486\Desktop\emri\facsheet221\Implementation\results\totRes-color7.xlsx")

'=================== Create a factsheet for each province
'=================== based on the opened document having design and table elements

' wb.Sheets.Count each sheet shoud be related to one province data
For i = 0 To 30

'--------------- Map insertion, Croping and province name insertion
ActiveDocument.Tables(1).Cell(2, 2).Range.Text = wb.sheets(CStr(i)).Cells(1, 15).Value
If ActiveDocument.Tables(1).Cell(1, 2).Range.InlineShapes.Count = 1 Then
    ActiveDocument.Tables(1).Cell(1, 2).Range.InlineShapes(1).Delete
End If

ActiveDocument.Tables(1).Cell(1, 2).Range.InlineShapes.AddPicture ("C:\Users\2740554486\Desktop\emri\facsheet221\Implementation\png\" + CStr(i) + ".png")
With ActiveDocument.Tables(1).Cell(1, 2).Range.InlineShapes(1)
    .PictureFormat.CropTop = 70
    .PictureFormat.CropBottom = 80
    .PictureFormat.CropLeft = 120
    .PictureFormat.CropRight = 105
    .Height = 150
    .Width = 180
End With
problems = wb.sheets("provs").Cells(2, "r").Value + wb.sheets(CStr(i)).Cells(1, 15).Value + wb.sheets("provs").Cells(2, "p").Value + " "

probCount = 0
'--------------  Metabolic Table filling
For j = 1 To 16
For k = 3 To 8
    ActiveDocument.Tables(3).Cell(j + 3, k).Range.Text = customFormat(wb.sheets(CStr(i)).Cells(j + 1, k).Value)
    ActiveDocument.Tables(3).Cell(j + 3, k).Range.Shading.BackgroundPatternColor = wb.sheets(CStr(i)).Cells(j + 1, k).DisplayFormat.Interior.Color
    If k = 5 Or k = 8 Then
    probCount = 1
        If customFormat(wb.sheets(CStr(i)).Cells(j + 1, k).Value) > 27 Then
        problems = problems + Left(ActiveDocument.Tables(3).Cell(j + 3, 1).Range.Text, Len(ActiveDocument.Tables(2).Cell(j + 3, 1).Range.Text) - 2)
        
            If k = 5 Then
            problems = problems + wb.sheets("provs").Cells(4, "p").Value
            End If
            If k = 8 Then
            problems = problems + wb.sheets("provs").Cells(3, "p").Value
         End If
         End If
         End If
Next
Next

'-------------- Behavioral Table filling
For j = 1 To 7
For k = 3 To 8
    ActiveDocument.Tables(2).Cell(j + 3, k).Range.Text = customFormat(wb.sheets(CStr(i)).Cells(j + 17, k).Value)
    ActiveDocument.Tables(2).Cell(j + 3, k).Range.Shading.BackgroundPatternColor = wb.sheets(CStr(i)).Cells(j + 17, k).DisplayFormat.Interior.Color
    If k = 5 Or k = 8 Then
    probCount = 1
        If customFormat(wb.sheets(CStr(i)).Cells(j + 17, k).Value) > 27 Then
        problems = problems + Left(ActiveDocument.Tables(2).Cell(j + 3, 1).Range.Text, Len(ActiveDocument.Tables(3).Cell(j + 3, 1).Range.Text) - 2)
        
            If k = 5 Then
            problems = problems + wb.sheets("provs").Cells(4, "p").Value
            End If
            If k = 8 Then
            problems = problems + wb.sheets("provs").Cells(3, "p").Value
         End If
         End If
         End If
Next
Next
'change . to /
    With Selection.Find
        .Text = "."
        .Replacement.Text = "/"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
     
'If probCount <> 0 Then
' ActiveDocument.Tables(4).Cell(1, 1).Range.Text = Left(problems, Len(problems) - 2) + " " + wb.sheets("provs").Cells(2, "n").Value
'Else
'ActiveDocument.Tables(4).Cell(1, 1).Range.Text = wb.sheets("provs").Cells(7, "s").Value + wb.sheets(CStr(i)).Cells(1, 15).Value + " " + wb.sheets("provs").Cells(7, "r").Value
'End If
'modakhelat
'For j = 1 To 12
'For k = 1 To 2
'ActiveDocument.Tables(5).Cell(j + 2, k).Range.Text = wb.Sheets(CStr(i)).Cells(j + 26, k).Value
'ActiveDocument.Tables(5).Cell(j + 2, k + 1).Range.Shading.BackgroundPatternColor = wb.Sheets(CStr(i)).Cells(j + 18, k + 1).DisplayFormat.Interior.Color
'Next
'Next
Debug.Print "province " + CStr(i) + " finished."
ActiveDocument.SaveAs2 FileName:="C:\Users\2740554486\Desktop\emri\facsheet221\Implementation\pdf\v7\factsheet1401-ostandari-" + CStr(i) + "-v1.pdf", FileFormat:=wdFormatPDF
ActiveDocument.SaveAs2 FileName:="C:\Users\2740554486\Desktop\emri\facsheet221\Implementation\pdf\v7\factsheet1401-ostandari-" + CStr(i) + "-v1.docx", FileFormat:=wdFormatDocumentDefault


'End If

Next
wb.Close




End Sub


Sub changeNumbers()
Dim vDir As String
Dim oDoc As Document
Dim fso As New Scripting.FileSystemObject

vDir = "C:\Users\Y.farzi\Desktop\Implementation\pdf\v1\"
vFile = Dir(vDir & "*.docx")

Do While vFile <> ""
Set oDoc = Documents.Open(FileName:=vDir & vFile)
oName = fso.GetBaseName(oDoc.Name)
oDoc.SaveAs2 FileName:="C:\Users\Y.farzi\Desktop\Implementation\pdf\v1\" + Left(oName, Len(oName) - 3) + "-v8.pdf", FileFormat:=wdFormatPDF
oDoc.SaveAs2 FileName:="C:\Users\Y.farzi\Desktop\Implementation\pdf\v1\" + Left(oName, Len(oName) - 3) + "-v8.docx", FileFormat:=wdFormatDocumentDefault

oDoc.Close SaveChanges:=False
vFile = Dir



Loop







End Sub
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
ActiveDocument.Tables(2).Range.Select

    With Selection.Find
        .Text = "."
        .Replacement.Text = "/"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
'
'
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro3"
'
' Macro3 Macro
'
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "WWW/VIZIT/REPORT"
        .Replacement.Text = "WWW.VIZIT.REPORT"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll, Forward:=True
    
    
End Sub
