Attribute VB_Name = "Module1"

Sub colorizeRanks()
For i = 1 To Sheets.Count
Sheets(i).Activate
Sheets(i).Select
'change all values to number format for coloring
    Range("c2:c27").Select
     Selection.NumberFormat = "General"
     Selection.Value = Selection.Value
     Selection.NumberFormat = "0.0"
      
    Range("f2:f27").Select
     Selection.NumberFormat = "General"
     Selection.Value = Selection.Value
     Selection.NumberFormat = "0.0"
    
    Range("E2:E27").Select
     Selection.NumberFormat = "General"
     Selection.Value = Selection.Value
    
    Range("h2:h27").Select
     Selection.NumberFormat = "General"
     Selection.Value = Selection.Value

'coloring  country-level ranks
    Range("e2:e27,h2:h27").Select
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(1).Value = 1
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 31
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With



    
    Next
End Sub


