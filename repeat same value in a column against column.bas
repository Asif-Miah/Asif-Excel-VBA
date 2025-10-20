Attribute VB_Name = "Module3"
Function FillRepeated(keyRange As Range, fillRange As Range, lookupValue As Variant) As Variant
    Dim i As Long
    
    ' Loop through all rows
    For i = 1 To keyRange.Rows.Count
        If keyRange.Cells(i, 1).Value = lookupValue Then
            If fillRange.Cells(i, 1).Value <> "" Then
                FillRepeated = fillRange.Cells(i, 1).Value
                Exit Function
            End If
        End If
    Next i
    
    ' If nothing found, return blank
    FillRepeated = ""
End Function

