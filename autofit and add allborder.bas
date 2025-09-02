Attribute VB_Name = "Module3"
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim c As Range
    
    On Error GoTo SafeExit
    Application.EnableEvents = False
    
    ' Autofit the changed rows & columns
    Target.EntireColumn.AutoFit
    Target.EntireRow.AutoFit
    
    ' Loop through changed cells
    For Each c In Target
        If IsEmpty(c.Value) Then
            ' Remove border if cell cleared
            c.Borders.LineStyle = xlNone
        Else
            ' Apply border if cell has value
            With c.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = vbBlack
            End With
        End If
    Next c
    
SafeExit:
    Application.EnableEvents = True
End Sub


