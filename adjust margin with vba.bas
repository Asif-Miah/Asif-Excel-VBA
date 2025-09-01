
Sub SetMarginsForSelection()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Set print area to current selection
    ws.PageSetup.PrintArea = Selection.Address
    
    ' Set all margins to 0.1 inch
    With ws.PageSetup
        .TopMargin = Application.InchesToPoints(0.1)
        .BottomMargin = Application.InchesToPoints(0.1)
        .LeftMargin = Application.InchesToPoints(0.1)
        .RightMargin = Application.InchesToPoints(0.1)
        .HeaderMargin = Application.InchesToPoints(0.1)
        .FooterMargin = Application.InchesToPoints(0.1)
    End With
    
    MsgBox "Margins set to 0.1 inch for the selection.", vbInformation
End Sub


