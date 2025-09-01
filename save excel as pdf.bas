Attribute VB_Name = "Module1"
Sub SaveFullPageAsPDF_AutoFit()
    Dim ws As Worksheet
    Dim FilePath As String
    
    Set ws = ActiveSheet
    
    ' Autofit rows and columns
    ws.Cells.EntireColumn.AutoFit
    ws.Cells.EntireRow.AutoFit
    
    ' Set margins to 0.2 inch
    With ws.PageSetup
        .TopMargin = Application.InchesToPoints(0.2)
        .BottomMargin = Application.InchesToPoints(0.2)
        .LeftMargin = Application.InchesToPoints(0.2)
        .RightMargin = Application.InchesToPoints(0.2)
    End With
    
    ' Clear print area (use whole sheet)
    ws.PageSetup.PrintArea = ""
    
    ' File path (same folder as workbook)
    FilePath = ThisWorkbook.Path & "\" & ws.Name & "_FullPage.pdf"
    
    ' Export as PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                           Filename:=FilePath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, _
                           IgnorePrintAreas:=False, _
                           OpenAfterPublish:=True
    
    MsgBox "Saved full page as PDF: " & FilePath, vbInformation
End Sub


