Sub CopyPasteValues()
    ' Copy values from Sheet1 A1:A10
    Sheets("Sheet1").Range("A1:A10").Copy
    
    ' Paste only values (no formatting) into Sheet2 B1:B10
    Sheets("Sheet2").Range("B1").PasteSpecial Paste:=xlPasteValues
    
    ' Clear the clipboard to remove the copy selection border
    Application.CutCopyMode = False
End Sub
