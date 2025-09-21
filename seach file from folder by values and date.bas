Attribute VB_Name = "Module1"

Sub callall()
AddSearchLabels
End Sub





Sub AddSearchLabels()
    With ActiveSheet
        ' Set values
        .Range("A1").Value = "Search by value"
        .Range("B1").Value = "Search by date"
        .Range("A2:B2").Value = ""  ' Keep second row empty for now (optional)
        
        ' Format A1
        With .Range("A1")
            .Interior.Color = RGB(52, 152, 219)   ' Blue background
            .Font.Color = RGB(255, 255, 255)      ' White text
            .Font.Bold = True
            .Font.Size = 20
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' Format B1
        With .Range("B1")
            .Interior.Color = RGB(46, 204, 113)   ' Green background
            .Font.Color = RGB(255, 255, 255)      ' White text
            .Font.Bold = True
            .Font.Size = 20
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' Format A1:B2 borders
        With .Range("A1:B2").Borders
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        
        ' Make A1:B2 bold and set font size
        With .Range("A1:B2").Font
            .Bold = True
            .Size = 20
        End With
        
        ' Adjust column width for better look
        .Columns("A:B").AutoFit
    End With
End Sub





Private Sub Worksheet_Change(ByVal Target As Range)
    Dim folderPath As String
    Dim fileName As String
    Dim rowNum As Long
    Dim searchString As String
    Dim startPos As Long
    Dim fileExt As String
    Dim fileDate As Date
    Dim ws As Worksheet
    Dim filterDate As Date
    Dim applyDateFilter As Boolean
    
    Set ws = Me
    
    ' Only trigger if A2 or B2 changes
    If Not Intersect(Target, ws.Range("A2:B2")) Is Nothing Then
        
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        ' Clear previous table (starting from A4, columns A:D)
        ws.Range("A4:D1000").Clear
        
        ' Add headers
        With ws.Range("A3:D3")
            .Value = Array("File Name", "File Type", "Date Modified", "Open Link")
            .Font.Bold = True
            .Interior.Color = RGB(79, 129, 189) ' modern blue header
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        ' Enable autofilter
        ws.Range("A3:D3").AutoFilter
        
        ' Folder containing files
        folderPath = "D:\My Documents\Desktop\file\"
        
        ' Get search string from A2
        searchString = ws.Range("A2").Value
        
        ' Get date from B2 (if any)
        If IsDate(ws.Range("B2").Value) Then
            filterDate = ws.Range("B2").Value
            applyDateFilter = True
        Else
            applyDateFilter = False
        End If
        
        ' Start writing from row 4
        rowNum = 4
        
        ' Get first file
        fileName = Dir(folderPath & "*.*")
        
        ' Loop through all files
        Do While fileName <> ""
            
            fileDate = FileDateTime(folderPath & fileName)
            
            ' Show files matching search string and date filter
            If (searchString = "" Or InStr(1, fileName, searchString, vbTextCompare) > 0) And _
               (Not applyDateFilter Or Int(fileDate) >= Int(filterDate)) Then
               
                ' Column A – file name
                ws.Cells(rowNum, 1).Value = fileName
                
                ' Column B – file type
                fileExt = Mid(fileName, InStrRev(fileName, ".") + 1)
                ws.Cells(rowNum, 2).Value = UCase(fileExt)
                
                ' Column C – Date Modified
                ws.Cells(rowNum, 3).Value = fileDate
                ws.Cells(rowNum, 3).NumberFormat = "dd-mmm-yyyy hh:mm"
                
                ' Highlight matching part in Column A
                If searchString <> "" Then
                    startPos = InStr(1, fileName, searchString, vbTextCompare)
                    If startPos > 0 Then
                        With ws.Cells(rowNum, 1).Characters(Start:=startPos, Length:=Len(searchString)).Font
                            .Color = RGB(255, 0, 0)
                            .Bold = True
                        End With
                    End If
                End If
                
                ' Column D – clickable hyperlink
                ws.Hyperlinks.Add _
                    Anchor:=ws.Cells(rowNum, 4), _
                    Address:=folderPath & fileName, _
                    TextToDisplay:="Open File"
                
                ' Alternate row coloring
                If rowNum Mod 2 = 0 Then
                    ws.Range("A" & rowNum & ":D" & rowNum).Interior.Color = RGB(221, 235, 247)
                Else
                    ws.Range("A" & rowNum & ":D" & rowNum).Interior.Color = RGB(255, 255, 255)
                End If
                
                rowNum = rowNum + 1
            End If
            
            ' Next file
            fileName = Dir
        Loop
        
        ' Apply borders
        Dim tblRange As Range
        Set tblRange = ws.Range("A3:D" & rowNum - 1)
        With tblRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
        
        ' Autofit columns
        ws.Columns("A:D").AutoFit
        
        ' Minimum width for readability
        If ws.Columns("A").ColumnWidth < 25 Then ws.Columns("A").ColumnWidth = 25
        If ws.Columns("B").ColumnWidth < 12 Then ws.Columns("B").ColumnWidth = 12
        If ws.Columns("C").ColumnWidth < 18 Then ws.Columns("C").ColumnWidth = 18
        If ws.Columns("D").ColumnWidth < 15 Then ws.Columns("D").ColumnWidth = 15
        
        ' Freeze panes from row 4 (headers stay visible)
        ws.Rows("4:4").Select
        ActiveWindow.FreezePanes = True
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
End Sub


