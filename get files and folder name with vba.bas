
Option Explicit

Sub ListAllFilesAndFolders()
    Dim folderPath As String
    Dim ws As Worksheet
    Dim rowNum As Long
    
    ' Choose folder/drive
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder or Drive"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    
    Set ws = ActiveSheet
    ws.Cells.Clear
    
    ' Headers
    ws.Range("A1").Value = "Type"
    ws.Range("B1").Value = "Name"
    ws.Range("C1").Value = "Full Path"
    ws.Range("D1").Value = "Extension"
    
    ' Format header
    With ws.Range("A1:D1")
        .Interior.Color = RGB(200, 230, 250) ' Light blue
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    rowNum = 2
    ' Start recursive listing
    Call RecursiveList(folderPath, ws, rowNum)
    
    ' Apply table-like formatting
    With ws.Range("A1:D" & rowNum - 1)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Columns.AutoFit
    End With
    
    ' Zebra stripes
    Dim i As Long
    For i = 2 To rowNum - 1 Step 2
        ws.Range("A" & i & ":D" & i).Interior.Color = RGB(242, 242, 242) ' Light gray
    Next i
    
    MsgBox "Files and folders listed successfully!", vbInformation
End Sub

Private Sub RecursiveList(ByVal folderPath As String, ByRef ws As Worksheet, ByRef rowNum As Long)
    Dim fso As Object, folder As Object, subFolder As Object, fileItem As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' List folder itself (highlight)
    ws.Cells(rowNum, 1).Value = "Folder"
    ws.Cells(rowNum, 2).Value = folder.Name
    ws.Cells(rowNum, 3).Value = folder.Path
    ws.Cells(rowNum, 4).Value = "" ' no extension for folders
    
    With ws.Rows(rowNum)
        .Font.Bold = True
        .Interior.Color = RGB(230, 240, 250) ' light blue-gray
    End With
    
    rowNum = rowNum + 1
    
    ' List files in this folder
    For Each fileItem In folder.Files
        ws.Cells(rowNum, 1).Value = "File"
        ws.Cells(rowNum, 2).Value = fileItem.Name
        ws.Cells(rowNum, 3).Value = fileItem.Path
        ws.Cells(rowNum, 4).Value = fso.GetExtensionName(fileItem.Path)
        rowNum = rowNum + 1
    Next fileItem
    
    ' Recurse into subfolders
    For Each subFolder In folder.SubFolders
        Call RecursiveList(subFolder.Path, ws, rowNum)
    Next subFolder
End Sub



