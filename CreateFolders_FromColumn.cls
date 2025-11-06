VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CreateFoldersFromColumn()
    Dim ws As Worksheet
    Dim folderPath As String
    Dim cell As Range
    Dim folderName As String
    Dim fDialog As FileDialog
    
    '?? Use the active sheet
    Set ws = ActiveSheet
    
    '?? Ask user to pick a folder path
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Select a folder where new folders will be created"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "No folder selected. Operation cancelled.", vbExclamation
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
        If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    End With
    
    '?? Loop through each used cell in column A (change column if needed)
    For Each cell In ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
        folderName = Trim(cell.Value)
        If folderName <> "" Then
            ' Check if folder already exists
            If Dir(folderPath & folderName, vbDirectory) = "" Then
                MkDir folderPath & folderName
            End If
        End If
    Next cell
    
    MsgBox "Folders created successfully in:" & vbCrLf & folderPath, vbInformation
End Sub

