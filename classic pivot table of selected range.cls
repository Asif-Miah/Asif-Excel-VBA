VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CreateClassicPivotTable_NoSubtotal_DateFilter()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim srcData As String
    Dim pivotName As String
    Dim header As String
    Dim firstDataCell As Range
    Dim i As Long
    Dim pf As PivotField
    Dim dateFieldAdded As Boolean
    Dim tblRange As Range
    Dim titleRow As Long
    Dim titleText As String
    
    titleText = "Sales Pivot Report" ' Change this label as needed
    
    ' --- Check selection ---
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a data range first!", vbExclamation
        Exit Sub
    End If
    
    Set wsData = ActiveSheet
    Set dataRange = Selection
    
    If WorksheetFunction.CountA(dataRange) = 0 Or dataRange.Rows.Count < 2 Then
        MsgBox "Please select a valid range including headers.", vbExclamation
        Exit Sub
    End If
    
    ' Convert range to string reference
    srcData = "'" & wsData.Name & "'!" & dataRange.Address(ReferenceStyle:=xlR1C1)
    
    ' Delete old pivot sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("PivotTable_Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create pivot sheet
    Set wsPivot = Sheets.Add
    wsPivot.Name = "PivotTable_Report"
    
    ' Create pivot cache & table
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcData)
    pivotName = "ClassicPivot"
    Set pt = pc.CreatePivotTable(TableDestination:=wsPivot.Range("A3"), TableName:=pivotName)
    
    ' --- Classic Layout ---
    With pt
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
        .ShowDrillIndicators = False
        .HasAutoFormat = False
        .ColumnGrand = False
        .RowGrand = False
        
        ' --- Merge and center row labels automatically ---
        .MergeLabels = True
    End With
    
    ' --- Add fields ---
    dateFieldAdded = False
    For i = 1 To dataRange.Columns.Count
        header = CStr(dataRange.Cells(1, i).Value)
        If header <> "" Then
            Set firstDataCell = dataRange.Cells(2, i)
            On Error Resume Next
            
            If IsDate(firstDataCell.Value) And Not dateFieldAdded Then
                ' Date field ? Filter
                Set pf = pt.PivotFields(header)
                pf.Orientation = xlPageField
                pf.Position = 1
                dateFieldAdded = True
            ElseIf IsNumeric(firstDataCell.Value) Then
                ' Numeric ? Values
                pt.AddDataField pt.PivotFields(header), "Sum of " & header, xlSum
            Else
                ' Text ? Row Field
                Set pf = pt.PivotFields(header)
                pf.Orientation = xlRowField
                pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            End If
            
            On Error GoTo 0
        End If
    Next i
    
    ' --- Clean formatting ---
    wsPivot.Columns.AutoFit
    wsPivot.Cells.Font.Name = "Calibri"
    wsPivot.Cells.Font.Size = 11
    
    ' --- Add borders to entire pivot table ---
    Set tblRange = pt.TableRange2
    If Not tblRange Is Nothing Then
        With tblRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 0
        End With
    End If
    
    ' --- Add merged & centered title above pivot table safely ---
    If Not tblRange Is Nothing Then
        titleRow = tblRange.Row - 1
        If titleRow < 1 Then titleRow = 1 ' Ensure row is valid
        
        wsPivot.Rows(titleRow).Insert Shift:=xlDown
        With wsPivot
            .Range(.Cells(titleRow, 1), .Cells(titleRow, tblRange.Columns.Count)).Merge
            .Cells(titleRow, 1).Value = titleText
            .Cells(titleRow, 1).HorizontalAlignment = xlCenter
            .Cells(titleRow, 1).VerticalAlignment = xlCenter
            .Cells(titleRow, 1).Font.Bold = True
            .Cells(titleRow, 1).Font.Size = 14
            .Rows(titleRow).RowHeight = 25
        End With
    End If
    
    MsgBox "? Classic Pivot created!" & vbCrLf & _
           "• No subtotals on any column" & vbCrLf & _
           "• Date field added to filter (first date column detected)" & vbCrLf & _
           "• Borders applied to entire pivot table" & vbCrLf & _
           "• Merge & Center Cells with Labels enabled" & vbCrLf & _
           "• Title row merged and centered", vbInformation
End Sub


