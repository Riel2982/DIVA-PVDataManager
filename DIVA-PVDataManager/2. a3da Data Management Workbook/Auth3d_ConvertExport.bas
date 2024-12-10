Attribute VB_Name = "Module1"
' Automatically adjust table range
Public Sub AdjustTableRange(sheetName As String, tableIndex As Integer)
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim lastCol As Long
    Dim ws As Worksheet
    Dim tableName As String
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Dynamically generate table name
    tableName = "EditAuth3dList" & tableIndex
    
    ' Specify the name of the table
    Set tbl = ws.ListObjects(tableName)
    
    ' Get the last row and last column
    lastRow = tbl.Range.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = tbl.Range.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    ' Adjust the table range
    tbl.Resize tbl.Range.Resize(lastRow - tbl.HeaderRowRange.Row + 1, lastCol - tbl.Range.Column + 1)
End Sub

' Convert to mod_auth_3d_db format
Public Sub SortAuth3d(sheetName As String)
    Dim ws As Worksheet, ws2 As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim categories As Collection
    Dim maxUid As Long
    Dim data As Variant
    Dim tempArray() As String
    Dim currentRow As Long

    Set ws = ThisWorkbook.Sheets(sheetName)
    Set ws2 = ThisWorkbook.Sheets("Preview")
    Set categories = New Collection

    ' Clear Preview
    ws2.Cells.Clear
    ws2.Cells(1, 1).Value = "#A3DA__________"
    ws2.Cells(2, 1).Value = "# date time was eliminated."

    ' Get unique categories
    On Error Resume Next
    For i = 2 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        categories.Add ws.Cells(i, "B").Value, CStr(ws.Cells(i, "B").Value)
    Next i
    On Error GoTo 0

    ' Write categories to auth_3d_db starting from row 3
    currentRow = 3
    For i = 1 To categories.Count
        ws2.Cells(currentRow, 1).Value = "category." & (i - 1) & ".value=" & categories(i)
        currentRow = currentRow + 1
    Next i
    ws2.Cells(currentRow, 1).Value = "category.length=" & categories.Count
    currentRow = currentRow + 1

    ' Write UID to auth_3d_db from the current row
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    j = 0
    maxUid = 0
    For i = 2 To lastRow
        ' Skip rows with "š"
        If ws.Cells(i, "A").Value <> "š" Then
            ' Add error check
            If IsEmpty(ws.Cells(i, "C").Value) Then
                Exit Sub
            End If

            ws2.Cells(currentRow, 1).Value = "uid." & j & ".category=" & ws.Cells(i, "B").Value
            ws2.Cells(currentRow + 1, 1).Value = "uid." & j & ".org_uid=" & ws.Cells(i, "C").Value
            ws2.Cells(currentRow + 2, 1).Value = "uid." & j & ".size=" & ws.Cells(i, "D").Value
            ws2.Cells(currentRow + 3, 1).Value = "uid." & j & ".value=A " & ws.Cells(i, "E").Value
            currentRow = currentRow + 4
            j = j + 1
            If ws.Cells(i, "C").Value > maxUid Then
                maxUid = ws.Cells(i, "C").Value
            End If
        End If
    Next i
    ws2.Cells(currentRow, 1).Value = "uid.length=" & j
    ws2.Cells(currentRow + 1, 1).Value = "uid.max=" & maxUid

    ' Sort data alphabetically starting from row 3
    lastRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    data = ws2.Range("A3:A" & lastRow).Value
    ReDim tempArray(1 To UBound(data, 1))

    For i = 1 To UBound(data, 1)
        tempArray(i) = data(i, 1)
    Next i

    Call QuickSort(tempArray, LBound(tempArray), UBound(tempArray))

    For i = 1 To UBound(tempArray)
        ws2.Cells(i + 2, 1).Value = tempArray(i)
    Next i
End Sub

' Export as mod_3d_db.bin
Public Sub ExportAuth3dDataBaseBin(sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim cellContent As String
    Dim filePath As String
    Dim fileDialog As fileDialog
    Dim fileNumber As Integer
    Dim folderPath As String
    Dim fileName As String
    Dim newFileName As String
    Dim currentDateTime As String
    Dim byteArray() As Byte
    Dim cellValue As String
    
    ' Set the target sheet
    Set ws = ThisWorkbook.Worksheets("Preview")

    ' Display save dialog (folder selection only)
    Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
    With fileDialog
        .Title = "Please select a folder to save"
        .Show
        If .SelectedItems.Count > 0 Then
            folderPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' Set fixed file name
    fileName = "mod_auth_3d_db.bin"
    filePath = folderPath & "\" & fileName
    
    ' Rename if the same file name exists
    If Dir(filePath) <> "" Then
        currentDateTime = Format(Now, "yyyy-mm-dd_hh-nn-ss")
        newFileName = "mod_auth_3d_db_" & currentDateTime & ".bin"
        Name filePath As folderPath & "\" & newFileName
    End If
    
    ' Open the file in text mode
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    
    ' Loop through all cells in column A
    Dim cell As Range
    For Each cell In ws.UsedRange.Columns(1).Cells
        cellValue = cell.Value
        Print #fileNumber, cellValue
    Next cell
    Close #fileNumber ' Close the file
End Sub

Public Sub QuickSort(arr() As String, first As Long, last As Long)
    Dim low As Long, high As Long, mid As String, temp As String
    low = first
    high = last
    mid = arr((first + last) \ 2)

    Do While (low <= high)
        Do While (arr(low) < mid)
            low = low + 1
        Loop
        Do While (arr(high) > mid)
            high = high - 1
        Loop
        If (low <= high) Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If (first < high) Then QuickSort arr, first, high
    If (low < last) Then QuickSort arr, low, last
End Sub

' Procedure for direct macro registration
Public Sub RunPreview1()
    PreviewRoutine ActiveSheet.Name, 1
End Sub

Public Sub RunExport1()
    ExportRoutine ActiveSheet.Name, 1
End Sub

Public Sub RunPreview2()
    PreviewRoutine ActiveSheet.Name, 2
End Sub

Public Sub RunExport2()
    ExportRoutine ActiveSheet.Name, 2
End Sub

' Procedure called within the sheet module
Public Sub PreviewRoutine(sheetName As String, tableIndex As Integer)
    AdjustTableRange sheetName, tableIndex
    SortAuth3d sheetName
End Sub

Public Sub ExportRoutine(sheetName As String, tableIndex As Integer)
    AdjustTableRange sheetName, tableIndex
    SortAuth3d sheetName
    ExportAuth3dDataBaseBin sheetName
End Sub
