Attribute VB_Name = "Auth3d_OptExport"
Sub OptExportAuth3dDB(sheetName As String, Optional isExport As Boolean = False)
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim lastRow As Long, destRow As Long, i As Long, uidCounter As Long
    Dim categoryDict As Object, categoryIndex As Long, currentCategory As String
    Dim maxOrgUid As Long, currentOrgUid As Long, data As Variant, tempArray() As String
    Dim uidDict As Object, uidFound As Boolean
    
    Set ws1 = ThisWorkbook.Sheets(sheetName)
    Set ws2 = ThisWorkbook.Sheets("Temp")
    Set ws3 = ThisWorkbook.Sheets("OptAuth3dDB")
    
    ' Clear Temp sheet and OptAuth3dDB sheet before starting
    ws2.Cells.Clear
    ws3.Cells.Clear

    ' Remove highlight from EditAuth3dDataBase before starting the process
    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow
        ws1.Cells(i, 1).Interior.ColorIndex = xlNone
    Next i

    ' Copy data to Temp sheet (only rows with "uid.number.")
    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    destRow = 1
    For i = 1 To lastRow
        If left(ws1.Cells(i, 1).Value, 4) = "uid." Then
            Dim uidPart As String
            uidPart = Split(ws1.Cells(i, 1).Value, ".")(1)
            If IsNumeric(uidPart) Then
                ws2.Cells(destRow, 1).Value = ws1.Cells(i, 1).Value
                destRow = destRow + 1
            End If
        End If
    Next i

    ' Create a dictionary to keep count of uids
    Set uidDict = CreateObject("Scripting.Dictionary")
    uidFound = False
    For i = 1 To lastRow
        If left(ws1.Cells(i, 1).Value, 4) = "uid." Then
            Dim uid As String
            uid = Split(ws1.Cells(i, 1).Value, ".")(1)
            ' Exclude uid.length and uid.max
            If IsNumeric(uid) Then
                If uidDict.Exists(uid) Then
                    uidDict(uid) = uidDict(uid) + 1
                Else
                    uidDict(uid) = 1
                End If
            End If
        End If
    Next i

    ' Highlight uids that appear exactly 1, 2, or 3 times
    For i = 1 To lastRow
        If left(ws1.Cells(i, 1).Value, 4) = "uid." Then
            uid = Split(ws1.Cells(i, 1).Value, ".")(1)
            If IsNumeric(uid) Then
                If uidDict(uid) = 1 Or uidDict(uid) = 2 Or uidDict(uid) = 3 Then
                    ws1.Cells(i, 1).Interior.Color = RGB(255, 0, 0)
                    uidFound = True
                End If
            End If
        End If
    Next i

    ' Organize data in Temp sheet
    lastRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    uidCounter = 0
    maxOrgUid = 0
    For i = 1 To lastRow Step 4
        If i + 3 <= lastRow Then
            ws2.Cells(i, 1).Value = "uid." & uidCounter & ".category=" & Split(ws2.Cells(i, 1).Value, "=")(1)
            ws2.Cells(i + 1, 1).Value = "uid." & uidCounter & ".org_uid=" & Split(ws2.Cells(i + 1, 1).Value, "=")(1)
            ws2.Cells(i + 2, 1).Value = "uid." & uidCounter & ".size=" & Split(ws2.Cells(i + 2, 1).Value, "=")(1)
            ws2.Cells(i + 3, 1).Value = "uid." & uidCounter & ".value=" & Split(ws2.Cells(i + 3, 1).Value, "=")(1)
            uidCounter = uidCounter + 1
            
            ' Get the maximum value of org_uid
            currentOrgUid = Val(Split(ws2.Cells(i + 1, 1).Value, "=")(1))
            If currentOrgUid > maxOrgUid Then
                maxOrgUid = currentOrgUid
            End If
        End If
    Next i

    ' Organize data in OptAuth3dDB
    ws3.Cells(1, 1).Value = "#A3DA__________"
    ws3.Cells(2, 1).Value = "# date time was eliminated."
    
    Set categoryDict = CreateObject("Scripting.Dictionary")
    destRow = 3
    lastRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRow Step 4
        currentCategory = Split(ws2.Cells(i, 1).Value, "=")(1)
        If Not categoryDict.Exists(currentCategory) Then
            categoryDict.Add currentCategory, categoryDict.Count
            ws3.Cells(destRow, 1).Value = "category." & categoryDict(currentCategory) & ".value=" & currentCategory
            destRow = destRow + 1
        End If
    Next i
    ws3.Cells(destRow, 1).Value = "category.length=" & categoryDict.Count
    destRow = destRow + 1
    
    For i = 1 To lastRow
        ws3.Cells(destRow, 1).Value = ws2.Cells(i, 1).Value
        destRow = destRow + 1
    Next i

    ' Sort data in alphabetical order starting from row 3
    lastRow = ws3.Cells(ws3.Rows.Count, 1).End(xlUp).Row
    data = ws3.Range("A3:A" & lastRow).Value
    ReDim tempArray(1 To UBound(data, 1))
    For i = 1 To UBound(data, 1)
        tempArray(i) = data(i, 1)
    Next i
    Call QuickSort(tempArray, LBound(tempArray), UBound(tempArray))
    For i = 1 To UBound(tempArray, 1)
        ws3.Cells(i + 2, 1).Value = tempArray(i)
    Next i
    
    ws3.Cells(lastRow + 1, 1).Value = "uid.length=" & uidCounter
    ws3.Cells(lastRow + 2, 1).Value = "uid.max=" & maxOrgUid

    ' Activate OptAuth3dDB sheet if no uids found with 1, 2, or 3 occurrences (only when optimizing)
    If Not uidFound And Not isExport Then
        ws3.Activate
    End If

    
    ' Export to binary file if needed
    If isExport Then
        ExportAuth3dDataBaseBin
    End If
End Sub

Sub QuickSort(arr As Variant, left As Long, right As Long)
    Dim i As Long, j As Long, pivot As String
    i = left
    j = right
    pivot = arr((left + right) \ 2)
    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            Swap arr, i, j
            i = i + 1
            j = j - 1
        End If
    Loop
    If left < j Then QuickSort arr, left, j
    If i < right Then QuickSort arr, i, right
End Sub

Sub Swap(arr As Variant, a As Long, b As Long)
    Dim temp As String
    temp = arr(a)
    arr(a) = arr(b)
    arr(b) = temp
End Sub


Public Sub RunOptimize()
    OptExportAuth3dDB ActiveSheet.Name
End Sub

Public Sub RunExport()
    OptExportAuth3dDB ActiveSheet.Name, True
End Sub

' Output as mod_3d_db.bin
Public Sub ExportAuth3dDataBaseBin()
    Dim ws As Worksheet, fileDialog As fileDialog, fileNumber As Integer
    Dim folderPath As String, fileName As String, filePath As String
    Dim newFileName As String, currentDateTime As String, cellValue As String
    
    ' Set the target sheet
    Set ws = ThisWorkbook.Worksheets("OptAuth3dDB")
    
    ' Show the save dialog (folder selection only)
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
    
    ' Set the fixed file name
    fileName = "mod_auth_3d_db.bin"
    filePath = folderPath & "\" & fileName
    
    ' If a file with the same name exists, rename it
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
