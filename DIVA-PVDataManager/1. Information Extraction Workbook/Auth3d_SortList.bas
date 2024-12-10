Attribute VB_Name = "Module1"
Sub Auth3dSortList()
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim i As Long, lastRow As Long, destRow As Long
    Dim tempCategory As String, tempOrgUid As String, tempSize As Long, tempValue As String
    Dim dataArray As Variant
    Dim filteredData() As Variant
    Dim filteredIndex As Long

    ' Set sheets
    Set ws1 = ThisWorkbook.Sheets("mod_auth_3d_db")
    Set ws2 = ThisWorkbook.Sheets("ConvertAuth3dList")
    Set ws3 = ThisWorkbook.Sheets("Temp")

    ' Clear all rows except the first row in ws2
    If ws2.UsedRange.Rows.Count > 1 Then
        ws2.Rows("2:" & ws2.Rows.Count).Clear
    End If
    ' Clear ws3
    ws3.Cells.Clear

    ' Load data from ws1 into an array
    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row

    ' Load data from ws1 into an array
    If lastRow > 1 Then
        dataArray = ws1.Range("A1:A" & lastRow).Value ' Get as a 2D array
    Else
        MsgBox "No data in range"
        Exit Sub
    End If

    ' Initialize array for filtering
    ReDim filteredData(1 To lastRow)
    filteredIndex = 0

    ' Filtering
    For i = 1 To UBound(dataArray, 1)
        If left(dataArray(i, 1), 4) = "uid." Then
            filteredIndex = filteredIndex + 1
            filteredData(filteredIndex) = dataArray(i, 1)
        End If
    Next i

    ' Adjust the size of the filtered arra
    ReDim Preserve filteredData(1 To filteredIndex)

    ' Sort by numbers
    Call BubbleSort(filteredData)

    ' Write sorted data to ws3
    For i = 1 To UBound(filteredData)
        ws3.Cells(i, 1).Value = filteredData(i)
    Next i

    ' Write data from ws3 to ws2
    destRow = 2
    For i = 1 To UBound(filteredData) Step 4
        If i + 3 <= UBound(filteredData) Then
            tempCategory = Split(filteredData(i), "=")(1)
            tempOrgUid = Split(filteredData(i + 1), "=")(1)
            tempSize = Split(filteredData(i + 2), "=")(1)
            tempValue = Split(filteredData(i + 3), "=")(1)

            ws2.Cells(destRow, 2).Value = tempCategory
            ws2.Cells(destRow, 3).Value = tempOrgUid
            ws2.Cells(destRow, 4).Value = tempSize
            ws2.Cells(destRow, 5).Value = Trim(Mid(tempValue, 3))

            destRow = destRow + 1
        End If
    Next i

    ' Write uid.max one cell below the last
    ws2.Cells(destRow + 1, 2).Value = "uid.max"
    ws2.Cells(destRow + 1, 3).Value = Split(dataArray(lastRow, 1), "=")(1)

    ' Clear the content of "Empty"
    ws3.Cells.Clear
End Sub

Sub BubbleSort(arr As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim index1 As Long, index2 As Long
    Dim success As Boolean

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            success = False
            On Error Resume Next
            index1 = CLng(Split(Split(arr(i), "=")(0), ".")(1))
            index2 = CLng(Split(Split(arr(j), "=")(0), ".")(1))
            If Err.Number = 0 Then
                If index1 > index2 Then
                    temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = temp
                End If
                success = True
            End If
            On Error GoTo 0
            If Not success Then Exit For
        Next j
    Next i
End Sub

' Workbook activated
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If Sh.Name = "ConvertAuth3dList" Then
        Call Auth3dSortList
    End If
End Sub



Sub CopyToClipboard(sheetName As String)
    Dim lastRow As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Get the last row with data (until an empty cell is found in columns B to E)
    For i = 2 To ws.Rows.Count
        If IsEmpty(ws.Cells(i, "B")) Or IsEmpty(ws.Cells(i, "C")) Or IsEmpty(ws.Cells(i, "D")) Or IsEmpty(ws.Cells(i, "E")) Then
            lastRow = i - 1
            Exit For
        End If
    Next i
    
    ' Select the range from A1 to the last row with data in column E
    Dim rng As Range
    Set rng = ws.Range("A2:E" & lastRow)
    
    ' Copy to clipboard
    rng.Copy
End Sub

' Procedure to register directly as macro
Public Sub RunCopy()
    CopyToClipboard ActiveSheet.Name
End Sub

