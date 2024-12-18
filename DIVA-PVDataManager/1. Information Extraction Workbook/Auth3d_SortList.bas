Attribute VB_Name = "Auth3d_SortList"
Sub Auth3dSortList(sheetName As String)
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim i As Long, lastRow As Long, destRow As Long
    Dim tempCategory As String, tempOrgUid As String, tempSize As Long, tempValue As String
    Dim dataArray As Variant
    Dim filteredData() As Variant
    Dim filteredIndex As Long
    
    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Set sheets
    Set ws1 = ThisWorkbook.Sheets(sheetName)
    Set ws2 = ThisWorkbook.Sheets("ConvertAuth3dList")
    Set ws3 = ThisWorkbook.Sheets("Temp")

    ' Clear all rows in ws2 and set headers
    ws2.Cells.Clear
    ws2.Cells(1, 2).Value = "Category"
    ws2.Cells(1, 3).Value = "org_uid"
    ws2.Cells(1, 4).Value = "size"
    ws2.Cells(1, 5).Value = "a3da_Name"
    ' Clear ws3
    ws3.Cells.Clear

    ' Load data from ws1 into an array
    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row

    ' Load data from ws1 into an array
    If lastRow > 1 Then
        dataArray = ws1.Range("A2:A" & lastRow).Value ' Get as a 2D array
    Else
        MsgBox "No data in range"
        Exit Sub
    End If

    ' Initialize array for filtering
    ReDim filteredData(1 To lastRow - 1) ' A1 is not included, so -1
    filteredIndex = 0

    ' Filtering
    For i = 1 To UBound(dataArray, 1)
        If left(dataArray(i, 1), 4) = "uid." Then
            filteredIndex = filteredIndex + 1
            filteredData(filteredIndex) = dataArray(i, 1)
        End If
    Next i

    ' Adjust the size of the filtered array
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
    If lastRow - 1 <= UBound(dataArray, 1) Then
        ws2.Cells(destRow + 1, 2).Value = "uid.max"
        ws2.Cells(destRow + 1, 3).Value = Split(dataArray(lastRow - 1, 1), "=")(1)
    Else
        MsgBox "Data array range issue"
        Exit Sub
    End If

    ' Clear the content of "Temp"
    ws3.Cells.Clear
    
    ' Re-enable screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    
    ' Activate the ConvertAuth3dList
    ws2.Activate

End Sub


Sub BubbleSort(arr As Variant)
    Dim i As Long, j As Long
    Dim Temp As Variant
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
                    Temp = arr(i)
                    arr(i) = arr(j)
                    arr(j) = Temp
                End If
                success = True
            End If
            On Error GoTo 0
            If Not success Then Exit For
        Next j
    Next i
End Sub

Sub CopyToClipboard(sheetName As String)
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Get the last row with data (until an empty cell is found in column B)
    For i = 2 To ws.Rows.Count
        If IsEmpty(ws.Cells(i, "B")) Then
            lastRow = i - 1
            Exit For
        End If
    Next i
    
    ' If no empty cell is found in column B, set lastRow to the last row in the sheet
    If lastRow = 0 Then lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Determine the range to copy
    Dim rng As Range
    If WorksheetFunction.CountA(ws.Range("A2:A" & lastRow)) = 0 Then
        ' If there is no data in column A but there is data in column B
        If WorksheetFunction.CountA(ws.Range("B2:B" & lastRow)) > 0 Then
            ' Include F-H columns if any data is found in F-H columns
            If WorksheetFunction.CountA(ws.Range("F2:H" & lastRow)) > 0 Then
                Set rng = ws.Range("A2:H" & lastRow)
            Else
                Set rng = ws.Range("A2:E" & lastRow)
            End If
        End If
    Else
        ' If there is data in column A, copy from A2 to the first empty cell in column A
        For i = 2 To ws.Rows.Count
            If IsEmpty(ws.Cells(i, "A")) Then
                lastRow = i - 1
                Exit For
            End If
        Next i
        Set rng = ws.Range("A2:A" & lastRow)
    End If
    
    ' Copy to clipboard
    rng.Copy
End Sub


' Procedure to register directly as macro

Public Sub RunConvert()
    Auth3dSortList activeSheet.name
End Sub

Public Sub RunCopy()
    CopyToClipboard activeSheet.name
End Sub
