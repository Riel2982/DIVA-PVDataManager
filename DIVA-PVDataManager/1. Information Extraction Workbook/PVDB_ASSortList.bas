Attribute VB_Name = "PVDB_ASSortList"
Sub ASSortList(sheetName As String)
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim ws4 As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim slotNum As String
    Dim songCount As Long
    Dim destRow As Long
    Dim currentRow As Long
    Dim anotherSongIndex As Long

    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Set sheets
    Set ws1 = ThisWorkbook.Sheets(sheetName)
    Set ws2 = ThisWorkbook.Sheets("Temp")
    Set ws3 = ThisWorkbook.Sheets("AnotherSongList")
    Set ws4 = ThisWorkbook.Sheets("ExtractPVDB")

    ' Clear sheets
    ws2.Cells.Clear
    ws3.Cells.Clear
    ws4.Rows("2:" & ws4.Rows.Count).Clear

    ' Add header row
    ws3.Cells(1, 2).Value = "pv_slot"
    ws3.Cells(1, 3).Value = "another_song"
    ws3.Cells(1, 4).Value = "SongDispName"
    ws3.Cells(1, 5).Value = "SongEngDispName"
    ws3.Cells(1, 6).Value = "Songfile"
    ws3.Cells(1, 7).Value = "Vocal"
    ws3.Cells(1, 8).Value = "EngVocal"

    ' Get the last row in the pv_db sheet
    lastRow = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row and separate based on the presence of another_song
    For i = 1 To lastRow
        If InStr(ws1.Cells(i, 1).Value, "another_song") > 0 Then
            destRow = ws2.Cells(ws2.Rows.Count, 3).End(xlUp).Row + 1
            ws2.Cells(destRow, 3).Value = ws1.Cells(i, 1).Value
        Else
            destRow = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).Row + 1
            ws2.Cells(destRow, 2).Value = ws1.Cells(i, 1).Value
        End If
    Next i

    ' Loop through each slot number
    lastRow = ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Row
    Dim extractedSlots As Object
    Set extractedSlots = CreateObject("Scripting.Dictionary")

    For i = 1 To lastRow
        ' Get the slot number
        slotNum = GetSlotNumber(ws2.Cells(i, 3).Value)

        ' Skip already extracted slot numbers
        If extractedSlots.exists(slotNum) Then GoTo NextSlot

        ' Get the length of another_song
        songCount = GetAnotherSongLength(ws1, slotNum)

        ' If another_song does not exist, paste data in column B of ws2
        If songCount = 0 Then
            destRow = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).Row + 1
            ws2.Cells(destRow, 2).Value = ws2.Cells(i, 3).Value
        Else
            ' Paste another_song in column A
            For j = 0 To songCount - 1
                destRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row + 1
                ws2.Cells(destRow, 1).Value = "pv_" & slotNum & ".another_song." & j & ".name=" & GetValue(ws1, "pv_" & slotNum & ".another_song." & j & ".name")
                ws2.Cells(destRow + 1, 1).Value = "pv_" & slotNum & ".another_song." & j & ".name_en=" & GetValue(ws1, "pv_" & slotNum & ".another_song." & j & ".name_en")
                ws2.Cells(destRow + 2, 1).Value = "pv_" & slotNum & ".another_song." & j & ".vocal_disp_name=" & GetValue(ws1, "pv_" & slotNum & ".another_song." & j & ".vocal_disp_name")
                ws2.Cells(destRow + 3, 1).Value = "pv_" & slotNum & ".another_song." & j & ".vocal_disp_name_en=" & GetValue(ws1, "pv_" & slotNum & ".another_song." & j & ".vocal_disp_name_en")
                ws2.Cells(destRow + 4, 1).Value = "pv_" & slotNum & ".another_song." & j & ".song_file_name=" & GetValue(ws1, "pv_" & slotNum & ".another_song." & j & ".song_file_name")
            Next j

            ' Paste another_song.length in column A
            destRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row + 1
            ws2.Cells(destRow, 1).Value = "pv_" & slotNum & ".another_song.length=" & songCount

            ' Paste other data in column B of ws2
            destRow = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).Row + 1
            ws2.Cells(destRow, 2).Value = ws2.Cells(i, 3).Value
        End If

        ' Add slot number to dictionary
        extractedSlots.Add slotNum, True

NextSlot:
    Next i

    ' Set currentRow based on the presence of a header row in ws3
    If ws3.Cells(1, 2).Value = "" And ws3.Cells(1, 3).Value = "" Then
        currentRow = 2 ' No header row
    Else
        currentRow = 2 ' Header row present
    End If

    ' Get the last row in the Temp sheet
    lastRow = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row

    ' Copy data
    For i = 1 To lastRow
        Dim cellValue As String
        cellValue = ws2.Cells(i, 1).Value

        ' Extract slot number and another_song number
        If InStr(cellValue, ".name=") > 0 Then
            slotNum = Split(Split(cellValue, "pv_")(1), ".")(0)
            anotherSongIndex = Split(Split(cellValue, ".another_song.")(1), ".")(0)
            ws3.Cells(currentRow, 2).Value = slotNum
            ws3.Cells(currentRow, 3).Value = anotherSongIndex
            ws3.Cells(currentRow, 4).Value = Split(cellValue, "=")(1)
        ' Process English name
        ElseIf InStr(cellValue, ".name_en=") > 0 Then
            ws3.Cells(currentRow, 5).Value = Split(cellValue, "=")(1)
        ' Process file name
        ElseIf InStr(cellValue, ".song_file_name=") > 0 Then
            Dim fileNameParts() As String
           fileNameParts = Split(cellValue, "song/")
           If UBound(fileNameParts) > 0 Then
                Dim fileName As String
                fileName = fileNameParts(1)
             ' Remove the extension
             fileName = Replace(fileName, ".ogg", "")
                ws3.Cells(currentRow, 6).Value = fileName
            Else
             ws3.Cells(currentRow, 6).Value = ""
            End If
        ' Process vocal display name
        ElseIf InStr(cellValue, ".vocal_disp_name=") > 0 Then
            ws3.Cells(currentRow, 7).Value = Split(cellValue, "=")(1)
        ' Process English vocal display name
        ElseIf InStr(cellValue, ".vocal_disp_name_en=") > 0 Then
            ws3.Cells(currentRow, 8).Value = Split(cellValue, "=")(1)
        End If

        ' Process another_song.length
        If InStr(cellValue, "length=") > 0 Then
            slotNum = Split(Split(cellValue, "pv_")(1), ".")(0)
            anotherSongLength = Split(Split(cellValue, ".length=")(1), ".")(0)
            ws3.Cells(currentRow, 2).Value = slotNum
            ws3.Cells(currentRow, 3).Value = anotherSongLength
        End If

        ' Move to the next row for each another_song
        If InStr(cellValue, ".name=") > 0 Or InStr(cellValue, ".another_song.length=") > 0 Then
            currentRow = currentRow + 1
        End If
    Next i

    ' Clear cells E2:H2 and move the cells below them up
    ws3.Range("E2:H2").ClearContents

    ' Move (cut) data to fill the gap
    ws3.Range("E3:H" & ws3.Cells(ws3.Rows.Count, "E").End(xlUp).Row).Cut ws3.Range("E2")

    ' Copy data from column B of Temp sheet to ws4
    lastRow = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
     currentRow = 1
    For i = 1 To lastRow
        Dim tempCellValue As String
        tempCellValue = ws2.Cells(i, 2).Value
        If tempCellValue <> "" And left(tempCellValue, 1) = "p" Then
            ws4.Cells(currentRow + 1, 1).Value = tempCellValue
            currentRow = currentRow + 1
        End If
    Next i

    ' Remove rows in ws4 that contain "another_song.0.name"
    lastRow = ws4.Cells(ws4.Rows.Count, "A").End(xlUp).Row
    For i = lastRow To 1 Step -1
    If InStr(ws4.Cells(i, 1).Value, "another_song.0.name") > 0 Then
        ws4.Rows(i).Delete
    End If
    Next i

    ' Additional code to remove rows where only columns B and C have data in ws3
    lastRow = ws3.Cells(ws3.Rows.Count, "B").End(xlUp).Row
    For i = lastRow To 2 Step -1 ' Start from the last row and go upwards, skipping the header
        If ws3.Cells(i, 2).Value <> "" And ws3.Cells(i, 3).Value <> "" _
        And ws3.Cells(i, 4).Value = "" And ws3.Cells(i, 5).Value = "" _
        And ws3.Cells(i, 6).Value = "" And ws3.Cells(i, 7).Value = "" _
        And ws3.Cells(i, 8).Value = "" Then
            ws3.Rows(i).Delete
        End If
    Next i

    ' Clear Temp sheet
    ws2.Cells.Clear
 
    ' Re-enable screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Set ws3 as active sheet
    ws3.Activate


End Sub
    


Function GetSlotNumber(cellValue As String) As String
    ' Function to extract slot number
    Dim parts() As String
    parts = Split(cellValue, ".")
    
    ' Check array size and if it is within bounds
    If UBound(parts) >= 1 Then
        parts = Split(parts(0), "_")
        If UBound(parts) >= 1 Then
            GetSlotNumber = parts(1)
        Else
            GetSlotNumber = "000" ' Set default value
        End If
    Else
        GetSlotNumber = "000" ' Set default value
    End If
End Function

Function GetAnotherSongLength(ws As Worksheet, slotNum As String) As Long
    ' Function to get another_song.length
    Dim i As Long
    Dim cellValue As String
    
    For i = 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        cellValue = ws.Cells(i, 1).Value
        If InStr(cellValue, "pv_" & slotNum & ".another_song.length") > 0 Then
            GetAnotherSongLength = CLng(Split(cellValue, "=")(1))
            Exit Function
        End If
    Next i
    GetAnotherSongLength = 0
End Function

Function GetValue(ws As Worksheet, searchString As String) As String
    ' Function to get the value of the cell containing the specified string
    Dim i As Long
    For i = 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If InStr(ws.Cells(i, 1).Value, searchString) > 0 Then
            If InStr(ws.Cells(i, 1).Value, "=") > 0 Then
                GetValue = Split(ws.Cells(i, 1).Value, "=")(1)
            Else
                GetValue = ""
            End If
            Exit Function
        End If
    Next i
    GetValue = ""
End Function


' Procedure to register directly as macro
Public Sub RunExtractA()
    ASSortList activeSheet.name
End Sub
