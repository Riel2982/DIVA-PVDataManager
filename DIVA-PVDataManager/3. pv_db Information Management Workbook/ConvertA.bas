Attribute VB_Name = "ConvertA"
Sub ConvertAnotherSong(sheetName As String)
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim lastRow As Long
    Dim ws2Row As Long
    Dim i As Long
    Dim startRow As Long
    Dim currentGroup As String
    Dim previousGroup As String
    Dim rowOffset As Integer
    Dim lastAnotherSong As Long
    Dim fileNameOnly As String
    Dim fileExtension As String
    
    ' Set the file extension
    fileExtension = ".ogg"

    ' Set the worksheets
    Set ws1 = Worksheets("AnotherSongList")
    Set ws2 = Worksheets("Temp")

    ' Set the slot groups to be processed
    Dim includeGroups As Variant
    Dim excludeGroups As Variant
    includeGroups = Split(ThisWorkbook.Sheets(sheetName).Cells(11, 9).value, "/") ' Split the value in I11 into an array
    excludeGroups = Split(ThisWorkbook.Sheets(sheetName).Cells(12, 9).value, "/") ' Split the value in I12 into an array

    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManua

    ' Convert includeGroups and excludeGroups to integers
    Dim k As Integer
    For k = LBound(includeGroups) To UBound(includeGroups)
        includeGroups(k) = CInt(includeGroups(k))
    Next k

    For k = LBound(excludeGroups) To UBound(excludeGroups)
        excludeGroups(k) = CInt(excludeGroups(k))
    Next k

    ' Initialize startRow
    startRow = 2
    
    ' Clear column A in ws2
    ws2.Columns("A").Clear
    
    ' Set the first row in ws2
    ws2Row = 1
    
    ' Get the last row in ws1
    lastRow = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row

    ' Copy data from ws1 to ws2 by group (if column A is empty)
    For i = startRow To lastRow
        If ws1.Cells(i, 1).value = "" Then ' If column A is empty
            currentGroup = "pv_" & Format(ws1.Cells(i, 2).value, "000")
            
            ' Check the slot group conditions
            If IsGroupIncluded(currentGroup, includeGroups) And Not IsGroupIncluded(currentGroup, excludeGroups) Then
                ' When a new group starts
                If currentGroup <> previousGroup Then
                    If previousGroup <> "" Then
                        ' Write the length of the previousGroup's last another_song + 1
                        ws2.Cells(ws2Row, 1).value = previousGroup & ".another_song.length=" & lastAnotherSong + 1
                        ws2Row = ws2Row + 1 ' Skip 1 row
                    End If
                    previousGroup = currentGroup
                    lastAnotherSong = 0 ' Reset for the new group
                End If
                
                rowOffset = 0

                ' Data writing process
                If ws1.Cells(i, 4).value <> "" Then
                    ws2.Cells(ws2Row + rowOffset, 1).value = "pv_" & ws1.Cells(i, 2).value & ".another_song." & ws1.Cells(i, 3).value & ".name=" & ws1.Cells(i, 4).value
                    lastAnotherSong = ws1.Cells(i, 3).value ' Update the value of the last another_song
                    rowOffset = rowOffset + 1
                End If
                If ws1.Cells(i, 5).value <> "" Then
                    ws2.Cells(ws2Row + rowOffset, 1).value = "pv_" & ws1.Cells(i, 2).value & ".another_song." & ws1.Cells(i, 3).value & ".name_en=" & ws1.Cells(i, 5).value
                    rowOffset = rowOffset + 1
                End If
                If ws1.Cells(i, 6).value <> "" Then
                    ' Fix the processing of song_file_name to append the extension only to the file name
                    fileNameOnly = ws1.Cells(i, 6).value
                    ws2.Cells(ws2Row + rowOffset, 1).value = "pv_" & ws1.Cells(i, 2).value & ".another_song." & ws1.Cells(i, 3).value & ".song_file_name=rom/sound/song/" & fileNameOnly & fileExtension
                    rowOffset = rowOffset + 1
                End If
                If ws1.Cells(i, 7).value <> "" Then
                    ws2.Cells(ws2Row + rowOffset, 1).value = "pv_" & ws1.Cells(i, 2).value & ".another_song." & ws1.Cells(i, 3).value & ".vocal_disp_name=" & ws1.Cells(i, 7).value
                    rowOffset = rowOffset + 1
                End If
                If ws1.Cells(i, 8).value <> "" Then
                    ws2.Cells(ws2Row + rowOffset, 1).value = "pv_" & ws1.Cells(i, 2).value & ".another_song." & ws1.Cells(i, 3).value & ".vocal_disp_name_en=" & ws1.Cells(i, 8).value
                    rowOffset = rowOffset + 1
                End If

                ' Update ws2Row by the number of rows written
                If rowOffset > 0 Then
                    ws2Row = ws2Row + rowOffset
                End If
            End If
        End If
    Next i
    
    ' Length processing for the last group
    If previousGroup <> "" Then
        ws2.Cells(ws2Row, 1).value = previousGroup & ".another_song.length=" & lastAnotherSong + 1
    End If

    ' Re-enable screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Function IsGroupIncluded(group As String, groups As Variant) As Boolean
    Dim i As Integer
    For i = LBound(groups) To UBound(groups)
        If group = "pv_" & Format(groups(i), "000") Then
            IsGroupIncluded = True
            Exit Function
        End If
    Next i
    IsGroupIncluded = False
End Function

