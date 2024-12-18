Attribute VB_Name = "ConvertC"
Sub ConvertExSong(sheetName As String)
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow1 As Long
    Dim startRow As Long
    Dim pvSlot As String
    Dim fileExtension As String

    ' Set the file extension
    fileExtension = ".ogg"

    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set ws1 = ThisWorkbook.Sheets("ExSongList")
    Set ws2 = ThisWorkbook.Sheets("Temp")

    ' Set the slot groups to be processed
    Dim includeGroups As Variant
    Dim excludeGroups As Variant
    includeGroups = Split(ThisWorkbook.Sheets("Temp").Cells(3, 6).value, "/") ' Split the value in I11 into an array
    excludeGroups = Split(ThisWorkbook.Sheets(sheetName).Cells(12, 9).value, "/") ' Split the value in I12 into an array

    ' Convert includeGroups and excludeGroups to integers
    Dim k As Integer
    For k = LBound(includeGroups) To UBound(includeGroups)
        includeGroups(k) = CInt(includeGroups(k))
    Next k

    For k = LBound(excludeGroups) To UBound(excludeGroups)
        excludeGroups(k) = CInt(excludeGroups(k))
    Next k

    ' Initialize: Clear columns C and D in ws2
    ws2.Range("C:D").ClearContents

    lastRow1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    startRow = 2

    Do While startRow <= lastRow1
        ' Pre-store the value of the cell in a variable and check its existence
        Dim cellValue As Variant
        cellValue = ws1.Cells(startRow, 2).value

        ' Check if the value is not empty
        If IsEmpty(cellValue) Then
            startRow = startRow + 1
            GoTo SkipProcessing
        End If

        ' Check if the value is numeric
        If IsNumeric(cellValue) Then
            pvSlot = CInt(cellValue)
        Else
            startRow = startRow + 1
            GoTo SkipProcessing
        End If

        ' Proceed only if the conditions for includeGroups and excludeGroups are met
        If IsGroupIncluded(CStr(pvSlot), includeGroups) And Not IsGroupIncluded(CStr(pvSlot), excludeGroups) Then
            ' Process the ex_song group
            Dim exSongCount As Integer, exAuthCount As Integer
            Dim exAuthExists As Boolean
            Dim songFile As String, chara As String, orgName As String, newName As String
            Dim tempFile As String
            Dim exSong As Integer, exAuth As Integer

            exSongCount = 0

            ' Process ex_song
            Do While startRow <= lastRow1 And CInt(ws1.Cells(startRow, 2).value) = pvSlot
                exSong = ws1.Cells(startRow, 3).value
                chara = ws1.Cells(startRow, 4).value
                songFile = ws1.Cells(startRow, 5).value
                exAuthCount = 0
                exAuthExists = False
                tempFile = ""

                ' Process the current ex_song
                Do While startRow <= lastRow1 And CInt(ws1.Cells(startRow, 2).value) = pvSlot And ws1.Cells(startRow, 3).value = exSong
                    If Not IsEmpty(ws1.Cells(startRow, 4).value) Then
                        ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Offset(1, 0).value = "pv_" & Format(pvSlot, "000") & ".ex_song." & exSong & ".chara=" & ws1.Cells(startRow, 4).value
                    End If
                    If Not IsEmpty(ws1.Cells(startRow, 5).value) Then
                        tempFile = "pv_" & Format(pvSlot, "000") & ".ex_song." & exSong & ".file=rom/sound/song/" & ws1.Cells(startRow, 5).value & fileExtension
                    End If

                    ' Process the ex_auth group
                    If Not IsEmpty(ws1.Cells(startRow, 6).value) And Not IsEmpty(ws1.Cells(startRow, 7).value) And Not IsEmpty(ws1.Cells(startRow, 8).value) Then
                        exAuth = ws1.Cells(startRow, 6).value
                        orgName = ws1.Cells(startRow, 7).value
                        newName = ws1.Cells(startRow, 8).value

                        ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Offset(1, 0).value = "pv_" & Format(pvSlot, "000") & ".ex_song." & exSong & ".ex_auth." & exAuth & ".name=" & newName
                        ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Offset(1, 0).value = "pv_" & Format(pvSlot, "000") & ".ex_song." & exSong & ".ex_auth." & exAuth & ".org_name=" & orgName
                        exAuthCount = exAuthCount + 1
                        exAuthExists = True
                    End If
                    startRow = startRow + 1
                Loop

                If exAuthExists Then
                    ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Offset(1, 0).value = "pv_" & Format(pvSlot, "000") & ".ex_song." & exSong & ".ex_auth.length=" & exAuthCount
                    ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Offset(1, 0).value = tempFile
                Else
                    ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Offset(1, 0).value = tempFile
                End If

                exSongCount = exSongCount + 1
            Loop

            ' Process ex_song.length
            If exSongCount > 0 Then
                ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Offset(1, 0).value = "pv_" & Format(pvSlot, "000") & ".ex_song.length=" & exSongCount
            End If
        End If

SkipProcessing:
        ' Advance startRow to the next to avoid infinite loop
        startRow = startRow + 1
        If startRow > lastRow1 Then Exit Do
    Loop

    ' Clear column D in ws2
    ws2.Range("D:D").ClearContents

    Exit Sub

    ' Re-enable screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Function IsGroupIncluded(group As String, groups As Variant) As Boolean
    Dim i As Integer
    For i = LBound(groups) To UBound(groups)
        If group = CStr(groups(i)) Then
            IsGroupIncluded = True
            Exit Function
        End If
    Next i
    IsGroupIncluded = False
End Function

