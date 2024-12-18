Attribute VB_Name = "PVDB_ExSSortList"
Sub ExSSortList(sheetName As String)
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim ws4 As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim slotNum As String
    Dim destRow As Long
    Dim exSongIndex As String
    Dim chara As String
    Dim songFile As String
    Dim exAuthIndex As String
    Dim fileRow As Long
    Dim tempLastRow As Long
    Dim bData As String
    Dim cData As String
    Dim name As String
    Dim orgName As String

    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Set sheets
    Set ws1 = ThisWorkbook.Sheets(sheetName)
    Set ws2 = ThisWorkbook.Sheets("Temp")
    Set ws3 = ThisWorkbook.Sheets("ExSongList")
    Set ws4 = ThisWorkbook.Sheets("ExtractPVDB")

    ' Clear sheets
    ws2.Cells.Clear
    ws3.Cells.Clear
    ws4.Rows("2:" & ws4.Rows.Count).Clear

    ' Add header row
    ws3.Cells(1, 2).Value = "pv_slot"
    ws3.Cells(1, 3).Value = "ex_song"
    ws3.Cells(1, 4).Value = "Character"
    ws3.Cells(1, 5).Value = "SongFile"
    ws3.Cells(1, 6).Value = "ex_auth"
    ws3.Cells(1, 7).Value = "org_name"
    ws3.Cells(1, 8).Value = "Replace"

    ' Get the last row in the data sheet
    lastRow = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    fileRow = 1 ' Initialize fileRow to keep track of Temp sheet rows

    ' Loop through each row and separate based on the presence of ex_song
    destRow = 2 ' Start writing data from the second row

    For i = 1 To lastRow
        If InStr(ws1.Cells(i, 1).Value, "ex_song") > 0 Then
            Dim parts() As String
            parts = Split(ws1.Cells(i, 1).Value, ".")
            slotNum = Mid(parts(0), 4) ' Remove "pv_" to get the slot number
            exSongIndex = parts(2) ' Get ex_song index

            ' Check if it is ex_song length line
            If InStr(ws1.Cells(i, 1).Value, "ex_song.length") > 0 Then GoTo NextRow

            ' Check if it is ex_auth length line
            If InStr(ws1.Cells(i, 1).Value, "ex_auth.length") > 0 Then GoTo NextRow

            ' Initialize variables
            chara = ""
            songFile = ""
            exAuthIndex = ""

            ' Check if it is ex_auth data
            If InStr(ws1.Cells(i, 1).Value, "ex_auth.") > 0 Then
                exAuthIndex = parts(4) ' Get ex_auth index

                If exAuthIndex = "0" Then
                    ' In case of ex_auth.0, place it in the character name row
                    If InStr(ws1.Cells(i, 1).Value, "org_name") > 0 Then
                        orgName = Split(ws1.Cells(i, 1).Value, "=")(1)
                        ws3.Cells(destRow - 1, 7).Value = orgName
                    ElseIf InStr(ws1.Cells(i, 1).Value, "name") > 0 Then
                        name = Split(ws1.Cells(i, 1).Value, "=")(1)
                        ws3.Cells(destRow - 1, 8).Value = name
                    End If
                    ws3.Cells(destRow - 1, 6).Value = exAuthIndex
                Else
                    ' For data of ex_auth.1 and beyond, place it in a new row as a single row
                    bData = slotNum ' Data for column B
                    cData = exSongIndex ' Data for column C

                    If InStr(ws1.Cells(i, 1).Value, "org_name") > 0 Then
                        orgName = Split(ws1.Cells(i, 1).Value, "=")(1)
                        ' Read the next row
                        If i < lastRow And InStr(ws1.Cells(i + 1, 1).Value, "ex_auth." & exAuthIndex & ".name") > 0 Then
                            name = Split(ws1.Cells(i + 1, 1).Value, "=")(1)
                            i = i + 1 ' Skip the already processed next row
                        End If
                    ElseIf InStr(ws1.Cells(i, 1).Value, "name") > 0 Then
                        name = Split(ws1.Cells(i, 1).Value, "=")(1)
                        ' Read the next row
                        If i < lastRow And InStr(ws1.Cells(i + 1, 1).Value, "ex_auth." & exAuthIndex & ".org_name") > 0 Then
                            orgName = Split(ws1.Cells(i + 1, 1).Value, "=")(1)
                            i = i + 1 ' Skip the already processed next row
                        End If
                    End If

                    ws3.Cells(destRow, 2).Value = bData ' Place slotNum in column B
                    ws3.Cells(destRow, 3).Value = cData ' Place exSongIndex in column C
                    ws3.Cells(destRow, 6).Value = exAuthIndex
                    ws3.Cells(destRow, 7).Value = orgName
                    ws3.Cells(destRow, 8).Value = name

                    destRow = destRow + 1
                End If
            Else
                ' Extract ex_song data
                If InStr(ws1.Cells(i, 1).Value, "chara") > 0 Then
                    chara = Split(ws1.Cells(i, 1).Value, "=")(1)
                End If

                If InStr(ws1.Cells(i, 1).Value, "file") > 0 Then
                    Dim filePath As String
                    filePath = Split(ws1.Cells(i, 1).Value, "=")(1)
                    Dim songFileParts() As String
                    songFileParts = Split(filePath, "/")
                    If UBound(songFileParts) >= 3 Then
                        ' Extract the core file name without path and extension
                        songFile = Replace(songFileParts(UBound(songFileParts)), ".ogg", "")
                        ws2.Cells(fileRow, 2).Value = songFile ' Temporarily store in Temp sheet B column
                        fileRow = fileRow + 1 ' Increment fileRow
                    End If
                End If

                ' Write to the worksheet if Character is present
                If chara <> "" Then
                    ws3.Cells(destRow, 2).Value = slotNum ' Place in column B
                    ws3.Cells(destRow, 3).Value = exSongIndex ' Place exSongIndex in column C
                    ws3.Cells(destRow, 4).Value = chara ' Character
                    destRow = destRow + 1
                End If
            End If
        Else
            ' Non-ex_song data goes to ws4
            If left(ws1.Cells(i, 1).Value, 1) <> "p" Then
                ' Skip the row to be deleted
                GoTo NextRow
            Else
                ws4.Cells(ws4.Cells(ws4.Rows.Count, "A").End(xlUp).Row + 1, 1).Value = ws1.Cells(i, 1).Value
            End If
        End If

NextRow:
    Next i

    ' Re-enable screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Get the last row in the Temp sheet
    tempLastRow = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row

    ' Reposition SongFile data from Temp to ExSongList
    Dim charaRow As Long
    fileRow = 1 ' Reset fileRow
    For charaRow = 2 To destRow - 1
        If ws3.Cells(charaRow, 4).Value <> "" Then
            If fileRow <= tempLastRow Then
                ws3.Cells(charaRow, 5).Value = ws2.Cells(fileRow, 2).Value ' Place SongFile data next to Character
                fileRow = fileRow + 1
            End If
        End If
    Next charaRow

    ' Set ws3 as active sheet
    ws3.Activate
End Sub

' Procedure to register directly as macro
Public Sub RunExtractEx()
    ExSSortList activeSheet.name
End Sub
