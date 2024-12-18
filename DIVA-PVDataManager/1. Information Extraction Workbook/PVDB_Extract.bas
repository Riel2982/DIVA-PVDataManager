Attribute VB_Name = "PVDB_Extract"
Sub ExtractA(sheetName As String)
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
                ws3.Cells(currentRow, 6).Value = fileNameParts(1)
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

    ' Copy data from column B of Temp sheet to column E of ws2
    lastRow = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    currentRow = 1
    For i = 1 To lastRow
        Dim tempCellValue As String
        tempCellValue = ws2.Cells(i, 2).Value
        If tempCellValue <> "" And left(tempCellValue, 1) = "p" Then
            ws2.Cells(currentRow + 1, 5).Value = tempCellValue ' Column E is the 5th column
            currentRow = currentRow + 1
        End If
    Next i

    ' Remove rows in ws2 that contain "another_song.0.name" in column E
        lastRow = ws2.Cells(ws2.Rows.Count, "E").End(xlUp).Row
        For i = lastRow To 1 Step -1
            If InStr(ws2.Cells(i, 5).Value, "another_song.0.name") > 0 Then
                ws2.Rows(i).Delete
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


    ' Re-enable screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' Set ws2 as active sheet
    ws2.Activate

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

Sub ExtractEx(sheetName As String)
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
    lastRow = ws2.Cells(ws2.Rows.Count, "E").End(xlUp).Row
    fileRow = 1 ' Initialize fileRow to keep track of Temp sheet rows

    ' Loop through each row and separate based on the presence of ex_song
    destRow = 2 ' Start writing data from the second row

    For i = 1 To lastRow
        If InStr(ws2.Cells(i, 5).Value, "ex_song") > 0 Then
            Dim parts() As String
            parts = Split(ws2.Cells(i, 5).Value, ".")
            slotNum = Mid(parts(0), 4) ' Remove "pv_" to get the slot number
            exSongIndex = parts(2) ' Get ex_song index

            ' Check if it is ex_song length line
            If InStr(ws2.Cells(i, 5).Value, "ex_song.length") > 0 Then GoTo NextRow

            ' Check if it is ex_auth length line
            If InStr(ws2.Cells(i, 5).Value, "ex_auth.length") > 0 Then GoTo NextRow

            ' Initialize variables
            chara = ""
            songFile = ""
            exAuthIndex = ""

            ' Check if it is ex_auth data
            If InStr(ws2.Cells(i, 5).Value, "ex_auth.") > 0 Then
                exAuthIndex = parts(4) ' Get ex_auth index

                If exAuthIndex = "0" Then
                    ' In case of ex_auth.0, place it in the character name row
                    If InStr(ws2.Cells(i, 5).Value, "org_name") > 0 Then
                        orgName = Split(ws2.Cells(i, 5).Value, "=")(1)
                        ws3.Cells(destRow - 1, 7).Value = orgName
                    ElseIf InStr(ws2.Cells(i, 5).Value, "name") > 0 Then
                        name = Split(ws2.Cells(i, 5).Value, "=")(1)
                        ws3.Cells(destRow - 1, 8).Value = name
                    End If
                    ws3.Cells(destRow - 1, 6).Value = exAuthIndex
                Else
                    ' For data of ex_auth.1 and beyond, place it in a new row as a single row
                    bData = slotNum ' Data for column B
                    cData = exSongIndex ' Data for column C

                    If InStr(ws2.Cells(i, 5).Value, "org_name") > 0 Then
                        orgName = Split(ws2.Cells(i, 5).Value, "=")(1)
                        ' Read the next row
                        If i < lastRow And InStr(ws2.Cells(i + 1, 5).Value, "ex_auth." & exAuthIndex & ".name") > 0 Then
                            name = Split(ws2.Cells(i + 1, 5).Value, "=")(1)
                            i = i + 1 ' Skip the already processed next row
                        End If
                    ElseIf InStr(ws2.Cells(i, 5).Value, "name") > 0 Then
                        name = Split(ws2.Cells(i, 5).Value, "=")(1)
                        ' Read the next row
                        If i < lastRow And InStr(ws2.Cells(i + 1, 5).Value, "ex_auth." & exAuthIndex & ".org_name") > 0 Then
                            orgName = Split(ws2.Cells(i + 1, 5).Value, "=")(1)
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
                If InStr(ws2.Cells(i, 5).Value, "chara") > 0 Then
                    chara = Split(ws2.Cells(i, 5).Value, "=")(1)
                End If

                If InStr(ws2.Cells(i, 5).Value, "file") > 0 Then
                    Dim filePath As String
                    filePath = Split(ws2.Cells(i, 5).Value, "=")(1)
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
                    ws3.Cells(destRow, 3).Value = exSongIndex ' ex_song index
                    ws3.Cells(destRow, 4).Value = chara ' Character
                    destRow = destRow + 1
                End If
            End If
        Else
            ' Non-ex_song data goes to ws4
            If left(ws2.Cells(i, 5).Value, 1) <> "p" Then
                ' Skip the row to be deleted
                GoTo NextRow
            Else
                ws4.Cells(ws4.Cells(ws4.Rows.Count, "A").End(xlUp).Row + 1, 1).Value = ws2.Cells(i, 5).Value
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

    ' Clear sheets
    ws2.Cells.Clear

    ' Set ws4 as active sheet
    ws4.Activate
End Sub


Public Sub ExportPVDBTxt(sheetName As String)
    Dim ws As Worksheet
    Dim folderPath As String, fileName As String, filePath As String
    Dim sakuraPath As String
    Dim wsh As Object

    ' Set the target sheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Get the path to the save location from cell A1
    folderPath = ws.Range("A1").Value
    If folderPath = "" Or Dir(folderPath, vbDirectory) = "" Then
        ' Show the save dialog (folder selection only) if A1 is empty or invalid
        With Application.fileDialog(msoFileDialogFolderPicker)
            .Title = "Please select a folder to save"
            .Show
            If .SelectedItems.Count > 0 Then
                folderPath = .SelectedItems(1)
            Else
                Exit Sub
            End If
        End With
    End If

    ' Get the path to Sakura Editor from cell E1
    sakuraPath = ws.Range("E1").Value
    If sakuraPath = "" Then
        MsgBox "Please enter the file path to Sakura Editor in cell E1.", vbExclamation
        Exit Sub
    End If

    ' Set the fixed file name
    fileName = "mod_pv_db.txt"
    filePath = folderPath & "\" & fileName

    ' If a file with the same name exists, rename it
    If Dir(filePath) <> "" Then
        Dim currentDateTime As String
        currentDateTime = Format(Now, "yyyy-mm-dd_hh-nn-ss")
        Dim newFileName As String
        newFileName = "mod_pv_db_" & currentDateTime & ".txt"
        Name filePath As folderPath & "\" & newFileName
    End If

    ' Create WshShell object
    Set wsh = CreateObject("WScript.Shell")

    ' Run Sakura Editor
    wsh.Run """" & sakuraPath & """"

    ' Wait for Sakura Editor to open
    Application.Wait Now + TimeValue("00:00:06")

    ' Try to activate the Sakura Editor window
    On Error Resume Next
    wsh.AppActivate "sakura"
    On Error GoTo 0

    ' Check if the window is active
    If Err.Number <> 0 Then
        MsgBox "Unable to activate Sakura Editor window.", vbExclamation
        Exit Sub
    End If

    ' Send Ctrl+V to paste the clipboard content into Sakura Editor
    wsh.SendKeys "^v"
    Application.Wait Now + TimeValue("00:00:02")

    ' Send Ctrl+A to select all content
    wsh.SendKeys "^a"
    Application.Wait Now + TimeValue("00:00:02")

    ' Send Alt+A to perform the sort
    wsh.SendKeys "%a"
    Application.Wait Now + TimeValue("00:00:02")

    ' Send Ctrl+S to save the file
    wsh.SendKeys "^s"
    Application.Wait Now + TimeValue("00:00:02")

    ' Send the file path and Enter
    wsh.SendKeys filePath
    wsh.SendKeys "{ENTER}"
    Application.Wait Now + TimeValue("00:00:02")

    ' Close the current tab (file) only if tabs are in use, otherwise close the editor
    On Error Resume Next
    wsh.AppActivate "sakura"
    ' Attempt to use Ctrl+F4 to close the tab
    wsh.SendKeys "^" & "{F4}"
    ' Check if the Sakura window is still active after sending Ctrl+F4
    If wsh.AppActivate("sakura") Then
        ' If still active, use Alt+F4 to close the editor
        wsh.SendKeys "%{F4}"
    End If
    On Error GoTo 0
End Sub


' Procedure to register directly as macro
Public Sub RunExtractAll()
    ExtractA activeSheet.name
    ExtractEx activeSheet.name
End Sub

Public Sub RunExportT()
    CopyToClipboard activeSheet.name
    ExportPVDBTxt activeSheet.name
End Sub
