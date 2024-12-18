Attribute VB_Name = "SelectMarge"
Sub SelectPV(sheetName As String)
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim selectedHeader As String
    Dim selectedColumn As Range
    Dim lastRow As Long
    Dim i As Long
    Dim copyRow As Long
    Dim pvNumbers As String
    Dim cellValue As String
    Dim additionalData As String
    Dim pvNumber As String
    Dim tempRow As Long
    Dim dict As Object
    Dim f2Data As String
    Dim i11Data As String
    Dim f2Array() As String
    Dim i11Array() As String
    Dim commonNumbers As String
    Dim j As Long
    
    ' Set worksheets
    Set ws1 = ThisWorkbook.Sheets(sheetName)
    Set ws2 = ThisWorkbook.Sheets("Temp")
    Set ws3 = ThisWorkbook.Sheets("pv_db")
    
    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Clear contents of columns F to H in ws2
    ws2.Range("F:H").ClearContents
    
    ' Get the last row in ws3
    lastRow = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row
    
    ' Initialize variables
    pvNumbers = ""
    copyRow = 1
    tempRow = 1
    
    ' Get the selected header from the dropdown list in ws1
    selectedHeader = ws1.Range("C2").value
    
    ' Temporarily disable error handling
    On Error Resume Next
    ' Find the column in ws3 that matches the selected header
    Set selectedColumn = ws3.Rows(1).Find(What:=selectedHeader, LookIn:=xlValues, LookAt:=xlWhole)
    ' Re-enable error handling
    On Error GoTo 0
    
    If Not selectedColumn Is Nothing Then
        ' Use a Dictionary object to prevent duplicates
        Set dict = CreateObject("Scripting.Dictionary")
        
        ' Process data in the matching column
        For i = 2 To lastRow
            cellValue = ws3.Cells(i, selectedColumn.Column).value
            
            If cellValue <> "" Then
                If left(cellValue, 3) = "pv_" Then
                    pvNumber = mid(cellValue, 4, InStr(4, cellValue, ".") - 4)
                    If Not dict.exists(pvNumber) Then
                        dict.Add pvNumber, pvNumber
                        ws2.Cells(tempRow, "H").value = pvNumber
                        tempRow = tempRow + 1
                    End If
                    ws2.Cells(copyRow, "G").value = cellValue
                    copyRow = copyRow + 1
                End If
            End If
        Next i
        
        ' Convert data from column H to a string with slash separators
        For i = 1 To tempRow - 1
            pvNumbers = pvNumbers & "/" & ws2.Cells(i, "H").value
        Next i
        
        ' Remove the leading slash
        If Len(pvNumbers) > 0 Then
            pvNumbers = mid(pvNumbers, 2)
        End If
        
        ' Write pv_numbers to F2 in ws2
        ws2.Range("F2").value = pvNumbers
        
        ' Get data from I11 in ws1
        additionalData = ws1.Range("I11").value
        
        ' If I11 is empty, write F2 data to F3
        If Len(additionalData) = 0 Then
            ws2.Range("F3").value = pvNumbers
        Else
            ' Split F2 and I11 data into arrays
            f2Array = Split(pvNumbers, "/")
            i11Array = Split(additionalData, "/")
            
            ' Find common numbers
            commonNumbers = ""
            For i = LBound(f2Array) To UBound(f2Array)
                For j = LBound(i11Array) To UBound(i11Array)
                    If f2Array(i) = i11Array(j) And f2Array(i) <> "" Then
                        commonNumbers = commonNumbers & "/" & f2Array(i)
                        Exit For
                    End If
                Next j
            Next i
            
            ' Remove the leading slash
            If Len(commonNumbers) > 0 Then
                commonNumbers = mid(commonNumbers, 2)
            End If
            
            ' Write common numbers to F3 in ws2
            ws2.Range("F3").value = commonNumbers
        End If
    Else
        MsgBox "Selected header not found."
        Exit Sub
    End If
    
    ' Clear contents of columns G and H in ws2
    ws2.Range("G:H").ClearContents
    
    ' Re-enable screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Sub SelectMarge(sheetName As String)
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim lastRow As Long
    Dim copyRange As Range
    Dim selectedColumn As Range
    Dim selectedHeader As String
    Dim cell As Range
    Dim i As Long
    
    ' Set the sheets
    Set ws1 = ThisWorkbook.Sheets(sheetName)
    Set ws2 = ThisWorkbook.Sheets("Temp")
    Set ws3 = ThisWorkbook.Sheets("pv_db")
    
    ' Clear column E in Temp sheet
    ws2.Columns("E:E").ClearContents

    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
        
    ' Copy data selected in dropdown list in C2 to column E
    selectedHeader = ws1.Range("C2").value
    On Error Resume Next
    Set selectedColumn = ws3.Rows(1).Find(What:=selectedHeader, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If Not selectedColumn Is Nothing Then
        ws3.Range(ws3.Cells(2, selectedColumn.Column), ws3.Cells(ws3.Rows.Count, selectedColumn.Column).End(xlUp)).Copy _
            destination:=ws2.Range("E" & ws2.Cells(ws2.Rows.Count, "E").End(xlUp).Row + 1)
    End If
    
    ' Copy data to column E if B2, B3, B4 are T
    If ws1.Range("B2").value = "T" Then
        Call ConvertAnotherSong(sheetName)
        Call AdjustTableRange("AnotherSongList", "AnotherSongList1")
        ws2.Range("A2:A" & ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row).Copy destination:=ws2.Range("E2")
    End If

    If ws1.Range("B3").value = "T" Then
        Call ConvertByModule(sheetName)
        Call AdjustTableRange("ByModuleList", "ByModuleList1")
        ws2.Range("B2:B" & ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row).Copy destination:=ws2.Range("E" & ws2.Cells(ws2.Rows.Count, "E").End(xlUp).Row + 1)
    End If

    If ws1.Range("B4").value = "T" Then
        Call ConvertExSong(sheetName)
        Call AdjustTableRange("ExSongList", "ExSongList1")
        Call AdjustTableRange("ByModuleList", "ByModuleList1")
        ws2.Range("C2:C" & ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Row).Copy destination:=ws2.Range("E" & ws2.Cells(ws2.Rows.Count, "E").End(xlUp).Row + 1)
    End If


    ' Re-enable screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Sub CopyToClipboard()
    Dim lastRow As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Temp")
    
    ' Get the last row with data in column E
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    ' Determine the range and copy to clipboard
    ws.Range("E1:E" & lastRow).Copy
End Sub

Public Sub RunExport()
    Call SelectPV(ActiveSheet.name)
    Call SelectMarge(ActiveSheet.name)
    Call CopyToClipboard
    Call ExportPVDBTxt(ActiveSheet.name)
End Sub


Public Sub ExportPVDBTxt(sheetName As String)
    Dim ws As Worksheet
    Dim folderPath As String, fileName As String, filePath As String
    Dim sakuraPath As String
    Dim wsh As Object

    ' Set the target sheet to Temp sheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Get the path to the save location from cell I4
    folderPath = ws.Range("I4").value
    If folderPath = "" Or Dir(folderPath, vbDirectory) = "" Then
        ' Show the save dialog (folder selection only) if I4 is empty or invalid
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

    ' Get the path to Sakura Editor from cell I2
    sakuraPath = ws.Range("I2").value
    If sakuraPath = "" Then
        MsgBox "Please enter the file path to Sakura Editor in cell I2.", vbExclamation
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
    Application.Wait Now + TimeValue("00:00:05")

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

