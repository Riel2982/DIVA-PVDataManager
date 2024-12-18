Attribute VB_Name = "Auth3d_ImportClear"
Sub ImportFile(sheetName As String)
    Dim filePath As String
    Dim ws As Worksheet
    Dim txtStream As Object
    Dim FileNum As Integer
    Dim fileContent As String
    Dim Lines() As String
    Dim i As Long
    Dim dict As Object ' Declare Dictionary object


    ' Open the dialog to select the bin or txt file
    filePath = Application.GetOpenFilename("Bin and Text Files (*.bin;*.txt), *.bin;*.txt", , "Select Bin or Text File")
    
    If filePath = "False" Then Exit Sub
    
    ' Check for the existence of the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet name not found: " & sheetName, vbExclamation
        Exit Sub
    End If
    
    ' Clear the contents from row 2 onwards in columns A and E
    ws.Range("A2:A1048576").ClearContents
    ws.Range("E2:E1048576").ClearContents
    
    ' Create Dictionary object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Determine the file extension
    If right(filePath, 4) = ".txt" Then
        ' Read UTF-8 encoded text file
        Set txtStream = CreateObject("ADODB.Stream")
        txtStream.Open
        txtStream.Type = 2 ' Text stream
        txtStream.Charset = "utf-8"
        txtStream.LoadFromFile filePath
        
        ' Get the text content
        fileContent = txtStream.ReadText
        txtStream.Close
        
        ' Split the content by newline character
        If InStr(fileContent, vbCrLf) > 0 Then
            Lines = Split(fileContent, vbCrLf)
        ElseIf InStr(fileContent, vbLf) > 0 Then
            Lines = Split(fileContent, vbLf)
        ElseIf InStr(fileContent, vbCr) > 0 Then
            Lines = Split(fileContent, vbCr)
        End If
    ElseIf right(filePath, 4) = ".bin" Then
        ' Open the binary file
        FileNum = FreeFile
        Open filePath For Binary As FileNum
        
        ' Read the file content
        fileContent = Input$(LOF(FileNum), FileNum)
        Close FileNum
        
        ' Split the content by newline character
        If InStr(fileContent, vbCrLf) > 0 Then
            Lines = Split(fileContent, vbCrLf)
        ElseIf InStr(fileContent, vbLf) > 0 Then
            Lines = Split(fileContent, vbLf)
        ElseIf InStr(fileContent, vbCr) > 0 Then
            Lines = Split(fileContent, vbCr)
        End If
    End If
    
    ' Insert the split content into the cells
    i = 2 ' Start from cell A2
    
' Avoid overflow by limiting rows
For Each Line In Lines
    If i > 1048576 Then ' If rows exceed 1,048,576, switch to column E
        If j > 1048576 Then Exit For ' Avoid overflow in column E
        If Not dict.exists(Line) Then ' Check for duplicate rows
            dict.Add Line, Nothing ' Add to dictionary
            ws.Cells(j, 5).Value = Line ' Insert into column E
            j = j + 1
        End If
    Else
        If Not dict.exists(Line) Then ' Check for duplicate rows
            dict.Add Line, Nothing ' Add to dictionary
            ws.Cells(i, 1).Value = Line ' Insert into column A
            i = i + 1
        End If
    End If
Next Line

    
    ' Re-enable screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Sub ClearColumns(sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Explicitly clear the contents from row 2 onwards in columns A to F
    ws.Range("A2:F1048576").ClearContents
    
End Sub


' Procedure to register directly as macro

Public Sub RunImport()
    ImportFile activeSheet.name
End Sub

Public Sub RunClear()
    ClearColumns activeSheet.name
End Sub
