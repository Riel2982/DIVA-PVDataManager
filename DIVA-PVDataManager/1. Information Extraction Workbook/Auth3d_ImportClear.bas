Attribute VB_Name = "Auth3d_ImportClear"
Sub ImportBinFile(sheetName As String)
    Dim FilePath As String
    Dim FileNum As Integer
    Dim FileContent As String
    Dim Lines() As String
    Dim i As Integer
    Dim ws As Worksheet
    
    ' Open the dialog to select the text file
    FilePath = Application.GetOpenFilename("Bin Files (*.bin), *.bin", , "Select Bin File")
    
    If FilePath = "False" Then Exit Sub
    
    ' Check for the existence of the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet name not found: " & sheetName, vbExclamation
        Exit Sub
    End If
    
    ' Open the text file
    FileNum = FreeFile
    Open FilePath For Binary As FileNum
    
    ' Read the file content
    FileContent = Input$(LOF(FileNum), FileNum)
    Close FileNum
    
    ' Split the content by newline character
    If InStr(FileContent, vbCrLf) > 0 Then
        Lines = Split(FileContent, vbCrLf)
    ElseIf InStr(FileContent, vbLf) > 0 Then
        Lines = Split(FileContent, vbLf)
    ElseIf InStr(FileContent, vbCr) > 0 Then
        Lines = Split(FileContent, vbCr)
    End If
    
    ' Insert the split content into the cells
    i = 2 ' ' Start from cell A2
    For Each Line In Lines
        ws.Cells(i, 1).Value = Line
        i = i + 1
    Next Line
End Sub

Sub ClearColumns(sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Explicitly clear the contents from row 2 onwards in columns A to F
    ws.Range("A2:F1048576").ClearContents
    
End Sub


' Procedure to register directly as macro

Public Sub RunImport()
    ImportBinFile ActiveSheet.Name
End Sub

Public Sub RunClear()
    ClearColumns ActiveSheet.Name
End Sub
