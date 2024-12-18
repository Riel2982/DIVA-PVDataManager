Attribute VB_Name = "HighlightDuplicate"
' Duplicate detection and sorting
Sub HighlightDuplicate()
    Dim rng As Range
    Dim cell As Range
    Dim cellValues() As String
    Dim dict As Object
    Dim cellValue As Variant
    Dim alreadyHighlighted As Object

    Set dict = CreateObject("Scripting.Dictionary")
    Set alreadyHighlighted = CreateObject("Scripting.Dictionary")
    Set rng = Union(Range("D1:D" & Cells(Rows.Count, "D").End(xlUp).Row), Range("I11"), Range("I12"))

    ' Reset font for all cells
    rng.Font.colorIndex = xlAutomatic
    rng.Font.bold = False

    ' Process non-empty cells
    For Each cell In rng
        If Trim(cell.value) <> "" Then
            ' Sort values
            cell.value = SortValues(cell.value)
            cellValues = Split(cell.value, "/")
            For Each cellValue In cellValues
                If dict.Exists(cellValue) Then
                    ' Highlight all occurrences of the duplicate value
                    HighlightAllOccurrences rng, CStr(cellValue)
                    alreadyHighlighted(cellValue) = True
                Else
                    dict.Add cellValue, cell
                End If
            Next cellValue
        End If
    Next cell

    ' Ask if user wants to reset highlight
    If dict.Count > 0 Then
        Dim shouldReset As VbMsgBoxResult
        shouldReset = MsgBox("Do you want to reset the highlight?", vbYesNo + vbQuestion, "Reset Highlight")
        If shouldReset = vbYes Then
            rng.Font.colorIndex = xlAutomatic
            rng.Font.bold = False
        End If
    End If
End Sub

Sub HighlightAllOccurrences(rng As Range, word As String)
    Dim cell As Range
    Dim startPos As Integer

    For Each cell In rng
        If Trim(cell.value) <> "" Then
            startPos = InStr(1, cell.value, word, vbTextCompare)
            Do While startPos > 0
                If IsWholeWord(cell.value, word, startPos) Then
                    cell.Characters(startPos, Len(word)).Font.Color = vbRed
                    cell.Characters(startPos, Len(word)).Font.bold = True
                End If
                startPos = InStr(startPos + Len(word), cell.value, word, vbTextCompare)
            Loop
        End If
    Next cell
End Sub

Function IsWholeWord(cellText As String, word As String, position As Integer) As Boolean
    Dim beforeChar As String
    Dim afterChar As String
    IsWholeWord = True

    If position > 1 Then
        beforeChar = mid(cellText, position - 1, 1)
        If beforeChar Like "[0-9]" Then IsWholeWord = False
    End If

    If position + Len(word) <= Len(cellText) Then
        afterChar = mid(cellText, position + Len(word), 1)
        If afterChar Like "[0-9]" Then IsWholeWord = False
    End If
End Function

Function SortValues(cellValue As String) As String
    Dim values() As String
    Dim temp As String
    Dim i As Integer, j As Integer

    values = Split(cellValue, "/")
    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If Val(values(i)) > Val(values(j)) Then
                temp = values(i)
                values(i) = values(j)
                values(j) = temp
            End If
        Next j
    Next i

    SortValues = Join(values, "/")
End Function

