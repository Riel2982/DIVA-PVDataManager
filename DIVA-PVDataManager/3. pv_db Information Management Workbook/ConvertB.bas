Attribute VB_Name = "ConvertB"
Sub ConvertByModule(sheetName As String)
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim i As Long, j As Long
    Dim moduleIDs() As String
    Dim lastRow As Long
    Dim counter As Long
    Dim authCounter As Long
    Dim currentPV As String
    Dim previousPV As String
    
    Set ws1 = ThisWorkbook.Sheets("ByModuleList")
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

    ' Clear column B in the Temp sheet
    ws2.Columns("B").ClearContents
    
    lastRow = ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
    
    counter = 1
    previousPV = ""
    authCounter = 0
    
    For i = 2 To lastRow
        currentPV = "pv_" & Format(ws1.Cells(i, 2).value, "000")

        ' Check the slot group conditions
        If IsGroupIncluded(currentPV, includeGroups) And Not IsGroupIncluded(currentPV, excludeGroups) Then
            If currentPV <> previousPV Then
                If authCounter > 0 Then
                    ws2.Cells(counter, 2).value = previousPV & ".auth_replace_by_module.length=" & authCounter
                    counter = counter + 1
                End If
                authCounter = 0
            End If
            
            moduleIDs = Split(ws1.Cells(i, 7).value, "/")
            For j = LBound(moduleIDs) To UBound(moduleIDs)
                ws2.Cells(counter, 2).value = currentPV & ".auth_replace_by_module." & authCounter & ".id=" & ws1.Cells(i, 3).value - 1
                ws2.Cells(counter + 1, 2).value = currentPV & ".auth_replace_by_module." & authCounter & ".module_id=" & moduleIDs(j)
                
                ' Modify the name processing based on the condition
                If ws1.Cells(i, 6).value = "P" Then
                    ws2.Cells(counter + 2, 2).value = currentPV & ".auth_replace_by_module." & authCounter & ".name=" & ws1.Cells(i, 4).value & "_" & ws1.Cells(i, 5).value
                ElseIf ws1.Cells(i, 6).value = "F" Then
                    ws2.Cells(counter + 2, 2).value = currentPV & ".auth_replace_by_module." & authCounter & ".name=" & ws1.Cells(i, 5).value
                End If
                ws2.Cells(counter + 3, 2).value = currentPV & ".auth_replace_by_module." & authCounter & ".org_name=" & ws1.Cells(i, 4).value
                
                authCounter = authCounter + 1
                counter = counter + 4
            Next j
            
            previousPV = currentPV
        End If
    Next i
    
    ' Length processing for the last PV slot group
    If authCounter > 0 Then
        ws2.Cells(counter, 2).value = previousPV & ".auth_replace_by_module.length=" & authCounter
    End If
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

