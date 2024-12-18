Attribute VB_Name = "SelectToggle"
Sub Toggle_A()
    Dim ws1 As Worksheet
    
    Set ws1 = ThisWorkbook.Sheets("★Control Center")

    If ws1.Range("B2").value = "F" Then
        ws1.Range("B2").value = "T"
    ElseIf ws1.Range("B2").value = "T" Then
        ws1.Range("B2").value = "F"
    Else
        ws1.Range("B2").value = "F" ' Set the initial value
    End If
End Sub

Sub Toggle_B()
    Dim ws1 As Worksheet
    
    Set ws1 = ThisWorkbook.Sheets("★Control Center")

    If ws1.Range("B3").value = "F" Then
        ws1.Range("B3").value = "T"
    ElseIf ws1.Range("B3").value = "T" Then
        ws1.Range("B3").value = "F"
    Else
        ws1.Range("B3").value = "F" ' Set the initial value
    End If
End Sub

Sub Toggle_C()
   Dim ws1 As Worksheet
        
   Set ws1 = ThisWorkbook.Sheets("★Control Center")
    
   If ws1.Range("B4").value = "F" Then
       ws1.Range("B4").value = "T"
   ElseIf ws1.Range("B4").value = "T" Then
       ws1.Range("B4").value = "F"
   Else
       ws1.Range("B4").value = "F" ' Set the initial value
   End If
End Sub

