Attribute VB_Name = "AdjustTable"
' Subroutine to automatically adjust table range
Public Sub AdjustTableRange(sheetName As String, tableName As String)
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim lastCol As Long
    Dim ws As Worksheet

    ' Disable screen updating and automatic calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Set the specified table name
    Set tbl = ws.ListObjects(tableName)

    ' Get the last row and column
    lastRow = tbl.Range.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = tbl.Range.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' Adjust the table range
    tbl.Resize tbl.Range.Resize(lastRow - tbl.HeaderRowRange.Row + 1, lastCol - tbl.Range.Column + 1)

    ' Re-enable screen updating and automatic calculation
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

