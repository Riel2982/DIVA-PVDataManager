VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Import_Click()
    RunImport
End Sub

Private Sub Clear_Click()
    RunClear
End Sub

Private Sub Opt_Click()
    RunOptimize
End Sub

Private Sub Export_Click()
    RunExport
End Sub
