VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' restore the formula bar seems the setting is excel-wise
    Application.DisplayFormulaBar = True
End Sub

Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    
    ' show the cover page only
    Application.DisplayFormulaBar = False
    gotoCover
     Exit Sub
ErrorHandler:
    MsgBox prompt:="This workbook could not start correctly. Please close and reopen the file." & vbNewLine & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Open Error"
    Err.Clear
End Sub

