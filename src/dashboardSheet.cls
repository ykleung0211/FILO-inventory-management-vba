VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dashboardSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim ws As Worksheet

Private Sub Worksheet_Activate()
    updateDashboard
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Check whether the user is editing the yellow input cell (minimum stock level and material name)
    ' Validate the input and if the input is invalid, show the error message cell
    On Error GoTo ErrorHandler
    
    Set ws = dashboardSheet
    
    If Not Intersect(Target, ws.Range("minStock")) Is Nothing Then ' User is editing minimum stock cell
        validateMinStock
        If checkMinStock(ws.Range("minStock")) Then
            updateDashboard
        End If
    End If
    If Not Intersect(Target, ws.Range("materialName")) Is Nothing Then ' User is editing material name cell
        validateMaterialName
    End If
    Exit Sub
ErrorHandler:
    MsgBox prompt:="The input could not be checked. Please re-enter the value." & vbNewLine & "Error: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Input Check Error"
    Err.Clear
End Sub

