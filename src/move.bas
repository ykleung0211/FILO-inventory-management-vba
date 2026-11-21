Attribute VB_Name = "move"
Option Explicit

Sub showOnly(ByVal Target As Worksheet)
    ' Set only the target worksheet visible while other worksheets invisible
    On Error GoTo ErrorHandler
    Target.Visible = True
    Target.Activate
    Target.Range("a1").Select
    
    If coverSheet.Name <> Target.Name Then _
        coverSheet.Visible = False
    If dashboardSheet.Name <> Target.Name Then _
        dashboardSheet.Visible = False
    If historySheet.Name <> Target.Name Then _
        historySheet.Visible = False
    Exit Sub
ErrorHandler:
    MsgBox prompt:="Unable to switch pages" & vbNewLine & "Error: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Error Occurred"
    Err.Clear
End Sub

Sub gotoDashboard()
    showOnly dashboardSheet
End Sub

Sub gotoCover()
    showOnly coverSheet
End Sub

Sub gotoHistory()
    showOnly historySheet
End Sub

