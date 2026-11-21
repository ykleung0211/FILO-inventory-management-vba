VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} transactionUserForm 
   Caption         =   "Record Transaction"
   ClientHeight    =   5950
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   7021
   OleObjectBlob   =   "transactionUserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "transactionUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub buyOption_Click()
    ' when user choose buy material, restore the error message and interior color
    On Error GoTo ErrorHandler
    With unitCostTextBox
        .Enabled = True
        .BackColor = vbWhite
    End With
    typeErrorLabel.Caption = ""
    If checkQty(qtyTextBox.Value, useOption.Value) Then ' if the error is due to not enough units on hand, then when user choose buy we should remove the error label
        If CLng(qtyTextBox.Value) > CLng(dashboardSheet.Range("unitsOnHand").Value) Then
            qtyTextBox.BackColor = vbWhite
            qtyErrorLabel.Caption = ""
        End If
    End If
    Exit Sub
ErrorHandler:
    MsgBox prompt:="Unable to check the current stock level" & vbNewLine & "Error: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Stock Check Error"
    Err.Clear
End Sub


Private Sub useOption_Click()
    ' When user choose use material, disable the unit cost box
    On Error GoTo ErrorHandler
    With unitCostTextBox
        .Enabled = False
        .Value = ""
        .BackColor = grey
    End With
    unitCostErrorLabel.Caption = ""
    If CLng(dashboardSheet.Range("unitsOnHand")) = 0 Then
        typeErrorLabel.ForeColor = &HFF&
        typeErrorLabel.Caption = "No stock on hand. Only Buy is available right now"
    Else
        typeErrorLabel.ForeColor = vbBlack
        typeErrorLabel.Caption = "At most " & CStr(dashboardSheet.Range("unitsOnHand").Value) & " units can be used"
    End If
    Exit Sub
ErrorHandler:
    MsgBox prompt:="Unable to check the current stock level" & vbNewLine & "Error: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Stock Check Error"
    Err.Clear
End Sub

Private Sub clearButton_Click()
    Call UserForm_Initialize
End Sub

Private Sub cancelButton_Click()
    Unload Me
End Sub

Private Sub refreshSettingLabel(ByRef validinput As Boolean)
    ' Reset the current setting label (material name and minimum stock level)
    ' Check the value of minimum stock level and material name, if not valid then show the error cell
    ' Setting display invalid input if fail
    Dim ws As Worksheet
    Dim materialName As String
    Dim minimumStock As String
    Dim validSetting As Boolean
    Set ws = dashboardSheet
    
    validSetting = True
    
    If Not checkMinStock(ws.Range("minStock").Value) Then ' Invalid input
        gotoDashboard
        validateMinStock
        ws.Range("minStock").Select
        validSetting = False
    End If
    If Not checkMaterialName(ws.Range("materialName").Value) Then ' invalid input
        gotoDashboard
        validateMaterialName
        ws.Range("materialName").Select
        validSetting = False
    End If
    If validSetting Then ' Valid input
        materialName = dashboardSheet.Range("materialName").Value
        minimumStock = dashboardSheet.Range("minStock").Value
        
        materialName = Trim(materialName)
        minimumStock = format(Trim(minimumStock), "#,##0")
        
        settingLabel.Caption = "Material: " & materialName & " | Minimum Stock Level: " & minimumStock
    Else ' Invalid input
        settingLabel.Caption = "Please enter valid input for material name and minimum stock level!"
    End If
    If validinput Then ' we will modify the validinput only when it is true before checking setting
        validinput = validSetting
    End If
End Sub

Private Sub submitButton_Click()
    Dim validinput As Boolean
    Dim validQty As Long
    Dim validUnitCost As Double
    Dim newId As Long
    Dim lastDataRow As Long
    On Error GoTo ErrorHandler
    
    validinput = True
    
    ' check unit cost
    If Not useOption.Value Then
        If Not checkUnitCost(unitCostTextBox.Value, useOption.Value) Then
            validinput = False
            unitCostTextBox.SetFocus
            unitCostTextBox.BackColor = lightRed
            unitCostErrorLabel.Caption = "Invalid unit cost. Please enter a numeric value within 0.01 to 1,000,000 up to 2 decimal places"
        End If
    End If
    ' check the qty
    If Not checkQty(qtyTextBox.Value, useOption.Value) Then
        qtyTextBox.SetFocus
        qtyTextBox.BackColor = lightRed
        validinput = False
        If useOption.Value Then
            qtyErrorLabel.Caption = "Invalid quantity. Please enter a whole number within 1 to 1,000,000 and smaller than or equal to current units on hand"
        Else
            qtyErrorLabel.Caption = "Invalid quantity. Please enter a whole number within 1 to 1,000,000"
        End If
    End If
        
    ' check the setting
    refreshSettingLabel validinput
    
    ' all input valid
    If validinput Then
        validQty = CLng(qtyTextBox.Text)
        lastDataRow = getLastDataRow
        newId = getNextId(lastDataRow)
        If buyOption.Value Then ' User choose buy
            validUnitCost = CDbl(unitCostTextBox.Text)
            writeBuy lastDataRow, newId, validQty, validUnitCost
            formatNewRow lastDataRow + 1
            updateDashboard
            gotoHistory
            Call UserForm_Initialize
            resultLabel.Caption = "Buy recorded: ID " & CStr(newId) & ", " & CStr(validQty) & " units at $" & CStr(format(historySheet.Cells(lastDataRow + 1, unitCostColumn).Value, "#,##0.00"))
        Else ' User choose use
            writeUse lastDataRow, newId, validQty
            formatNewRow lastDataRow + 1
            updateDashboard
            gotoHistory
            showStockWarning
            Call UserForm_Initialize
            resultLabel.Caption = "Use recorded: ID " & CStr(newId) & ", " & CStr(validQty) & " units. Usage cost: $" & CStr(format(historySheet.Cells(lastDataRow + 1, usageCostColumn).Value, "#,##0.00"))
        End If
    End If
    Exit Sub
ErrorHandler:
    MsgBox prompt:="Your transaction could not be recorded." & vbNewLine & "Error: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Transaction Error"
    Err.Clear
End Sub

Private Sub qtyTextBox_Change()
    ' Reset the background color and error label when user is editing
    qtyTextBox.BackColor = vbWhite
    qtyErrorLabel.Caption = ""
End Sub

Private Sub unitCostTextBox_Change()
    ' Reset the background color and error label when user is editing
    unitCostTextBox.BackColor = vbWhite
    unitCostErrorLabel.Caption = ""
End Sub

Private Sub UserForm_Initialize()
    Dim validinput As Boolean
    
    On Error GoTo ErrorHandler
     
    validinput = True
   
    ' Set the setting labe
    refreshSettingLabel validinput
 
    ' Set the option button for type of transaction
    buyOption.Value = True
    useOption.Value = False
    typeErrorLabel.Caption = ""
    
    ' Set the quantity text box
    qtyTextBox.Value = ""
    qtyTextBox.BackColor = vbWhite
    qtyErrorLabel.Caption = ""
    
    ' Set the unit cost text box
    unitCostTextBox.Enabled = True
    unitCostTextBox.Value = ""
    unitCostTextBox.BackColor = vbWhite
    unitCostErrorLabel.Caption = ""
    
    ' Set the result label
    resultLabel.Caption = ""
    
    qtyTextBox.SetFocus
    Exit Sub
ErrorHandler:
    MsgBox prompt:="The current settings could not be displayed." & vbNewLine & "Error: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Display Error"
    Err.Clear
End Sub
