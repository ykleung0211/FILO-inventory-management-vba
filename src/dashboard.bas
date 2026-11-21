Attribute VB_Name = "dashboard"
Option Explicit
Dim ws As Worksheet

Sub validateMinStock()
    ' Validate the minimum stock level and if it is invalid, show the error message cell
    Set ws = dashboardSheet
    
    If Not checkMinStock(ws.Range("minStock").Cells(1, 1).Value) Then ' Invalid input
        ws.Range("minStock").Interior.Color = lightRed
        ws.Range("minStockHint").Value = "Invalid input. Please enter a whole number from 0 to " & CStr(format(largestMinStock, "#,##0"))
    Else
        ws.Range("minStock").Interior.Color = inputYellow ' Reset the cell
        ws.Range("minStockHint").Value = "" ' Reset the cell
    End If
End Sub

Sub validateMaterialName()
     ' Validate the material name and if it is invalid, show the error message cell
     Set ws = dashboardSheet
     
     If Not checkMaterialName(ws.Range("materialName").Cells(1, 1).Value) Then ' Invalid input
        ws.Range("materialName").Interior.Color = lightRed
        ws.Range("materialNameHint").Value = "Invalid input. Please enter a nonempty name within " & CStr(format(largestMaterialName, "#,##0")) & " characters"
     Else
        ws.Range("materialName").Interior.Color = inputYellow ' Reset the cell
        ws.Range("materialNameHint").Value = "" ' Reset the cell
     End If
End Sub

Sub launchTracker()
    ' Check the value of minimum stock level and material name, if not valid then show the error cell
    ' Otherwise open the transaction user form
    On Error GoTo ErrorHandler
    
    Dim validinput As Boolean
    Set ws = dashboardSheet
    validinput = True
      
    If Not checkMinStock(ws.Range("minStock").Value) Then ' Invalid input
        validateMinStock
        ws.Range("minStock").Select
        validinput = False
    End If
    If Not checkMaterialName(ws.Range("materialName").Value) Then ' invalid input
        validateMaterialName
        ws.Range("materialName").Select
        validinput = False
    End If
    If validinput Then ' Valid input
        transactionUserForm.Show vbModeless
    End If
    Exit Sub
ErrorHandler:
    MsgBox prompt:="Unable to open the transaction form. Please try again." & vbNewLine & "Error: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Launch Error"
    Err.Clear
End Sub

Sub updateDashboard()
    ' update the dashboard based on transaction histroy
    Dim lastDataRow As Long
    Dim thisRow As Long
    Dim purchased As Long
    Dim used As Long
    Dim unitsOnHand As Long
    Dim inventoryValue As Double
    Dim usageCost As Double
    Dim minStock As Long
    Dim currentStatus As String
    
    On Error GoTo ErrorHandler
    
    Set ws = historySheet
    
    lastDataRow = getLastDataRow
    purchased = 0
    used = 0
    unitsOnHand = 0
    inventoryValue = 0#
    usageCost = 0#
    
    For thisRow = historyFirstDataRow To lastDataRow
        Select Case ws.Cells(thisRow, typeColumn).Value
            Case "Buy"
                purchased = purchased + ws.Cells(thisRow, qtyColumn).Value
                unitsOnHand = unitsOnHand + ws.Cells(thisRow, remainingQtyColumn).Value
                inventoryValue = inventoryValue + ws.Cells(thisRow, remainingQtyColumn).Value * ws.Cells(thisRow, unitCostColumn).Value
            Case "Use"
                used = used + ws.Cells(thisRow, qtyColumn).Value
                usageCost = usageCost + ws.Cells(thisRow, usageCostColumn).Value
        End Select
    Next
    
    Set ws = dashboardSheet
    
    ws.Range("totalPurchased").Value = purchased
    ws.Range("totalUsed").Value = used
    ws.Range("unitsOnHand").Value = unitsOnHand
    ws.Range("inventoryValue").Value = Round(inventoryValue, 2)
    ws.Range("usageCost").Value = Round(usageCost, 2)
    
    ' check the current stock status and formatting
    minStock = CLng(ws.Range("minStock").Value)
    currentStatus = getStockStatus(unitsOnHand, minStock)
    ws.Range("stockStatus").Value = currentStatus
    setStatusFormat currentStatus
    
    ws.Range("statusMessage").Value = "Stock Status: " & currentStatus & " with " & CStr(unitsOnHand) & " units on hand (minimum " & CStr(minStock) & " units)."
    Exit Sub
ErrorHandler:
    MsgBox prompt:="The dashboard could not be refreshed. Please try again or reset the system." & vbNewLine & "Error: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Dashboard Error"
    Err.Clear
End Sub

Function getStockStatus(ByVal units As Long, ByVal minStock As Long) As String
    Select Case units
        Case 0
            getStockStatus = "OUT OF STOCK"
        Case 1 To minStock
            getStockStatus = "LOW"
        Case (minStock + 1) To (2 * minStock)
            getStockStatus = "MEDIUM"
        Case Else
            getStockStatus = "HIGH"
    End Select
End Function

Sub setStatusFormat(ByVal status As String)
    ' update the status cell format of dashboard
    Set ws = dashboardSheet
    
    With ws.Range("stockStatus").Cells(1, 1)
        Select Case status
            Case "OUT OF STOCK"
                .Font.Color = vbWhite
                .MergeArea.Interior.Color = darkRed
            Case "LOW"
                .Font.Color = vbBlack
                .MergeArea.Interior.Color = lightRed
            Case "MEDIUM"
                .Font.Color = vbBlack
                .MergeArea.Interior.Color = yellow
            Case "HIGH"
                .Font.Color = vbBlack
                .MergeArea.Interior.Color = green
        End Select
    End With
    Application.ScreenUpdating = True ' excel have bug that merge cell can't fully change color
End Sub

Sub showStockWarning()
    ' show warning message when the stock level is drop to low or out of stock
    Dim status As String
    Dim unitsOnHand As Long
    Dim minStock As Long
    
    Set ws = dashboardSheet
    
    status = ws.Range("stockStatus").Value
    unitsOnHand = ws.Range("unitsOnHand").Value
    minStock = ws.Range("minStock").Value
    
    Select Case status
        Case "OUT OF STOCK"
            MsgBox prompt:="Warning: this material is currently OUT OF STOCK", _
                   Buttons:=vbExclamation, _
                   Title:="Stock Warning"
        Case "LOW"
            MsgBox prompt:="Warning: stock level is LOW. " & vbNewLine & CStr(unitsOnHand) & " units on hand (minimum level: " & CStr(minStock) & " units)", _
                   Buttons:=vbExclamation, _
                   Title:="Stock Warning"
    End Select
End Sub

Sub resetSystem()
    ' reset everything
    Dim confirm As Long
    
    On Error GoTo ErrorHandler
    
    confirm = MsgBox(prompt:="This will delete all transactions and restore the Dashboard to its original state. This cannot be undone. Continue?", _
                     Buttons:=vbYesNo + vbExclamation + vbDefaultButton2, _
                     Title:="Reset System")
    If confirm = vbNo Then
        Exit Sub
    End If
    
    ' clean the transaction history
    With historySheet
        With .Range(.Rows(historyFirstDataRow), .Rows(.Rows.count))
            .ClearContents
            .Interior.ColorIndex = xlNone
            .Borders.LineStyle = xlNone
        End With
    End With
    ' restore the setting
    With dashboardSheet
        .Range("materialName") = defaultMaterialName
        .Range("minStock") = defaultMinStock
        .Range("materialName").Interior.Color = inputYellow
        .Range("minStock").Interior.Color = inputYellow
        .Range("materialNameHint") = ""
        .Range("minStockHint") = ""
    
        ' restore the dashboard
        updateDashboard
        .Range("statusMessage") = "System reset. No transactions recorded yet."
    End With
    Exit Sub
ErrorHandler:
    MsgBox prompt:="The system could not be reset completely. Please try again." & vbNewLine & "Error: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Reset Error"
    Err.Clear
End Sub
