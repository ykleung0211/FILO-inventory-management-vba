Attribute VB_Name = "formatting"
Option Explicit

Sub formatNewRow(ByVal rowNumber As Long)
    ' format a new transaction row when a transaction is submitted
    On Error GoTo ErrorHandler
    Dim thisCell As Range
    
    With historySheet
        For Each thisCell In .Range(.Cells(rowNumber, idColumn), .Cells(rowNumber, remainingQtyColumn))
            thisCell.HorizontalAlignment = xlCenter
            thisCell.Borders.LineStyle = xlContinuous
            thisCell.Borders.Color = RGB(191, 191, 191)
        Next
        
        .Cells(rowNumber, qtyColumn).NumberFormat = "#,##0"
        .Cells(rowNumber, remainingQtyColumn).NumberFormat = "#,##0"
        .Cells(rowNumber, unitCostColumn).NumberFormat = "[$$-409]#,##0.00"
        .Cells(rowNumber, totalAmountColumn).NumberFormat = "[$$-409]#,##0.00"
        .Cells(rowNumber, usageCostColumn).NumberFormat = "[$$-409]#,##0.00"
    End With
    Exit Sub
ErrorHandler:
    MsgBox prompt:="Formatting could not be applied. Your data is unaffected." & vbNewLine & "Error: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Formatting Error"
    Err.Clear
End Sub
