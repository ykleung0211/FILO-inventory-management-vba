Attribute VB_Name = "transaction"
Option Explicit
Option Base 1

Function getLastDataRow() As Long
    ' get the last row from transaction sheet with data
    getLastDataRow = CLng(historySheet.Cells(historySheet.Rows.count, idColumn).End(xlUp).row)
End Function

Function getNextId(ByVal lastDataRow As Long) As Long
    ' get the next id for transaction
    If lastDataRow < historyFirstDataRow Then
        getNextId = firstID
    Else
        getNextId = historySheet.Cells(lastDataRow, idColumn).Value + 1
    End If
End Function

Sub writeBuy(ByVal lastDataRow As Long, ByVal newId As Long, ByVal qty As Long, ByVal unitCost As Double)
    ' write the buy transaction in the transaction history worksheet
    Dim newRow As Long
    
    newRow = lastDataRow + 1
    With historySheet
    
        .Cells(newRow, idColumn) = newId
        .Cells(newRow, typeColumn) = "Buy"
        .Cells(newRow, qtyColumn) = qty
        .Cells(newRow, unitCostColumn) = Round(unitCost, 2)
        .Cells(newRow, totalAmountColumn) = Round(qty * Round(unitCost, 2), 2)
        .Cells(newRow, usageCostColumn) = "-"
        .Cells(newRow, remainingQtyColumn) = qty
    End With
End Sub

Sub writeUse(ByVal lastDataRow As Long, ByVal newId As Long, ByVal qty As Long)
    ' write the use transaction in the transaction history worksheet using FILO logic
    Dim newRow As Long
    Dim thisRow As Long
    Dim buyRow() As Long
    Dim buyRemainingQty() As Long
    Dim buyUnitCost() As Double
    Dim buyNumber As Long
    Dim idx As Long
    Dim usageCost As Double
    Dim orderQty As Long
    
    With historySheet
        ' get the number of buy transaction with remaining quantity > 0
        buyNumber = 0
        For thisRow = historyFirstDataRow To lastDataRow
            If .Cells(thisRow, typeColumn) = "Buy" And .Cells(thisRow, remainingQtyColumn).Value > 0 Then
                buyNumber = buyNumber + 1
            End If
        Next
        
        ReDim buyRow(1 To buyNumber)
        ReDim buyUnitCost(1 To buyNumber)
        ReDim buyRemainingQty(1 To buyNumber)
        
        ' get arrays of row, remaining quantity and unit cost of each buy transaction with remaining quantity > 0 staring from the last transaction
        idx = 1
        For thisRow = lastDataRow To historyFirstDataRow Step -1
            If .Cells(thisRow, typeColumn) = "Buy" And .Cells(thisRow, remainingQtyColumn).Value > 0 Then
                buyRow(idx) = thisRow
                buyRemainingQty(idx) = .Cells(thisRow, remainingQtyColumn).Value
                buyUnitCost(idx) = .Cells(thisRow, unitCostColumn).Value
                idx = idx + 1
            End If
        Next
        
        ' loop through the array starting from last transaction,
        ' take the quantity needed from each buy transaction,
        ' add to the usage cost, subtract quantity taken from remaining quantity
        ' loop until the order is fulfilled
        idx = 1
        usageCost = 0
        orderQty = qty
        Do Until buyRemainingQty(idx) >= orderQty ' stop when the remaining quantity of a transaction is enough to fulfill remaining use order quantity
            orderQty = orderQty - buyRemainingQty(idx) ' use the remaining quantity of the transaction to fulfill the order
            usageCost = usageCost + buyRemainingQty(idx) * buyUnitCost(idx) ' add the usage cost for that buy transaction
            buyRemainingQty(idx) = 0 ' used up all the transaction
            idx = idx + 1 ' use the next buy transaction
        Loop
        ' this transaction remaining quantity is enough to fulfill the remaining use order
        ' usage cost is given by the unit cost of this buy transaction times the remaining quantity that we need to fulfill
        usageCost = usageCost + orderQty * buyUnitCost(idx)
        
        ' there will have some unused quantity remaining for this buy transaction (or 0)
        buyRemainingQty(idx) = buyRemainingQty(idx) - orderQty
        
        ' update the new buy transaction after usage
        For idx = 1 To buyNumber
            .Cells(buyRow(idx), remainingQtyColumn).Value = buyRemainingQty(idx)
            If .Cells(buyRow(idx), remainingQtyColumn).Value = 0 Then ' grey out the row if the remaining quantity of that buy transaction is 0
               .Range(.Cells(buyRow(idx), idColumn), .Cells(buyRow(idx), remainingQtyColumn)).Interior.Color = grey
            Else
                .Range(.Cells(buyRow(idx), idColumn), .Cells(buyRow(idx), remainingQtyColumn)).Interior.ColorIndex = xlNone
            End If
        Next
        
        ' write the use row
        newRow = lastDataRow + 1
        .Cells(newRow, idColumn) = newId
        .Cells(newRow, typeColumn) = "Use"
        .Cells(newRow, qtyColumn) = qty
        .Cells(newRow, unitCostColumn) = "N/A"
        .Cells(newRow, totalAmountColumn) = "N/A"
        .Cells(newRow, usageCostColumn) = Round(usageCost, 2)
        .Cells(newRow, remainingQtyColumn) = "N/A"
    End With
End Sub
