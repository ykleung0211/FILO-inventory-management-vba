Attribute VB_Name = "validation"
Option Explicit

Function checkMinStock(ByVal minStock As String) As Boolean
    ' Minimum stock level has to be nonempty, numeric,
    ' a whole number and greater than or equal to 0
    minStock = Trim(minStock)
    If Len(minStock) <> 0 And IsNumeric(minStock) Then
        minStock = CDbl(minStock)
        If minStock >= 0 And minStock <= largestMinStock And (minStock = Int(minStock)) Then
            checkMinStock = True
        Else
            checkMinStock = False
        End If
    Else
        checkMinStock = False
    End If
End Function

Function checkMaterialName(ByVal materialName As String) As Boolean
    ' Material name has to be nonempty and less than or equal to 50 characters
    materialName = Trim(materialName)
    If Len(materialName) = 0 Or Len(materialName) > largestMaterialName Then
        checkMaterialName = False
    Else
        checkMaterialName = True
    End If
End Function

Function checkQty(ByVal qty As String, ByVal isuse As Boolean) As Boolean
    ' Qty has to be nonempty, numeric, positive and whole number smaller or equal to largest quantity
    ' If type is use, it can not be more than the units on hand
    Dim dblQty As Double
    
    qty = Trim(qty)
    If Len(qty) <> 0 And IsNumeric(qty) Then
        dblQty = CDbl(qty)
        If dblQty = Int(dblQty) And dblQty > 0 And dblQty <= largestQty Then
            If isuse Then
                If dblQty <= CLng(dashboardSheet.Range("unitsOnHand")) Then
                    checkQty = True
                Else
                    checkQty = False
                End If
            Else
                checkQty = True
            End If
        Else
            checkQty = False
        End If
    Else
        checkQty = False
    End If
End Function

Function checkUnitCost(ByVal unitCost As String, ByVal isuse As Boolean) As Boolean
    ' Unit cost has to be nonempty, numeric, greater than 0 and smaller or equal to largest unit cost
    ' If type is use, no unit cost should be recorded
    Dim dblUnitCost As Double
    
    unitCost = Trim(unitCost)
    If Len(unitCost) <> 0 And IsNumeric(unitCost) Then
        dblUnitCost = CDbl(unitCost)
        If Round(dblUnitCost, 2) >= 0.01 And dblUnitCost <= largestUnitCost Then
            If isuse = False Then
                checkUnitCost = True
            Else
                checkUnitCost = False
            End If
        Else
            checkUnitCost = False
        End If
    Else
        checkUnitCost = False
    End If
End Function
