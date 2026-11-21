Attribute VB_Name = "report"
Option Explicit
Sub countBuyBatches(ByRef fullyUsed As Long, ByRef available As Long)
    ' count the number of fully used and available buy batches
    ' change the two values directly using byref
    Dim lastDataRow As Long
    Dim thisRow As Long
    
    fullyUsed = 0
    available = 0
    
    lastDataRow = getLastDataRow
    
    With historySheet
        For thisRow = historyFirstDataRow To lastDataRow
            If .Cells(thisRow, typeColumn) = "Buy" Then
                If .Cells(thisRow, remainingQtyColumn) = 0 Then
                    fullyUsed = fullyUsed + 1
                Else
                    available = available + 1
                End If
            End If
        Next
    End With
End Sub
Sub generateReport()
    ' generate the full report
    Dim reportBook As Workbook
    Dim reportSheet As Worksheet
    Dim fullyUsed As Long
    Dim available As Long
    Dim reportTime As Date
    Dim suggestedName As String
    Dim isSaved As Boolean
    Dim confirm As Long
    
    On Error GoTo ErrorHandler
    
    If Not checkMinStock(dashboardSheet.Range("minStock")) Or Not checkMaterialName(dashboardSheet.Range("materialName")) Then
        MsgBox prompt:="Please correct the red setting cells first", _
               Buttons:=vbInformation, _
               Title:="Invalid Setting"
        Exit Sub
    End If
    
    updateDashboard
    
    fullyUsed = 0
    available = 0
    
    countBuyBatches fullyUsed, available
    
    ' copy the template report
    reportTime = Now
    reportTemplateSheet.Visible = xlSheetVisible
    reportTemplateSheet.Copy ' this create a new workbook with the sheet copied and being active
    
    Set reportBook = ActiveWorkbook
    Set reportSheet = reportBook.Worksheets(1)
    
    reportTemplateSheet.Visible = xlSheetVeryHidden
    
    suggestedName = "Inventory_Report_" & format(reportTime, "yyyymmdd_hhnnss")
    reportSheet.Name = "Inventory Report"
     
    ' since template is used, it is better not to used named range for confusing
    With reportSheet
        .Range("a1").Select
        ' "F8": material name and minimum stock level "K13": current time
        ' "K14": Total units purchased "K15": total units used
        ' "K16": current units on hand "K17": total inventory value
        ' "K18": total usage cost "K19": fully used buy batch
        ' "K20": availble buy batch "K21": current stock status
        
        .Range("F8") = "Material: " & Trim(dashboardSheet.Range("materialName")) & " | Minimum Stock Level: " & dashboardSheet.Range("minStock")
        .Range("k13") = reportTime
        .Range("k14") = dashboardSheet.Range("totalPurchased")
        .Range("k15") = dashboardSheet.Range("totalUsed")
        .Range("k16") = dashboardSheet.Range("unitsOnHand")
        .Range("k17") = dashboardSheet.Range("inventoryValue")
        .Range("k18") = dashboardSheet.Range("usageCost")
        .Range("k19") = fullyUsed
        .Range("k20") = available
        .Range("k21") = dashboardSheet.Range("stockStatus")
    End With
    
    ' note that in read mode we can't change the name of the workbook
    ' user can change it when they want to save it
    confirm = MsgBox(prompt:="Do you want to save the report as a file?", _
                     Buttons:=vbYesNo + vbQuestion + vbDefaultButton1, _
                     Title:="Save Report")
    If confirm = vbYes Then
        isSaved = Application.Dialogs(xlDialogSaveAs).Show(suggestedName & ".xlsx")
    End If
    
    If isSaved = False Then
        reportBook.Windows(1).Caption = suggestedName
    End If
    ' user can close the workbook without message
    reportBook.Saved = True
    Exit Sub
ErrorHandler:
    On Error Resume Next
    reportTemplateSheet.Visible = xlSheetVeryHidden
    If Not reportBook Is Nothing Then
        reportBook.Close savechanges:=False
    End If
    MsgBox prompt:="The report could not be created. Please try again." & vbNewLine & "Error: " & Err.Description, _
           Buttons:=vbCritical, _
           Title:="Report Error"
    Err.Clear
End Sub
