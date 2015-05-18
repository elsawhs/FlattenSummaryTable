Sub Flatten()
' Flatten Macro
'
'  Make sure you have a summary table with Header
'  Flat table will produce 4 columns and delete the rows with empty value on "Sale"
Dim SummaryTable As Range, FlatTable As Range
Dim flatR As Long
Dim tempR As Long, tempC As Long
On Error Resume Next
Set SummaryTable = ActiveCell.CurrentRegion
If SummaryTable.Count = 1 Or SummaryTable.Rows.Count < 4 Then
MsgBox "Select a cell within the summary table.", vbCritical
Exit Sub
End If
SummaryTable.Select
Set FlatTable = Application.InputBox(prompt:="Select a cell for the 4-column output", Type:=8)
'   Convert the range
OutRow = 2
Application.ScreenUpdating = False
FlatTable.Range("A1:D1") = Array("Product", "Sale Rep", "Year", "Sale")
For tempR = 2 To SummaryTable.Rows.Count
For tempC = 3 To SummaryTable.Columns.Count
FlatTable.Cells(flatR, 1) = SummaryTable.Cells(tempR, 1)
FlatTable.Cells(flatR, 2) = SummaryTable.Cells(tempR, 2)
FlatTable.Cells(flatR, 3) = SummaryTable.Cells(tempR, tempC)
FlatTable.Cells(flatR, 4) = SummaryTable.Cells(tempR, tempC)
FlatTable.Cells(flatR, 4).NumberFormat = SummaryTable.Cells(tempR, tempC).NumberFormat
flatR = flatR + 1
Next tempC
Next tempR
FlatTable.Columns(4).SpecialCells(xlBlanks).EntireRow.Delete
'
End Sub