Attribute VB_Name = "Module1"
Sub ticker():
For Each ws In Worksheets



RowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

For i = 2 To RowCount
    
    ws.Cells(i, 9).Value = ws.Cells(i, 1).Value
    ws.Cells(i, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
    ws.Cells(i, 11).Value = (ws.Cells(i, 6).Value / ws.Cells(i, 3).Value) - 1
    ws.Cells(i, 12).Value = ws.Cells(i, 7).Value * ws.Cells(i, 6).Value
    
    
    
    
Next i



    ws.Range("P2").Value = WorksheetFunction.Max(ws.Range("K2:K" & RowCount))
    ws.Range("P3").Value = WorksheetFunction.Min(ws.Range("K2:K" & RowCount))
    ws.Range("P4").Value = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    ws.Range("O2").Value = ws.Cells(increase_number + 1, 9)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    ws.Range("O3").Value = ws.Cells(decrease_number + 1, 9)
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
    ws.Range("O4").Value = ws.Cells(increase_number + 1, 9)

    
    
Next ws


End Sub
