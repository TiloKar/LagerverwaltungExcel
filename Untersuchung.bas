Attribute VB_Name = "Untersuchung"
Sub untersuchung()

    Workbooks.Open Filename:=a & b, ReadOnly:=True
 
    Dim i As Integer
    Dim ws As Worksheet        'temporäre arbeitsblatt
    Dim neu As Workbook

    Set neu = Workbooks.Add
    For i = Workbooks(b).Worksheets.Count To 1 Step -1
        
        Set ws = Workbooks(b).Worksheets(i)
        neu.Sheets(1).Cells(i, 1).Value = ws.Name
        neu.Sheets(1).Cells(i, 2).Value = ws.UsedRange.Rows.Count
        neu.Sheets(1).Cells(i, 3).Value = ws.UsedRange.Columns.Count
    Next i
    
    Workbooks(b).Close
  
End Sub
