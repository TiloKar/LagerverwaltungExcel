Attribute VB_Name = "EagleExport"
Public Sub EagleExport()
    Dim Dateiname As String
    Dateiname = "Z:\12 interne Elektrodokumentation\EAGLE RESSOURCEN\ulps\ScancodesEagleDB.csv"
  
    On Error GoTo fehler
    Workbooks.Open Filename:=a & b, ReadOnly:=True
    Dim fileSaveName
    fileSaveName = Application.GetSaveAsFilename( _
        fileFilter:="CSV-Datei (*.csv), *.csv", _
        InitialFileName:="Z:\12 interne Elektrodokumentation\EAGLE RESSOURCEN\ulps\ScancodesEagleDB.csv")
    If fileSaveName <> False Then
        Dateiname = fileSaveName
    End If
    
    
    Workbooks(b).Sheets(1).Activate
    Workbooks(b).Sheets(1).Columns("D:Z").Delete
    Workbooks(b).SaveAs Filename:=Dateiname, FileFormat:=xlCSV, Local:=True
    
    ActiveWorkbook.Close SaveChanges:=False
    
    Exit Sub
fehler:
    
    MsgBox "Datei wurde nicht gespeichert"

End Sub
