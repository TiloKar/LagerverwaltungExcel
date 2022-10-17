Attribute VB_Name = "ButtonsReports"
Sub ReportProblemkinder()

    'Call Inits
    Select Case fixeduser
        Case "TK", "BW", "KM", "KW", "SD", "SBL", "MK", "MF"
            'nix
        Case Else
            MsgBox "Für Nutzer " & fixeduser & " nicht zulässig"
            Exit Sub
    End Select
    
    On Error GoTo FehlerOpenProblemkinder
    Workbooks.Open Filename:=a & d, ReadOnly:=False
    On Error GoTo 0
    Dim i As Integer
    Dim terminal As Worksheet        'temporäre arbeitsblatt
    Set terminal = Workbooks(Dateiname).Worksheets(1)
    Dim kinder As Worksheet        'temporäre arbeitsblatt
    Set kinder = Workbooks(d).Worksheets(1)
    
    Dim letzteZeile As Integer
    letzteZeile = kinder.UsedRange.Cells(kinder.UsedRange.Rows.Count, 1).Row 'kinder.UsedRange.Rows.Count
    
    If letzteZeile >= 2 Then
        For i = 2 To letzteZeile
            terminal.Rows(2).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
            kinder.Range(kinder.Cells(i, 1), kinder.Cells(i, 11)).Cut Destination:=terminal.Cells(2, 1)
        Next i
    Else
        MsgBox "Toll- keine Problemkinder vorhanden"
    End If
  
  
    Workbooks(d).Close SaveChanges:=True
    If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
      '  Beep 2000, 750
    End If
  
    Workbooks(Dateiname).Save
    Workbooks(Dateiname).Worksheets(1).Activate
    
Exit Sub
FehlerOpenProblemkinder:
    MsgBox "Fehler beim Öffnen von Problemkinder"
Exit Sub

End Sub
Sub ÖffneReportProjektmappe()
   ' Call Inits
   ' Call holeDatenbank
    
    
'
'    ReportProjektmappe.Projektauswahl.Clear
'    ReportProjektmappe.Projektauswahl.AddItem "Alle"
'
'
'     If plcount > 1 Then
'        ReportProjektmappe.Projektauswahl.ListIndex = 0
'        For i = 0 To plcount - 1
'            ReportProjektmappe.Projektauswahl.AddItem plnames(i) 'projektauswahlliste füllen
'        Next i
'    Else
'        MsgBox "keine Projektblätter gefunden"
'    End If
'    ReportProjektmappe.Projektauswahl.ListIndex = 0
'
'
'    ReportProjektmappe.Show
    
    
End Sub

Public Sub KopiereProjektmappe()          'kann theoretisch so gelassen werden, es sei denn Projekte sollen auch in die Listbocx. Praktisch?

    'to do umwandeln in kopiere Datenbank
    Dim i As Integer
    Dim k As Integer
    On Error Resume Next
    Workbooks.Open Filename:=a & b, ReadOnly:=True, Password:=pwlager
    fName = Application.GetSaveAsFilename( _
        fileFilter:="Excel-Datei (*.xlsx), *.xlsx", _
        InitialFileName:="Kopie_Lagerliste")
    
    Workbooks(b).SaveAs (fName)
    Workbooks(b).Close SaveChanges:=False

    
End Sub


Public Sub ReportJournal()
    'Call Inits
    Workbooks.Open Filename:=a & c, ReadOnly:=True, WriteResPassword:=pwjournal, Password:=pwjournal
    Workbooks.Add
    Set NeueMappe = ActiveWorkbook
    Workbooks(c).Worksheets(1).Copy Before:=NeueMappe.Sheets(1)
    NeueMappe.Sheets(1).Cells(1, 9).Value = "Buchungsjournal"
    NeueMappe.Sheets(1).Cells(2, 8).Value = "Stand:"
    NeueMappe.Sheets(1).Cells(2, 9).Value = Format(Now(), "DD.MM.YYYY   hh:mm:ss")
    Workbooks(c).Close SaveChanges:=False
    If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
      '  Beep 2000, 750
    End If
End Sub



Public Sub ReportRoterPunkt()
  '  Call Inits
    Dim k As Integer
    Dim reportzeile As Integer
    reportzeile = 2
    'On Error GoTo FehlerOpen
    Workbooks.Open Filename:=a & b, ReadOnly:=True, Password:=pwlager
    'On Error GoTo FehlerNachOpen
    If ActiveWorkbook.Sheets.Count > 1 Then
        Dim lagerliste As Worksheet
        Dim NeueMappe As Worksheet
        Dim treffer As Object
        
        Set lagerliste = Workbooks(b).Worksheets(1)
        
        Set treffer = lagerliste.Columns(1).Find(what:="", lookat:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False)
        If Not (treffer Is Nothing) Then
            If treffer.Row > 2 Then
                'MsgBox (treffer.Row)
                For k = 2 To (treffer.Row - 1)
                    If CStr(lagerliste.Cells(k, 9).Value) = "Nachbestellen" Then
                        If reportzeile = 2 Then 'bei erstem treffer neues Arbeitsblatt anlegen kopfzeile anlegen
                           Workbooks.Add (a + "\Report Roter Punkt.xltx")
                           Set NeueMappe = ActiveWorkbook.Sheets(1)
                        End If
                        lagerliste.Range(lagerliste.Cells(k, 1), lagerliste.Cells(k, 6)).Copy Destination:=NeueMappe.Cells(reportzeile, 1)
                        NeueMappe.Cells(reportzeile, 7).Value = lagerliste.Cells(k, 12).Value  'letzte bedarfsmeldung
                        NeueMappe.Cells(reportzeile, 8).Value = lagerliste.Cells(k, 13).Value  'zu wann
                        NeueMappe.Cells(reportzeile, 9).Value = lagerliste.Cells(k, 14).Value  'wer
                        reportzeile = reportzeile + 1
                    End If
                Next k
            End If
        End If

        If Workbooks(Dateiname).Worksheets(3).Cells(10, 4).Value <> "" Then
          '  Beep 2000, 750
        End If
       
    End If
    Workbooks(b).Close SaveChanges:=False
    
Exit Sub
    
FehlerOpen:
    MsgBox "Fehler in Sub ReportRoterPunkt_Click beim Öffnen" & vbCrLf & Err.Number & vbCrLf & Err.Description

Exit Sub

FehlerNachOpen:
    MsgBox "Fehler in Sub ReportRoterPunkt_Click nach Öffnen" & vbCrLf & Err.Number & vbCrLf & Err.Description
    On Error Resume Next
    Workbooks(b).Close SaveChanges:=False
End Sub
 
